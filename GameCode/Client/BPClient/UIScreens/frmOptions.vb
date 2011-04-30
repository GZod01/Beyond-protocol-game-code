Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmOptions
    Inherits UIWindow

    Private Enum eShowFrameType As Byte
        eGraphicsFrame = 0
        eAudioFrame = 1
        eControlsFrame = 2
    End Enum

    Private WithEvents fraCurrent As UIWindow   'contains our current frame being viewed...

    Private WithEvents btnGraphics As UIButton
    Private WithEvents btnAudio As UIButton
    Private WithEvents btnControls As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnQuitGame As UIButton
	Private WithEvents btnAliases As UIButton
    Private WithEvents btnSelfDestruct As UIButton
    Private WithEvents btnRestartTutorial As UIButton
    Private WithEvents btnClaim As UIButton
    Private WithEvents btnCredits As UIButton
    Private WithEvents chkRaiseFullInvul As UICheckBox

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmOptions initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eOptions
            .ControlName = "frmOptions"
            .Left = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 185
            .Top = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 200
            .Width = 660 '370
            .Height = 405
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True  'I want it to show up regardless
            .mbAcceptReprocessEvents = True
        End With

        'btnGraphics initial props
        btnGraphics = New UIButton(oUILib)
        With btnGraphics
            .ControlName = "btnGraphics"
            .Left = 20
            .Top = 20
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Graphics"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnGraphics, UIControl))

        'btnAudio initial props
        btnAudio = New UIButton(oUILib)
        With btnAudio
            .ControlName = "btnAudio"
            .Left = 135
            .Top = 18
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Audio"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAudio, UIControl))

        'btnControls initial props
        btnControls = New UIButton(oUILib)
        With btnControls
            .ControlName = "btnControls"
            .Left = 250
            .Top = 18
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Controls"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnControls, UIControl))

        'btnClaim initial props
        btnClaim = New UIButton(oUILib)
        With btnClaim
            .ControlName = "btnClaim"
            .Left = Me.Width - 130
            .Top = 18
            .Width = 110
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Claim Rewards"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to view the Rewards dialog for any potential rewards that you may be able to claim."
        End With
        Me.AddChild(CType(btnClaim, UIControl))

        'btnCredits initial props
        btnCredits = New UIButton(oUILib)
        With btnCredits
            .ControlName = "btnCredits"
            .Left = btnClaim.Left - 105
            .Top = 18
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Credits"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = ""
        End With
        Me.AddChild(CType(btnCredits, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 110 ' 260
            .Top = 375
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Close"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnQuitGame initial props
        btnQuitGame = New UIButton(oUILib)
        With btnQuitGame
            .ControlName = "btnQuitGame"
            .Left = 10
            .Top = 375
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Quit Game"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnQuitGame, UIControl))

        'btnAliases initial props
        btnAliases = New UIButton(oUILib)
        With btnAliases
            .ControlName = "btnAliases"
            .Left = btnQuitGame.Left + btnQuitGame.Width + 5
            .Top = btnQuitGame.Top
            .Width = 100
            .Height = 24
            .Enabled = (goCurrentPlayer Is Nothing = True OrElse (goCurrentPlayer.yPlayerPhase = 255 AndAlso goCurrentPlayer.lAccountStatus = AccountStatusType.eActiveAccount))
            .Visible = True
            .Caption = "Aliases"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
		Me.AddChild(CType(btnAliases, UIControl))

		'btnSelfDestruct initial props
		btnSelfDestruct = New UIButton(oUILib)
		With btnSelfDestruct
			.ControlName = "btnSelfDestruct"
			.Left = btnAliases.Left + btnAliases.Width + 5
			.Top = btnAliases.Top
            .Width = 150 '110
			.Height = 24
            .Enabled = Not gbAliased '(goCurrentPlayer Is Nothing = True) OrElse (goCurrentPlayer.yPlayerPhase = 255)
			.Visible = True
            .Caption = "Self Destruct"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Kills your current player account and restarts you initiating the contingency plan." & vbCrLf & "This is the same as killing off all colonies so that the contingency plan becomes available." & vbCrLf & "However, by doing so, you will not have a contingency plan."
		End With
		Me.AddChild(CType(btnSelfDestruct, UIControl))

        'btnRestartTutorial initial props
        btnRestartTutorial = New UIButton(oUILib)
        With btnRestartTutorial
            .ControlName = "btnRestartTutorial"
            .Left = btnSelfDestruct.Left + btnSelfDestruct.Width + 5
            .Top = btnAliases.Top
            .Width = 130
            .Height = 24
            .Enabled = True '(goCurrentPlayer Is Nothing = True) OrElse (goCurrentPlayer.yPlayerPhase = 255)
            .Visible = False
            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255 Then
                .Visible = True
            End If
            .Caption = "Restart Tutorial"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Kills your current progress and places you back at step 1 of the tutorial."
        End With
        Me.AddChild(CType(btnRestartTutorial, UIControl))

        'chkRaiseFullInvul initial props
        chkRaiseFullInvul = New UICheckBox(oUILib)
        With chkRaiseFullInvul
            .ControlName = "chkRaiseFullInvul"
            .Left = Me.Width - 210
            .Top = btnCredits.Top + btnCredits.Height + 3
            .Width = 190
            .Height = 18
            .Enabled = Not gbAliased
            .Visible = False

            .Caption = "Raise Full Invulnerability"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck


            If goCurrentPlayer Is Nothing = False Then
                .Visible = True
                .Value = (goCurrentPlayer.lStatusFlags And elPlayerStatusFlag.AlwaysRaiseFullInvul) <> 0
            End If

            .ToolTipText = ""
        End With
        Me.AddChild(CType(chkRaiseFullInvul, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        ShowFrame(eShowFrameType.eGraphicsFrame)
    End Sub

    Private Sub ShowFrame(ByVal yType As eShowFrameType)

        If fraCurrent Is Nothing = False Then
            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X).ControlName = fraCurrent.ControlName Then
                    Me.RemoveChild(X)
                    Exit For
                End If
            Next X
        End If
        fraCurrent = Nothing

        Select Case yType
            Case eShowFrameType.eAudioFrame
                fraCurrent = New fraAudio(MyBase.moUILib)
            Case eShowFrameType.eControlsFrame
                fraCurrent = New fraControls(MyBase.moUILib)
            Case eShowFrameType.eGraphicsFrame
				fraCurrent = New fraGraphics2(MyBase.moUILib)
        End Select

        If fraCurrent Is Nothing = False Then
            fraCurrent.Left = 10
            fraCurrent.Top = 64
            fraCurrent.mbAcceptReprocessEvents = True
            Me.AddChild(CType(fraCurrent, UIControl))
        End If

    End Sub

	Private Sub btnAudio_Click(ByVal sName As String) Handles btnAudio.Click
		ShowFrame(eShowFrameType.eAudioFrame)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmAliasing")
        If ofrm Is Nothing = True Then
            MyBase.moUILib.RemoveWindow("frmAliasLogins")
        End If
        ofrm = Nothing
        'force a save on the window settings
        muSettings.SaveSettings()
    End Sub

	Private Sub btnControls_Click(ByVal sName As String) Handles btnControls.Click
		ShowFrame(eShowFrameType.eControlsFrame)
	End Sub

	Private Sub btnGraphics_Click(ByVal sName As String) Handles btnGraphics.Click
		ShowFrame(eShowFrameType.eGraphicsFrame)
	End Sub

	Private Sub btnQuitGame_Click(ByVal sName As String) Handles btnQuitGame.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		frmMain.Close()
	End Sub

#Region " Frame Declarations "
	'Interface created from Interface Builder
	Private Class fraAudio
		Inherits UIWindow

		Private WithEvents chkAudio As UICheckBox
        Private WithEvents chkMusicOn As UICheckBox
        Private WithEvents chkPostional As UICheckBox
		Private WithEvents lblVolume As UILabel
		Private WithEvents lblMaster As UILabel
		Private WithEvents hscrMaster As UIScrollBar
		Private WithEvents lblMusicVol As UILabel
		Private WithEvents hscrVG0 As UIScrollBar
		Private WithEvents lblWpn As UILabel
		Private WithEvents lblEntityA As UILabel
		Private WithEvents lblEnvirA As UILabel
		Private WithEvents lblUnitSpeech As UILabel
		Private WithEvents lblUI As UILabel
		Private WithEvents lblGameNarr As UILabel
        Private WithEvents lblTutorialVoice As UILabel
        Private WithEvents hscrVG1 As UIScrollBar
		Private WithEvents hscrVG2 As UIScrollBar
		Private WithEvents hscrVG3 As UIScrollBar
		Private WithEvents hscrVG4 As UIScrollBar
		Private WithEvents hscrVG5 As UIScrollBar
		Private WithEvents hscrVG6 As UIScrollBar
        Private WithEvents hscrVG7 As UIScrollBar

		Private mbLoading As Boolean = True

		Private mlLastVals() As Int32
		Private msTestWAV() As String
		Private mlTestIdx() As Int32
		Private mlTestUsage() As Int32

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraAudio initial props
			With Me
				.ControlName = "fraAudio"
				.Left = 199
				.Top = 71
				.Width = 640 '350
				.Height = 300
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Caption = "Audio Settings"
				.Moveable = False
				.mbAcceptReprocessEvents = True
			End With

			'chkAudio initial props
			chkAudio = New UICheckBox(oUILib)
			With chkAudio
				.ControlName = "chkAudio"
				.Left = 15
				.Top = 15
				.Width = 123
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Audio Enabled"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
			Me.AddChild(CType(chkAudio, UIControl))

			'chkMusicOn initial props
			chkMusicOn = New UICheckBox(oUILib)
			With chkMusicOn
				.ControlName = "chkMusicOn"
				.Left = 15
				.Top = 35
				.Width = 123
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Music Enabled"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
            Me.AddChild(CType(chkMusicOn, UIControl))

            'chkMusicOn initial props
            chkPostional = New UICheckBox(oUILib)
            With chkPostional
                .ControlName = "chkPostional"
                .Left = 15
                .Top = 55
                .Width = 153
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Positional Override"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.PositionalSound
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If checked, may cause random crashing if your sound hardware cannot handle it properly." & vbCrLf & _
                               "If you experience issues, turn this option off to see if the issues go away. If they do" & vbCrLf & _
                               "update your sound drivers. If that does not remedy the problem, your sound card may not" & vbCrLf & _
                               "be able to handle the option on. Greatly improves sound quality of the game by playing" & vbCrLf & _
                               "closer sound effects over further effects without the use of additional processor."
            End With
            Me.AddChild(CType(chkPostional, UIControl))

			'lblVolume initial props
			lblVolume = New UILabel(oUILib)
			With lblVolume
				.ControlName = "lblVolume"
				.Left = 15
                .Top = 70
				.Width = 134
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Volume Settings"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblVolume, UIControl))

			'lblMaster initial props
			lblMaster = New UILabel(oUILib)
			With lblMaster
				.ControlName = "lblMaster"
				.Left = 40
				.Top = 90
				.Width = 65
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Master:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblMaster, UIControl))

			'hscrMaster initial props
			hscrMaster = New UIScrollBar(oUILib, False)
			With hscrMaster
				.ControlName = "hscrMaster"
				.Left = 150
				.Top = 91
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrMaster, UIControl))

			'lblMusicVol initial props
			lblMusicVol = New UILabel(oUILib)
			With lblMusicVol
				.ControlName = "lblMusicVol"
				.Left = 40
				.Top = 115
				.Width = 65
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Music:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblMusicVol, UIControl))

			'hscrVG0 initial props
			hscrVG0 = New UIScrollBar(oUILib, False)
			With hscrVG0
				.ControlName = "hscrVG0"
				.Left = 150
				.Top = 116
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrVG0, UIControl))

			'lblWpn initial props
			lblWpn = New UILabel(oUILib)
			With lblWpn
				.ControlName = "lblWpn"
				.Left = 40
				.Top = 140
				.Width = 65
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Combat:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblWpn, UIControl))

			'lblEntityA initial props
			lblEntityA = New UILabel(oUILib)
			With lblEntityA
				.ControlName = "lblEntityA"
				.Left = 40
				.Top = 165
				.Width = 96
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Entity Sound:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblEntityA, UIControl))

			'lblEnvirA initial props
			lblEnvirA = New UILabel(oUILib)
			With lblEnvirA
				.ControlName = "lblEnvirA"
				.Left = 40
				.Top = 190
				.Width = 95
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Environment:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblEnvirA, UIControl))

			'lblUnitSpeech initial props
			lblUnitSpeech = New UILabel(oUILib)
			With lblUnitSpeech
				.ControlName = "lblUnitSpeech"
                .Left = 340
                .Top = 165
				.Width = 89
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Unit Speech:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblUnitSpeech, UIControl))

            'lblUI initial props
            lblUI = New UILabel(oUILib)
            With lblUI
                .ControlName = "lblUI"
                .Left = 40
                .Top = 215
                .Width = 65
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "UI:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblUI, UIControl))

			'lblGameNarr initial props
			lblGameNarr = New UILabel(oUILib)
			With lblGameNarr
				.ControlName = "lblGameNarr"
                .Left = 340
                .Top = 190
				.Width = 95
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Game Voice:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblGameNarr, UIControl))

            'lblTutorialVoice initial props
            lblTutorialVoice = New UILabel(oUILib)
            With lblTutorialVoice
                .ControlName = "lblTutorialVoice"
                .Left = 340
                .Top = 215
                .Width = 105
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Tutorial Voice:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTutorialVoice, UIControl))

            'hscrVG1 initial props
			hscrVG1 = New UIScrollBar(oUILib, False)
			With hscrVG1
				.ControlName = "hscrVG1"
				.Left = 150
				.Top = 141
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrVG1, UIControl))

			'hscrVG2 initial props
			hscrVG2 = New UIScrollBar(oUILib, False)
			With hscrVG2
				.ControlName = "hscrVG2"
				.Left = 150
				.Top = 166
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrVG2, UIControl))

			'hscrVG3 initial props
			hscrVG3 = New UIScrollBar(oUILib, False)
			With hscrVG3
				.ControlName = "hscrVG3"
				.Left = 150
				.Top = 191
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrVG3, UIControl))

			'hscrVG4 initial props
			hscrVG4 = New UIScrollBar(oUILib, False)
			With hscrVG4
				.ControlName = "hscrVG4"
                .Left = 450
                .Top = 166
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrVG4, UIControl))

			'hscrVG5 initial props
			hscrVG5 = New UIScrollBar(oUILib, False)
			With hscrVG5
				.ControlName = "hscrVG5"
                .Left = 150
                .Top = 216
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrVG5, UIControl))

			'hscrVG6 initial props
			hscrVG6 = New UIScrollBar(oUILib, False)
			With hscrVG6
				.ControlName = "hscrVG6"
                .Left = 450
                .Top = 191
				.Width = 150
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrVG6, UIControl))

            'hscrVG7 initial props
            hscrVG7 = New UIScrollBar(oUILib, False)
            With hscrVG7
                .ControlName = "hscrVG7"
                .Left = 450
                .Top = 216
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 0
                .MaxValue = 100
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrVG7, UIControl))

			FillValues()

			mbLoading = False
		End Sub

		Private Sub FillValues()

			chkAudio.Value = muSettings.AudioOn

			If goSound Is Nothing = False Then
				hscrMaster.Value = goSound.MasterVolume ' DirectXVolumeToLinear(goSound.MasterVolume, hscrMaster.MaxValue)
				chkMusicOn.Value = goSound.MusicOn
                hscrVG0.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_Music), hscrVG0.MaxValue)
                hscrVG1.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_Fighting), hscrVG1.MaxValue)
                hscrVG2.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_EntityAmbience), hscrVG2.MaxValue)
                hscrVG3.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_EnvirAmbience), hscrVG3.MaxValue)
                hscrVG4.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_UnitSpeech), hscrVG4.MaxValue)
                hscrVG5.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_UserInterface), hscrVG5.MaxValue)
                hscrVG6.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_GameNarrative), hscrVG6.MaxValue)
                hscrVG7.Value = SoundMgr.DirectXVolumeToLinear(goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_TutorialVoice), hscrVG7.MaxValue)
                chkMusicOn.Enabled = True
				hscrMaster.Enabled = True
				hscrVG0.Enabled = True
				hscrVG1.Enabled = True
				hscrVG2.Enabled = True
				hscrVG3.Enabled = True
				hscrVG4.Enabled = True
				hscrVG5.Enabled = True
				hscrVG6.Enabled = True
                hscrVG7.Enabled = True
            Else
                chkMusicOn.Value = False
                chkMusicOn.Enabled = False
                hscrMaster.Enabled = False
                hscrVG0.Enabled = False
                hscrVG1.Enabled = False
                hscrVG2.Enabled = False
                hscrVG3.Enabled = False
                hscrVG4.Enabled = False
                hscrVG5.Enabled = False
                hscrVG6.Enabled = False
                hscrVG7.Enabled = False
            End If

			ReDim mlLastVals(SoundMgr.VolumeGroup.eVG_Count - 1)
			ReDim msTestWAV(SoundMgr.VolumeGroup.eVG_Count - 1)
			ReDim mlTestIdx(SoundMgr.VolumeGroup.eVG_Count - 1)
			ReDim mlTestUsage(SoundMgr.VolumeGroup.eVG_Count - 1)
			For X As Int32 = 0 To SoundMgr.VolumeGroup.eVG_Count - 1
				If goSound Is Nothing = False Then mlLastVals(X) = goSound.VolumeGroups(X)
				mlTestIdx(X) = -1
				msTestWAV(X) = ""
				mlTestUsage(X) = SoundMgr.SoundUsage.eALWAYS_PLAY
			Next X
			msTestWAV(SoundMgr.VolumeGroup.eVG_Fighting) = "Explosions\SmallGroundDeath1.wav"
			mlTestUsage(SoundMgr.VolumeGroup.eVG_Fighting) = SoundMgr.SoundUsage.eWeaponsFire
			msTestWAV(SoundMgr.VolumeGroup.eVG_EntityAmbience) = "Unit Sounds\MediumShipRoar2.wav"
			mlTestUsage(SoundMgr.VolumeGroup.eVG_EntityAmbience) = SoundMgr.SoundUsage.eUnitSounds
			msTestWAV(SoundMgr.VolumeGroup.eVG_UserInterface) = "UserInterface\ButtonClick.wav"
			mlTestUsage(SoundMgr.VolumeGroup.eVG_UserInterface) = SoundMgr.SoundUsage.eUserInterface
			msTestWAV(SoundMgr.VolumeGroup.eVG_GameNarrative) = "Game Narrative\ProductionComplete.wav"
			mlTestUsage(SoundMgr.VolumeGroup.eVG_GameNarrative) = SoundMgr.SoundUsage.eNarrative
			msTestWAV(SoundMgr.VolumeGroup.eVG_UnitSpeech) = "Unit Speech\General_Yes_2_1.wav"
            mlTestUsage(SoundMgr.VolumeGroup.eVG_UnitSpeech) = SoundMgr.SoundUsage.eUnitSpeech
            msTestWAV(SoundMgr.VolumeGroup.eVG_TutorialVoice) = "Tutorial\line 194.wav"
            mlTestUsage(SoundMgr.VolumeGroup.eVG_TutorialVoice) = SoundMgr.SoundUsage.eTutorialVoice

		End Sub

        Private Sub UpdateVolumeGroups() Handles hscrVG0.ValueChange, hscrVG1.ValueChange, hscrVG2.ValueChange, hscrVG3.ValueChange, hscrVG4.ValueChange, hscrVG5.ValueChange, hscrVG6.ValueChange, hscrVG7.ValueChange, hscrMaster.ValueChange
            If mbLoading = True Then Return
            If goSound Is Nothing = False Then
                'Go ahead and update all volume groups...
                goSound.MasterVolume = hscrMaster.Value 'goSound.LinearVolumeToDirectX(hscrMaster.Value, hscrMaster.MaxValue)
                Dim oINI As New InitFile
                oINI.WriteString("AUDIO", "MasterVolume", goSound.MasterVolume.ToString)
                oINI = Nothing

                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_Music) = SoundMgr.LinearVolumeToDirectX(hscrVG0.Value, hscrVG0.MaxValue)
                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_Fighting) = SoundMgr.LinearVolumeToDirectX(hscrVG1.Value, hscrVG1.MaxValue)
                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_EntityAmbience) = SoundMgr.LinearVolumeToDirectX(hscrVG2.Value, hscrVG2.MaxValue)
                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_EnvirAmbience) = SoundMgr.LinearVolumeToDirectX(hscrVG3.Value, hscrVG3.MaxValue)
                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_UnitSpeech) = SoundMgr.LinearVolumeToDirectX(hscrVG4.Value, hscrVG4.MaxValue)
                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_UserInterface) = SoundMgr.LinearVolumeToDirectX(hscrVG5.Value, hscrVG5.MaxValue)
                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_GameNarrative) = SoundMgr.LinearVolumeToDirectX(hscrVG6.Value, hscrVG6.MaxValue)
                goSound.VolumeGroups(SoundMgr.VolumeGroup.eVG_TutorialVoice) = SoundMgr.LinearVolumeToDirectX(hscrVG7.Value, hscrVG7.MaxValue)
                goSound.UpdateSoundVolumes()

                For X As Int32 = 0 To SoundMgr.VolumeGroup.eVG_Count - 1
                    If goSound.VolumeGroups(X) <> mlLastVals(X) Then
                        If msTestWAV(X) <> "" Then
                            If mlTestIdx(X) <> -1 Then goSound.StopSound(mlTestIdx(X))
                            mlTestIdx(X) = goSound.StartSound(msTestWAV(X), False, CType(mlTestUsage(X), SoundMgr.SoundUsage), Nothing, Nothing)
                            Exit For
                        End If
                    End If
                Next

                For X As Int32 = 0 To SoundMgr.VolumeGroup.eVG_Count - 1
                    mlLastVals(X) = goSound.VolumeGroups(X)
                Next X

            End If
        End Sub

		Private Sub chkMusicOn_Click() Handles chkMusicOn.Click
			If goSound Is Nothing = False Then
				goSound.MusicOn = chkMusicOn.Value


				If goSound.MusicOn = True Then
					Dim oINI As New InitFile
					oINI.WriteString("AUDIO", "MusicOn", "1")
					oINI = Nothing
					goSound.PlayListType = 1
					goSound.StartMusic()
				Else
					Dim oINI As New InitFile
					oINI.WriteString("AUDIO", "MusicOn", "0")
					oINI = Nothing
					goSound.StopMusic()
				End If
			End If
		End Sub

		Private Sub chkAudio_Click() Handles chkAudio.Click
			muSettings.AudioOn = chkAudio.Value
			If muSettings.AudioOn = True Then
				If goSound Is Nothing Then
					goSound = New SoundMgr(frmMain)
				End If
			Else
				goSound.DisposeMe()
				goSound = Nothing
			End If
			FillValues()
		End Sub

        Private Sub chkPostional_Click() Handles chkPostional.Click
            muSettings.PositionalSound = chkPostional.Value
        End Sub
    End Class

	'Interface created from Interface Builder
	Private Class fraGraphics
		Inherits UIWindow

		Private WithEvents chkWindowed As UICheckBox
		Private WithEvents lblRes As UILabel
		Private WithEvents cboResolution As UIComboBox
		Private WithEvents lblFOWRes As UILabel
		Private WithEvents cboFOWRes As UIComboBox
		Private WithEvents lblModelRes As UILabel
		Private WithEvents cboTextureRes As UIComboBox
		Private WithEvents lblWaterRes As UILabel
		Private WithEvents cboWaterRes As UIComboBox
		Private WithEvents lblEntityClipPlane As UILabel
		Private WithEvents hscrEntityClipPlane As UIScrollBar
		Private WithEvents chkSpecular As UICheckBox
		Private WithEvents chkShowMinimap As UICheckBox
		Private WithEvents chkDrawGrid As UICheckBox
		Private WithEvents chkRenderCache As UICheckBox
		Private WithEvents chkScreenShake As UICheckBox
		Private WithEvents lblVertex As UILabel
		Private WithEvents cboVertex As UIComboBox

		Private WithEvents lblTerrainTexRes As UILabel
		Private WithEvents cboTerrainTexRes As UIComboBox

		Private WithEvents chkHiResPlanets As UICheckBox
		Private WithEvents chkSmoothFOW As UICheckBox

		Private lblGlowFX As UILabel
		Private WithEvents hscrGlowFX As UIScrollBar

		Private lblEngineFX As UILabel
		Private WithEvents cboEngineFX As UIComboBox
		Private lblBurnFX As UILabel
		Private WithEvents cboBurnFX As UIComboBox
		Private WithEvents chkShieldFX As UICheckBox

		Private lblPlanetFX As UILabel
		Private WithEvents cboPlanetFX As UIComboBox

		Private mbLoading As Boolean = True

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraGraphics initial props
			With Me
				.ControlName = "fraGraphics"
				.Left = 145
				.Top = 123
				.Width = 640 '350
				.Height = 300
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.mbAcceptReprocessEvents = True
			End With

			'chkWindowed initial props
			chkWindowed = New UICheckBox(oUILib)
			With chkWindowed
				.ControlName = "chkWindowed"
				.Left = 15
				.Top = 10
				.Width = 92
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Windowed"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
			Me.AddChild(CType(chkWindowed, UIControl))

			'lblRes initial props
			lblRes = New UILabel(oUILib)
			With lblRes
				.ControlName = "lblRes"
				.Left = 15
				.Top = 55
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Fullscreen Resolution:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblRes, UIControl))

			'lblVertex initial props
			lblVertex = New UILabel(oUILib)
			With lblVertex
				.ControlName = "lblVertex"
				.Left = 15
				.Top = 30
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Vertex Processing:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblVertex, UIControl))

			'lblFOWRes initial props
			lblFOWRes = New UILabel(oUILib)
			With lblFOWRes
				.ControlName = "lblFOWRes"
				.Left = 15
				.Top = 80
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Fog of War Resolution:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblFOWRes, UIControl))

			'lblModelRes initial props
			lblModelRes = New UILabel(oUILib)
			With lblModelRes
				.ControlName = "lblModelRes"
				.Left = 15
				.Top = 105
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Texture Resolution:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblModelRes, UIControl))

			'lblWaterRes initial props
			lblWaterRes = New UILabel(oUILib)
			With lblWaterRes
				.ControlName = "lblWaterRes"
				.Left = 15
				.Top = 130
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Water Resolution:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblWaterRes, UIControl))

			'lblTerrainTexRes initial props
			lblTerrainTexRes = New UILabel(oUILib)
			With lblTerrainTexRes
				.ControlName = "lblTerrainTexRes"
				.Left = 15
				.Top = lblWaterRes.Top + 25
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Terrain Texture Res:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "Texture resolution for the terrain of planets." & vbCrLf & "Setting this value to low can increase performance on fill-rate limited video cards."
			End With
			Me.AddChild(CType(lblTerrainTexRes, UIControl))

			'lblEntityClipPlane initial props
			lblEntityClipPlane = New UILabel(oUILib)
			With lblEntityClipPlane
				.ControlName = "lblEntityClipPlane"
				.Left = 15
				.Top = lblTerrainTexRes.Top + 30
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Entity Clip Plane:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblEntityClipPlane, UIControl))

			'lblGlowFX initial props
			lblGlowFX = New UILabel(oUILib)
			With lblGlowFX
				.ControlName = "lblGlowFX"
				.Left = 15
				.Top = lblEntityClipPlane.Top + 30
				.Width = 157
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Glow FX Intensity:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblGlowFX, UIControl))

			'hscrEntityClipPlane initial props
			hscrEntityClipPlane = New UIScrollBar(oUILib, False)
			With hscrEntityClipPlane
				.ControlName = "hscrEntityClipPlane"
				.Left = 180
				.Top = lblEntityClipPlane.Top
				.Width = 155
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 100
				.MaxValue = 400
				.MinValue = 100
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(hscrEntityClipPlane, UIControl))

			'hscrGlowFX initial props
			hscrGlowFX = New UIScrollBar(oUILib, False)
			With hscrGlowFX
				.ControlName = "hscrGlowFX"
				.Left = 180
				.Top = lblGlowFX.Top
				.Width = 155
				.Height = 18
				.Enabled = True
				.Visible = True
				.Value = 5
				.MaxValue = 10
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 4
				.ReverseDirection = False
				.mbAcceptReprocessEvents = False
			End With
			Me.AddChild(CType(hscrGlowFX, UIControl))

			'chkSpecular initial props
			chkSpecular = New UICheckBox(oUILib)
			With chkSpecular
				.ControlName = "chkSpecular"
				.Left = 355	'15
				.Top = 10 '180
				.Width = 139
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Specular Lighting"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
			Me.AddChild(CType(chkSpecular, UIControl))

			'chkShowMinimap initial props
			chkShowMinimap = New UICheckBox(oUILib)
			With chkShowMinimap
				.ControlName = "chkShowMinimap"
				.Left = 355	'15
				.Top = 30 '205
				.Width = 122
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Show Mini Map"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
			Me.AddChild(CType(chkShowMinimap, UIControl))

			'chkDrawGrid initial props
			chkDrawGrid = New UICheckBox(oUILib)
			With chkDrawGrid
				.ControlName = "chkDrawGrid"
				.Left = 355	'15
				.Top = 55 '230
				.Width = 152
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Draw Grid in Space"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
			Me.AddChild(CType(chkDrawGrid, UIControl))

			'chkSmoothFOW initial props
			chkSmoothFOW = New UICheckBox(oUILib)
			With chkSmoothFOW
				.ControlName = "chkSmoothFOW"
				.Left = 355	'15
				.Top = 80 '255
				.Width = 153
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Smooth Fog-of-War"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
				.ToolTipText = "Click to toggle smoothing the fog-of-war texture layer."
			End With
			Me.AddChild(CType(chkSmoothFOW, UIControl))

			'chkRenderCache initial props
			chkRenderCache = New UICheckBox(oUILib)
			With chkRenderCache
				.ControlName = "chkRenderCache"
				.Left = 355	'15
				.Top = 105 '255
				.Width = 181
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Render Mineral Caches"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
			Me.AddChild(CType(chkRenderCache, UIControl))

			'chkScreenShake initial props
			chkScreenShake = New UICheckBox(oUILib)
			With chkScreenShake
				.ControlName = "chkScreenShake"
				.Left = 355	'15
				.Top = 130 '280
				.Width = 180
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Screen Shake Enabled"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
			End With
			Me.AddChild(CType(chkScreenShake, UIControl))

			'lblEngineFX initial props
			lblEngineFX = New UILabel(oUILib)
			With lblEngineFX
				.ControlName = "lblEngineFX"
				.Left = 355
				.Top = 150
				.Width = 66
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Engine FX:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "Manages the Density of Engine FX Particles" & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."
			End With
			Me.AddChild(CType(lblEngineFX, UIControl))

			'lblBurnFX initial props
			lblBurnFX = New UILabel(oUILib)
			With lblBurnFX
				.ControlName = "lblBurnFX"
				.Left = 355
				.Top = 180
				.Width = 66
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Burn FX:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "Manages the Density of Burn (fire) FX particles." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."
			End With
			Me.AddChild(CType(lblBurnFX, UIControl))

			'lblPlanetFX initial props
			lblPlanetFX = New UILabel(oUILib)
			With lblPlanetFX
				.ControlName = "lblPlanetFX"
				.Left = 355
				.Top = 210
				.Width = 66
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Planet FX:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "Manages the Density of Planet particle effects." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this will increase performance on most systems."
			End With
			Me.AddChild(CType(lblPlanetFX, UIControl))

			'chkShieldFX initial props
			chkShieldFX = New UICheckBox(oUILib)
			With chkShieldFX
				.ControlName = "chkShieldFX"
				.Left = 355
				.Top = 230
				.Width = 122
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Render Shield FX"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
				.ToolTipText = "If on, renders shield effects when an entity is hit." & vbCrLf & "Unchecking this value can increase performance in high combat situations."
			End With
			Me.AddChild(CType(chkShieldFX, UIControl))

			'chkHiResPlanets initial props
			chkHiResPlanets = New UICheckBox(oUILib)
			With chkHiResPlanets
				.ControlName = "chkHiResPlanets"
				.Left = 355
				.Top = 250
				.Width = 155
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Hi-Res Planet Textures"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
				.Value = False
				.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
				.ToolTipText = "If On, Planet Textures are rendered. If Off, the textures are calculated." & vbCrLf & "Turning this On will increase load times for NEW solar systems." & vbCrLf & "To see changes to this setting, you will need to reload the solar system."
			End With
			Me.AddChild(CType(chkHiResPlanets, UIControl))

			'==== AND NOW THE COMBO BOXES ====
			'cboPlanetFX initial props
			cboPlanetFX = New UIComboBox(oUILib)
			With cboPlanetFX
				.ControlName = "cboPlanetFX"
				.Left = 450
				.Top = 209
				.Width = 155
				.Height = 18
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = "Manages the Density of Planet particle effects." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this will increase performance on most systems."
            End With
            Me.AddChild(CType(cboPlanetFX, UIControl))

            'cboBurnFX initial props
            cboBurnFX = New UIComboBox(oUILib)
            With cboBurnFX
                .ControlName = "cboBurnFX"
                .Left = 450
                .Top = 179
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = "Manages the Density of Burn (fire) FX particles." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."
            End With
            Me.AddChild(CType(cboBurnFX, UIControl))

            'cboEngineFX initial props
            cboEngineFX = New UIComboBox(oUILib)
            With cboEngineFX
                .ControlName = "cboEngineFX"
                .Left = 450
                .Top = 149
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = "Manages the Density of Engine FX Particles" & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."
            End With
            Me.AddChild(CType(cboEngineFX, UIControl))

            'cboTerrainTexRes initial props
            cboTerrainTexRes = New UIComboBox(oUILib)
            With cboTerrainTexRes
                .ControlName = "cboTerrainTexRes"
                .Left = 180
                .Top = lblTerrainTexRes.Top
                .Width = 155
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
                .ToolTipText = lblTerrainTexRes.ToolTipText
            End With
            Me.AddChild(CType(cboTerrainTexRes, UIControl))

            'cboWaterRes initial props
            cboWaterRes = New UIComboBox(oUILib)
            With cboWaterRes
                .ControlName = "cboWaterRes"
                .Left = 180
                .Top = 130
                .Width = 155
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
            Me.AddChild(CType(cboWaterRes, UIControl))

            'cboTextureRes initial props
            cboTextureRes = New UIComboBox(oUILib)
            With cboTextureRes
                .ControlName = "cboTextureRes"
                .Left = 180
                .Top = 105
                .Width = 155
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
            Me.AddChild(CType(cboTextureRes, UIControl))

            'cboFOWRes initial props
            cboFOWRes = New UIComboBox(oUILib)
            With cboFOWRes
                .ControlName = "cboFOWRes"
                .Left = 180
                .Top = 80
                .Width = 155
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
            Me.AddChild(CType(cboFOWRes, UIControl))

            'cboResolution initial props
            cboResolution = New UIComboBox(oUILib)
            With cboResolution
                .ControlName = "cboResolution"
                .Left = 180
                .Top = 55
                .Width = 155
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
            Me.AddChild(CType(cboResolution, UIControl))

            'cboVertex initial props
            cboVertex = New UIComboBox(oUILib)
            With cboVertex
                .ControlName = "cboVertex"
                .Left = 180
                .Top = 30
                .Width = 155
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
            Me.AddChild(CType(cboVertex, UIControl))

            FillValues()
            mbLoading = False
        End Sub

        Private Sub FillValues()
            Dim lMaxTexRes As Int32 = Math.Min(MyBase.moUILib.oDevice.DeviceCaps.MaxTextureWidth, MyBase.moUILib.oDevice.DeviceCaps.MaxTextureHeight)

            chkWindowed.Value = MyBase.moUILib.oDevice.PresentationParameters.Windowed

            chkSmoothFOW.Enabled = moUILib.oDevice.DeviceCaps.TextureFilterCaps.SupportsMagnifyLinear
            If chkSmoothFOW.Enabled = False Then chkSmoothFOW.Value = False Else chkSmoothFOW.Value = muSettings.SmoothFOW

            With cboEngineFX
                .Clear()
                .AddItem("Off")
                .ItemData(.NewIndex) = 0
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .AddItem("Full")
                .ItemData(.NewIndex) = 4
                .FindComboItemData(muSettings.EngineFXParticles)
            End With
            With cboBurnFX
                .Clear()
                .AddItem("Off")
                .ItemData(.NewIndex) = 0
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .AddItem("Full")
                .ItemData(.NewIndex) = 4
                .FindComboItemData(muSettings.BurnFXParticles)
            End With
            chkShieldFX.Value = muSettings.RenderShieldFX
            With cboPlanetFX
                .Clear()
                .AddItem("Off")
                .ItemData(.NewIndex) = 0
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .AddItem("Full")
                .ItemData(.NewIndex) = 4
                .FindComboItemData(muSettings.PlanetFXParticles)
            End With

            With cboTerrainTexRes
                .Clear()
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .FindComboItemData(muSettings.TerrainTextureResolution)
            End With

            'fill in our vertex processing flags
            cboVertex.Clear()
            cboVertex.AddItem("Hardware")
            cboVertex.ItemData(cboVertex.NewIndex) = CreateFlags.HardwareVertexProcessing
            cboVertex.AddItem("Mixed")
            cboVertex.ItemData(cboVertex.NewIndex) = CreateFlags.MixedVertexProcessing
            cboVertex.AddItem("Software")
            cboVertex.ItemData(cboVertex.NewIndex) = CreateFlags.SoftwareVertexProcessing
            If MyBase.moUILib.oDevice.CreationParameters.Behavior.HardwareVertexProcessing = True Then
                cboVertex.FindComboItemData(CreateFlags.HardwareVertexProcessing)
            ElseIf MyBase.moUILib.oDevice.CreationParameters.Behavior.SoftwareVertexProcessing = True Then
                cboVertex.FindComboItemData(CreateFlags.SoftwareVertexProcessing)
            Else : cboVertex.FindComboItemData(CreateFlags.MixedVertexProcessing)
            End If

            'TODO: Remove this when display settings can be updated on the fly...
            chkWindowed.Enabled = False
            cboResolution.Enabled = False

            chkHiResPlanets.Value = muSettings.HiResPlanetTexture
            If MyBase.moUILib.oDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = False Then
                muSettings.HiResPlanetTexture = False
                chkHiResPlanets.Value = False
                chkHiResPlanets.Enabled = False
            End If

            'FOW Texture can be OFF, LOW (512), MEDIUM (1024) or HIGH (2048)
            'TODO: Include SIMPLE in there eventually...
            cboFOWRes.Clear()
            cboFOWRes.AddItem("Off - No FOW shading")
            cboFOWRes.ItemData(cboFOWRes.NewIndex) = 0

            'Texture Res can be Very Low, Low, Very Low and High
            cboTextureRes.Clear()
            cboTextureRes.AddItem("Very Low Resolution")
            cboTextureRes.ItemData(cboTextureRes.NewIndex) = CInt(EngineSettings.eTextureResOptions.eVeryLowResTextures)
            cboTextureRes.AddItem("Low Resolution")
            cboTextureRes.ItemData(cboTextureRes.NewIndex) = CInt(EngineSettings.eTextureResOptions.eLowResTextures)
            cboTextureRes.AddItem("Normal Resolution")
            cboTextureRes.ItemData(cboTextureRes.NewIndex) = CInt(EngineSettings.eTextureResOptions.eNormResTextures)

            'Water Quad Res can be... OFF, SIMPLE (do not draw), LOW (64), MEDIUM (128), HIGH (256)
            cboWaterRes.AddItem("Plain Water")
            cboWaterRes.ItemData(cboWaterRes.NewIndex) = 0
            If MyBase.moUILib.oDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
                cboWaterRes.AddItem("Shader 2.0")
                cboWaterRes.ItemData(cboWaterRes.NewIndex) = 1
            End If

            'cboWaterRes.AddItem("Low Resolution")
            'cboWaterRes.ItemData(cboWaterRes.NewIndex) = 64
            'cboWaterRes.AddItem("Medium Resolution")
            'cboWaterRes.ItemData(cboWaterRes.NewIndex) = 128
            'cboWaterRes.AddItem("High Resolution")
            'cboWaterRes.ItemData(cboWaterRes.NewIndex) = 256

            If lMaxTexRes >= 512 Then
                cboFOWRes.AddItem("Normal Resolution")
                cboFOWRes.ItemData(cboFOWRes.NewIndex) = 512

                cboTextureRes.AddItem("High Resolution")
                cboTextureRes.ItemData(cboTextureRes.NewIndex) = CInt(EngineSettings.eTextureResOptions.eHiResTextures)
            End If
            'If lMaxTexRes >= 1024 Then
            '    cboFOWRes.AddItem("Medium Resolution")
            '    cboFOWRes.ItemData(cboFOWRes.NewIndex) = 1024
            'End If
            'If lMaxTexRes >= 2048 Then
            '    cboFOWRes.AddItem("High Resolution")
            '    cboFOWRes.ItemData(cboFOWRes.NewIndex) = 2048
            'End If

            cboFOWRes.FindComboItemData(muSettings.FOWTextureResolution)
            cboTextureRes.FindComboItemData(muSettings.ModelTextureResolution)

            'If muSettings.DrawProceduralWater = False Then
            '	If muSettings.WaterQuadWidthHeight = 32 Then
            '		cboWaterRes.FindComboItemData(0)
            '	Else : cboWaterRes.FindComboItemData(1)
            '	End If
            'Else
            '	cboWaterRes.FindComboItemData(muSettings.WaterQuadWidthHeight)
            'End If
            cboWaterRes.FindComboItemData(0)
            If muSettings.WaterRenderMethod = 1 Then
                If MyBase.moUILib.oDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
                    cboWaterRes.FindComboItemData(1)
                Else : muSettings.WaterRenderMethod = 0
                End If
            End If

            hscrEntityClipPlane.Value = muSettings.EntityClipPlane \ 100
            hscrGlowFX.Value = CInt(muSettings.PostGlowAmt * 10)
            chkSpecular.Value = muSettings.SpecularEnabled
            chkShowMinimap.Value = muSettings.ShowMiniMap
            chkDrawGrid.Value = muSettings.DrawGrid
            chkRenderCache.Value = muSettings.RenderMineralCaches
            chkScreenShake.Value = muSettings.ScreenShakeEnabled
        End Sub

        Private Sub chkWindowed_Click() Handles chkWindowed.Click
            If mbLoading = True Then Return
            'TODO: Fix this for windowed mode...
        End Sub

        Private Sub cboResolution_ItemSelected(ByVal lItemIndex As Integer) Handles cboResolution.ItemSelected
            If mbLoading = True Then Return
            'TODO: adjust the resolution....
        End Sub

        Private Sub cboFOWRes_ItemSelected(ByVal lItemIndex As Integer) Handles cboFOWRes.ItemSelected
            If mbLoading = True Then Return
            If lItemIndex = -1 Then Return

            Dim lVal As Int32 = cboFOWRes.ItemData(lItemIndex)
            If lVal = 0 Then
                muSettings.ShowFOWTerrainShading = False
            Else
                muSettings.ShowFOWTerrainShading = True
                muSettings.FOWTextureResolution = cboFOWRes.ItemData(lItemIndex)
            End If
            'This tells the engine to recreate the FOW texture
            TerrainClass.ReleaseVisibleTexture()
        End Sub

        Private Sub cboTextureRes_ItemSelected(ByVal lItemIndex As Integer) Handles cboTextureRes.ItemSelected
            If mbLoading = True Then Return
            If lItemIndex = -1 Then Return

            frmMain.Cursor = Cursors.WaitCursor

            muSettings.ModelTextureResolution = CType(cboTextureRes.ItemData(lItemIndex), EngineSettings.eTextureResOptions)
            goResMgr.UpdateModelTextureResolution()

            frmMain.Cursor = Cursors.Arrow
        End Sub

        Private Sub cboWaterRes_ItemSelected(ByVal lItemIndex As Integer) Handles cboWaterRes.ItemSelected
            If mbLoading = True Then Return
            If lItemIndex = -1 Then Return

            Dim lVal As Int32 = cboWaterRes.ItemData(lItemIndex)

            muSettings.WaterRenderMethod = lVal
            TerrainClass.ForceWaterCreation()

            'If lVal = 0 Then
            '	muSettings.DrawProceduralWater = False
            '	muSettings.WaterQuadWidthHeight = 32

            'ElseIf lVal = 1 Then
            '	muSettings.DrawProceduralWater = False
            '	muSettings.WaterQuadWidthHeight = 64
            '	TerrainClass.ForceWaterCreation()
            'Else
            '	muSettings.DrawProceduralWater = True
            '	muSettings.WaterQuadWidthHeight = lVal
            '	TerrainClass.ForceWaterCreation()
            'End If
        End Sub

        Private Sub hscrEntityClipPlane_ValueChange() Handles hscrEntityClipPlane.ValueChange
            If mbLoading = True Then Return
            muSettings.EntityClipPlane = (hscrEntityClipPlane.Value * 100)
        End Sub

        Private Sub chkSpecular_Click() Handles chkSpecular.Click
            If mbLoading = True Then Return
            muSettings.SpecularEnabled = chkSpecular.Value
        End Sub

        Private Sub chkShowMinimap_Click() Handles chkShowMinimap.Click
            If mbLoading = True Then Return
            muSettings.ShowMiniMap = chkShowMinimap.Value
        End Sub

        Private Sub chkDrawGrid_Click() Handles chkDrawGrid.Click
            If mbLoading = True Then Return
            muSettings.DrawGrid = chkDrawGrid.Value
        End Sub

        Private Sub chkRenderCache_Click() Handles chkRenderCache.Click
            If mbLoading = True Then Return
            muSettings.RenderMineralCaches = chkRenderCache.Value
        End Sub

        Private Sub chkScreenShake_Click() Handles chkScreenShake.Click
            If mbLoading = True Then Return
            muSettings.ScreenShakeEnabled = chkScreenShake.Value
        End Sub

        Private Sub cboVertex_ItemSelected(ByVal lItemIndex As Integer) Handles cboVertex.ItemSelected
            If mbLoading = True Then Return

            If lItemIndex > -1 Then
                Dim oINI As New InitFile()
                oINI.WriteString("GRAPHICS", "VertexProcessing", cboVertex.ItemData(cboVertex.ListIndex).ToString)
                oINI = Nothing
                MyBase.moUILib.AddNotification("This will require you to restart Beyond Protocol to take effect!", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End Sub

        Private Sub cboEngineFX_ItemSelected(ByVal lItemIndex As Integer) Handles cboEngineFX.ItemSelected
            If mbLoading = True Then Return
            muSettings.EngineFXParticles = cboEngineFX.ItemData(lItemIndex)
            goPFXEngine32.ResetEngineFXPCnts()
        End Sub

        Private Sub cboBurnFX_ItemSelected(ByVal lItemIndex As Integer) Handles cboBurnFX.ItemSelected
            If mbLoading = True Then Return
            muSettings.BurnFXParticles = cboBurnFX.ItemData(lItemIndex)
            goPFXEngine32.ResetBurnFXPCnts()
        End Sub

        Private Sub chkShieldFX_Click() Handles chkShieldFX.Click
            If mbLoading = True Then Return
            muSettings.RenderShieldFX = chkShieldFX.Value
            If muSettings.RenderShieldFX = False Then goShldMgr.TurnOffMgr()
        End Sub

        Private Sub cboPlanetFX_ItemSelected(ByVal lItemIndex As Integer) Handles cboPlanetFX.ItemSelected
            If mbLoading = True Then Return
            muSettings.PlanetFXParticles = cboPlanetFX.ItemData(lItemIndex)
        End Sub

        Private Sub hscrGlowFX_ValueChange() Handles hscrGlowFX.ValueChange
            If mbLoading = True Then Return
            muSettings.PostGlowAmt = hscrGlowFX.Value / 10.0F
        End Sub

        Private Sub chkHiResPlanets_Click() Handles chkHiResPlanets.Click
            If mbLoading = True Then Return
            muSettings.HiResPlanetTexture = chkHiResPlanets.Value
            muSettings.SaveSettings()
            If goGalaxy Is Nothing = False Then
                For X As Int32 = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X) Is Nothing = False Then
                        For Y As Int32 = 0 To goGalaxy.moSystems(X).PlanetUB
                            If goGalaxy.moSystems(X).moPlanets(Y) Is Nothing = False Then goGalaxy.moSystems(X).moPlanets(Y).ClearPlanetTexture()
                        Next Y
                    End If
                Next X
            End If
        End Sub

        Private Sub cboTerrainTexRes_ItemSelected(ByVal lItemIndex As Integer) Handles cboTerrainTexRes.ItemSelected
            If mbLoading = True Then Return
            If lItemIndex = -1 Then Return

            frmMain.Cursor = Cursors.WaitCursor

            muSettings.TerrainTextureResolution = cboTerrainTexRes.ItemData(lItemIndex)
            TerrainClass.bReloadVertexBuffer = True

            frmMain.Cursor = Cursors.Arrow
        End Sub

        Private Sub chkSmoothFOW_Click() Handles chkSmoothFOW.Click
            muSettings.SmoothFOW = chkSmoothFOW.Value
        End Sub
    End Class

    'Interface created from Interface Builder
    Private Class fraControls
        Inherits UIWindow

        Private lblZoom As UILabel
        Private WithEvents hscrZoomRate As UIScrollBar
        Private lblGalScroll As UILabel
        Private WithEvents hscrScroll0 As UIScrollBar
        Private lblScrollSys As UILabel
        Private lblScrollPlanet As UILabel
        Private WithEvents hscrScroll3 As UIScrollBar
        Private WithEvents hscrScroll5 As UIScrollBar
        Private lblNotificationDly As UILabel
        Private WithEvents hscrNoteDly As UIScrollBar
        Private lblNoteDlyVal As UILabel

        Private lblScrollArea As UILabel
        Private WithEvents hscrScrollArea As UIScrollBar

        Private WithEvents chkCtrlQExits As UICheckBox
        Private WithEvents chkZoomCausesViewChange As UICheckBox

        Private lblKeyMap As UILabel
        Private txtKeyMap As UITextBox

        Private mbLoading As Boolean = True

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraControls initial props
            With Me
                .ControlName = "fraControls"
                .Left = 168
                .Top = 117
                .Width = 640 '350
                .Height = 300
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .mbAcceptReprocessEvents = True
            End With

            'lblZoom initial props
            lblZoom = New UILabel(oUILib)
            With lblZoom
                .ControlName = "lblZoom"
                .Left = 15
                .Top = 15
                .Width = 82
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Zoom Rate:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblZoom, UIControl))

            'hscrZoomRate initial props
            hscrZoomRate = New UIScrollBar(oUILib, False)
            With hscrZoomRate
                .ControlName = "hscrZoomRate"
                .Left = 180
                .Top = 16
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 100
                .MaxValue = 1000
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrZoomRate, UIControl))

            'lblGalScroll initial props
            lblGalScroll = New UILabel(oUILib)
            With lblGalScroll
                .ControlName = "lblGalScroll"
                .Left = 15
                .Top = 40
                .Width = 147
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Scroll Rate (Galaxy):"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblGalScroll, UIControl))

            'hscrScroll0 initial props
            hscrScroll0 = New UIScrollBar(oUILib, False)
            With hscrScroll0
                .ControlName = "hscrScroll0"
                .Left = 180
                .Top = 41
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 100
                .MaxValue = 1000
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrScroll0, UIControl))

            'lblScrollSys initial props
            lblScrollSys = New UILabel(oUILib)
            With lblScrollSys
                .ControlName = "lblScrollSys"
                .Left = 15
                .Top = 65
                .Width = 147
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Scroll Rate (System):"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblScrollSys, UIControl))

            'lblScrollPlanet initial props
            lblScrollPlanet = New UILabel(oUILib)
            With lblScrollPlanet
                .ControlName = "lblScrollPlanet"
                .Left = 15
                .Top = 90
                .Width = 147
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Scroll Rate (Planet):"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblScrollPlanet, UIControl))

            'hscrScroll3 initial props
            hscrScroll3 = New UIScrollBar(oUILib, False)
            With hscrScroll3
                .ControlName = "hscrScroll3"
                .Left = 180
                .Top = 66
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 100
                .MaxValue = 1000
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrScroll3, UIControl))

            'hscrScroll5 initial props
            hscrScroll5 = New UIScrollBar(oUILib, False)
            With hscrScroll5
                .ControlName = "hscrScroll5"
                .Left = 180
                .Top = 91
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 100
                .MaxValue = 1000
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrScroll5, UIControl))

            'lblNotificationDly initial props
            lblNotificationDly = New UILabel(oUILib)
            With lblNotificationDly
                .ControlName = "lblNotificationDly"
                .Left = 15
                .Top = 140
                .Width = 147
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Notification Delay:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNotificationDly, UIControl))

            'hscrNoteDly initial props
            hscrNoteDly = New UIScrollBar(oUILib, False)
            With hscrNoteDly
                .ControlName = "hscrNoteDly"
                .Left = 180
                .Top = 141
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 5
                .MaxValue = 16
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrNoteDly, UIControl))

            'lblNoteDlyVal initial props
            lblNoteDlyVal = New UILabel(oUILib)
            With lblNoteDlyVal
                .ControlName = "lblNoteDlyVal"
                .Left = 180
                .Top = 160
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "5 Seconds"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(5, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNoteDlyVal, UIControl))

            'lblScrollArea initial props
            lblScrollArea = New UILabel(oUILib)
            With lblScrollArea
                .ControlName = "lblScrollArea"
                .Left = 15
                .Top = 180
                .Width = 147
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Scroll Edge Area:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblScrollArea, UIControl))

            'hscrScrollArea initial props
            hscrScrollArea = New UIScrollBar(oUILib, False)
            With hscrScrollArea
                .ControlName = "hscrScrollArea"
                .Left = 180
                .Top = 181
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Value = 10
                .MaxValue = 20
                .MinValue = 1
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
                .ToolTipText = "Indicates the amount of area on the screen's edge before scrolling begins."
            End With
            Me.AddChild(CType(hscrScrollArea, UIControl))

            'chkCtrlQExits initial props
            chkCtrlQExits = New UICheckBox(oUILib)
            With chkCtrlQExits
                .ControlName = "chkCtrlQExits"
                .Left = 180
                .Top = 210
                .Width = 140
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Ctrl+Q Exits Game"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = False
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If checked, ctrl+Q will exit the game. Useful for German keyboard users who require the @ symbol."
            End With
            Me.AddChild(CType(chkCtrlQExits, UIControl))

            'chkZoomCausesViewChange initial props
            chkZoomCausesViewChange = New UICheckBox(oUILib)
            With chkZoomCausesViewChange
                .ControlName = "chkZoomCausesViewChange"
                .Left = 100
                .Top = 230
                .Width = 205
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Zoom Causes View Change"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = False
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If checked, zooming to certain extents will cause the view to change." & vbCrLf & _
                               "If left unchecked, the only way to change views would be the Home and Backspace keys."
            End With
            Me.AddChild(CType(chkZoomCausesViewChange, UIControl))

            'lblKeyMap initial props
            lblKeyMap = New UILabel(oUILib)
            With lblKeyMap
                .ControlName = "lblKeyMap"
                .Left = hscrZoomRate.Left + hscrZoomRate.Width + 10
                .Top = hscrZoomRate.Top
                .Width = Me.Width - .Left - 10
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Keymap Configuration"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblKeyMap, UIControl))

            'txtKeyMap initial props
            txtKeyMap = New UITextBox(oUILib)
            With txtKeyMap
                .ControlName = "txtKeyMap"
                .Left = hscrZoomRate.Left + hscrZoomRate.Width + 10
                .Top = lblKeyMap.Top + lblKeyMap.Height + 5
                .Width = Me.Width - .Left - 10
                .Height = Me.Height - .Top - 10
                .Enabled = True
                .Visible = True
                .Caption = frmChat.GetKeyMapText.Replace(vbCrLf, vbCrLf & vbCrLf)
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
                .MultiLine = True
                .Locked = True
            End With
            Me.AddChild(CType(txtKeyMap, UIControl))

            FillValues()

            mbLoading = False
        End Sub

        Private Sub FillValues()
            hscrZoomRate.Value = goCamera.ZoomRate
            hscrScroll0.Value = goCamera.ScrollRate(0)
            hscrScroll3.Value = goCamera.ScrollRate(3)
            hscrScroll5.Value = goCamera.ScrollRate(5)
            hscrScrollArea.Value = goCamera.ScrollEdgeWidth

            If muSettings.NotificationDisplayTime = -1 Then
                hscrNoteDly.Value = hscrNoteDly.MaxValue
                lblNoteDlyVal.Caption = "No Fade Out"
            Else
                hscrNoteDly.Value = muSettings.NotificationDisplayTime \ 1000
                lblNoteDlyVal.Caption = (muSettings.NotificationDisplayTime \ 1000) & " seconds"
            End If

            chkCtrlQExits.Value = muSettings.CtrlQExits
            chkZoomCausesViewChange.Value = muSettings.ZoomChangesView
        End Sub

        Private Sub hscrNoteDly_ValueChange() Handles hscrNoteDly.ValueChange
            If mbLoading = True Then Return
            Dim lVal As Int32 = hscrNoteDly.Value
            If lVal = hscrNoteDly.MaxValue Then
                muSettings.NotificationDisplayTime = -1
                lblNoteDlyVal.Caption = "No Fade Out"
            Else
                muSettings.NotificationDisplayTime = lVal * 1000
                lblNoteDlyVal.Caption = lVal & " seconds"
            End If
        End Sub

        Private Sub hscrScroll0_ValueChange() Handles hscrScroll0.ValueChange
            If mbLoading = True Then Return

            If goCamera Is Nothing = False Then goCamera.ScrollRate(0) = hscrScroll0.Value
        End Sub

        Private Sub hscrScroll3_ValueChange() Handles hscrScroll3.ValueChange
            If mbLoading = True Then Return

            If goCamera Is Nothing = False Then goCamera.ScrollRate(3) = hscrScroll3.Value
        End Sub

        Private Sub hscrScroll5_ValueChange() Handles hscrScroll5.ValueChange
            If mbLoading = True Then Return

            If goCamera Is Nothing = False Then goCamera.ScrollRate(5) = hscrScroll5.Value
        End Sub

        Private Sub hscrZoomRate_ValueChange() Handles hscrZoomRate.ValueChange
            If mbLoading = True Then Return

            If goCamera Is Nothing = False Then goCamera.ZoomRate = hscrZoomRate.Value
        End Sub

        Private Sub hscrScrollArea_ValueChange() Handles hscrScrollArea.ValueChange
            If mbLoading = True Then Return

            If goCamera Is Nothing = False Then goCamera.ScrollEdgeWidth = hscrScrollArea.Value

            Dim oINI As InitFile = New InitFile()
            oINI.WriteString("SETTINGS", "ScrollArea", hscrScrollArea.Value.ToString)
            oINI = Nothing
        End Sub

        Private Sub chkCtrlQExits_Click() Handles chkCtrlQExits.Click
            If mbLoading = True Then Return
            muSettings.CtrlQExits = chkCtrlQExits.Value
        End Sub

        Private Sub chkZoomCausesViewChange_Click() Handles chkZoomCausesViewChange.Click
            If mbLoading = True Then Return
            muSettings.ZoomChangesView = chkZoomCausesViewChange.Value
        End Sub
    End Class

    'Interface created from Interface Builder
    Public Class fraGraphics2
        Inherits UIWindow

        Private Enum elOptionSetItem As Int32
            General = 0
            Textures = 1
            Particles = 2
            Lighting = 3
            Shadowing = 4
            InterfaceColors = 5
            ChatWindowColors = 6
            IdentificationColors = 7
            Gameplay = 8
            BuildGhostColor = 9
            MineralCacheColor = 10
            MinimapViewAngle = 11
            WaterOptions = 12
        End Enum

        Private moColorSelect As fraColorSelect

        Private mbIgnoreEvents As Boolean = True

        Private WithEvents lstOptionSet As UIListBox
        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraGraphics initial props
            With Me
                .ControlName = "fraGraphics"
                .Left = 145
                .Top = 123
                .Width = 640 '350
                .Height = 300
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .mbAcceptReprocessEvents = True
            End With

            'lstOptionSet initial props
            lstOptionSet = New UIListBox(oUILib)
            With lstOptionSet
                .ControlName = "lstOptionSet"
                .Left = 5
                .Top = 5
                .Width = 175
                .Height = 290
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(lstOptionSet, UIControl))

            lstOptionSet.Clear()
            With lstOptionSet
                .AddItem("General") : .ItemData(.NewIndex) = elOptionSetItem.General
                .AddItem("Textures") : .ItemData(.NewIndex) = elOptionSetItem.Textures
                .AddItem("Lighting") : .ItemData(.NewIndex) = elOptionSetItem.Lighting
                '.AddItem("Shadows") : .ItemData(.NewIndex) = elOptionSetItem.Shadowing
                .AddItem("Particles") : .ItemData(.NewIndex) = elOptionSetItem.Particles
                .AddItem("Water/Lava/Acid") : .ItemData(.NewIndex) = elOptionSetItem.WaterOptions
                .AddItem("Gameplay") : .ItemData(.NewIndex) = elOptionSetItem.Gameplay
                .AddItem("Interface Colors") : .ItemData(.NewIndex) = elOptionSetItem.InterfaceColors
                .AddItem("Chat Window Colors") : .ItemData(.NewIndex) = elOptionSetItem.ChatWindowColors
                .AddItem("Identification Colors") : .ItemData(.NewIndex) = elOptionSetItem.IdentificationColors
                .AddItem("Build Ghost Colors") : .ItemData(.NewIndex) = elOptionSetItem.BuildGhostColor
                .AddItem("Mineral Cache Colors") : .ItemData(.NewIndex) = elOptionSetItem.MineralCacheColor
                .AddItem("Minimap Angle Colors") : .ItemData(.NewIndex) = elOptionSetItem.MinimapViewAngle
            End With
            lstOptionSet.ListIndex = 0
            lstOptionSet_ItemClick(0)
            'ConfigureGeneral()
            mbIgnoreEvents = False
        End Sub

        Private mlSyncLockObj(-1) As Int32
        Private Sub lstOptionSet_ItemClick(ByVal lIndex As Integer) Handles lstOptionSet.ItemClick
            mbIgnoreEvents = True

            'SyncLock mlSyncLockObj

            '    Me.Enabled = False
            Me.RemoveAllChildren()
            Me.AddChild(CType(lstOptionSet, UIControl))

            If lIndex > -1 Then
                Select Case lstOptionSet.ItemData(lIndex)
                    Case elOptionSetItem.BuildGhostColor
                        ConfigureBuildGhostColors()
                    Case elOptionSetItem.ChatWindowColors
                        ConfigureChatWindowColors()
                    Case elOptionSetItem.Gameplay
                        ConfigureGameplay()
                    Case elOptionSetItem.General
                        ConfigureGeneral()
                    Case elOptionSetItem.IdentificationColors
                        ConfigureIdentificationColors()
                    Case elOptionSetItem.InterfaceColors
                        ConfigureInterfaceColors()
                    Case elOptionSetItem.Lighting
                        ConfigureLighting()
                    Case elOptionSetItem.MineralCacheColor
                        ConfigureMineralCacheColors()
                    Case elOptionSetItem.MinimapViewAngle
                        ConfigureMinimapAngleColors()
                    Case elOptionSetItem.Particles
                        ConfigureParticles()
                    Case elOptionSetItem.Shadowing
                    Case elOptionSetItem.Textures
                        ConfigureTexture()
                    Case elOptionSetItem.WaterOptions
                        ConfigureWater()
                End Select
            End If

            If moColorSelect Is Nothing = False Then
                Me.AddChild(CType(moColorSelect, UIControl))
            End If
            'Me.Enabled = True
            mbIgnoreEvents = False

            '    Me.IsDirty = True
            'End SyncLock
            Me.IsDirty = True
        End Sub

        Private Sub ConfigureCheckboxClick()
            If mbIgnoreEvents = True Then Return

            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X) Is Nothing = False AndAlso Me.moChildren(X).ControlName Is Nothing = False Then

                    Select Case Me.moChildren(X).ControlName.ToUpper
                        Case "CHKSHOWMINIMAP"
                            muSettings.ShowMiniMap = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKDRAWGRID"
                            muSettings.DrawGrid = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKRENDERCACHE"
                            muSettings.RenderMineralCaches = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKSHOWINTRO"
                            muSettings.ShowIntro = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKWINDOWED"
                            If muSettings.Windowed <> CType(Me.moChildren(X), UICheckBox).Value Then
                                MyBase.moUILib.AddNotification("You will need to restart Beyond Protocol in order for the changes to take effect.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                muSettings.Windowed = CType(Me.moChildren(X), UICheckBox).Value
                            End If
                            'GFXEngine.ToggleFullScreen()
                        Case "CHKTRIPLEBUFFER"
                            muSettings.SavedTripleBuffer = CType(Me.moChildren(X), UICheckBox).Value
                            If muSettings.TripleBuffer <> muSettings.SavedTripleBuffer Then MyBase.moUILib.AddNotification("You will need to restart Beyond Protocol in order for the changes to take effect.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            'GFXEngine.gbPaused = True
                            'GFXEngine.gbDeviceLost = True
                            'GFXEngine.gbPaused = False
                        Case "CHKVSYNC"
                            muSettings.SavedVSync = CType(Me.moChildren(X), UICheckBox).Value
                            If muSettings.SavedVSync <> muSettings.VSync Then MyBase.moUILib.AddNotification("You will need to restart Beyond Protocol in order for the changes to take effect.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            'GFXEngine.gbPaused = True
                            'GFXEngine.gbDeviceLost = True
                            'GFXEngine.gbPaused = False
                        Case "CHKCHATTIMESTAMPS"
                            muSettings.ChatTimeStamps = CType(Me.moChildren(X), UICheckBox).Value
                            Dim oChat As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)
                            If oChat Is Nothing = False AndAlso oChat.Visible = True Then
                                oChat.RefreshChatText()
                            End If
                            oChat = Nothing
                        Case "CHKSMOOTHFOW"
                            muSettings.SmoothFOW = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKHIRESPLANETS"
                            muSettings.HiResPlanetTexture = CType(Me.moChildren(X), UICheckBox).Value
                            muSettings.SaveSettings()
                            If NewTutorialManager.TutorialOn = True Then
                                Planet.GetTutorialPlanet().ClearPlanetTexture()
                            Else
                                If goGalaxy Is Nothing = False Then
                                    For A As Int32 = 0 To goGalaxy.mlSystemUB
                                        If goGalaxy.moSystems(A) Is Nothing = False Then
                                            For Y As Int32 = 0 To goGalaxy.moSystems(A).PlanetUB
                                                If goGalaxy.moSystems(A).moPlanets(Y) Is Nothing = False Then goGalaxy.moSystems(A).moPlanets(Y).ClearPlanetTexture()
                                            Next Y
                                        End If
                                    Next A
                                End If
                            End If
                        Case "CHKRENDERCOSMOS"
                            muSettings.RenderCosmos = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKRENDERPLANETRINGS"
                            muSettings.RenderPlanetRings = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKSHIELDFX"
                            muSettings.RenderShieldFX = CType(Me.moChildren(X), UICheckBox).Value
                            If muSettings.RenderShieldFX = False Then goShldMgr.TurnOffMgr()
                        Case "CHKSPECULARMAPS"
                            muSettings.RenderSpecularMap = CType(Me.moChildren(X), UICheckBox).Value
                        Case "CHKSHOWTARGETBOXES"
                            muSettings.ShowTargetBoxes = CType(moChildren(X), UICheckBox).Value
                        Case "CHKBUMPTERRAIN"
                            muSettings.BumpMapTerrain = CType(moChildren(X), UICheckBox).Value
                        Case "CHKBUMPPLANETMODELS"
                            muSettings.BumpMapPlanetModel = CType(moChildren(X), UICheckBox).Value
                        Case "CHKILLUMTERRAIN"
                            muSettings.IlluminationMapTerrain = CType(moChildren(X), UICheckBox).Value
                        Case "CHKFILTERLANGUAGE"
                            muSettings.FilterBadWords = CType(moChildren(X), UICheckBox).Value
                        Case "CHKSHOWHPBARS"
                            muSettings.RenderHPBars = CType(moChildren(X), UICheckBox).Value
                        Case "CHKWORMHOLEAURA"
                            muSettings.WormholeAura = CType(moChildren(X), UICheckBox).Value
                        Case "CHKPULSEBOLTS"
                            muSettings.RenderPulseBolts = CType(moChildren(X), UICheckBox).Value
                        Case "CHKEXPLOSIONS"
                            muSettings.RenderExplosionFX = CType(moChildren(X), UICheckBox).Value
                        Case "CHKFLASHSELECTION"
                            muSettings.FlashSelections = CType(moChildren(X), UICheckBox).Value
                        Case "CHKDISABLEEXCLIMATIONS"
                            muSettings.gbDisableExclamations = CType(moChildren(X), UICheckBox).Value
                        Case "CHKDYK"
                            If CType(moChildren(X), UICheckBox).Value = True Then
                                For Y As Int32 = 0 To ChildrenUB
                                    If moChildren(Y) Is Nothing = False AndAlso moChildren(Y).ControlName Is Nothing = False Then
                                        If moChildren(Y).ControlName.ToUpper = "HSCRDYKFREQ" Then
                                            muSettings.lDYK_Frequency = CType(moChildren(Y), UIScrollBar).Value + 1
                                            Exit For
                                        End If
                                    End If
                                Next Y
                            Else
                                muSettings.lDYK_Frequency = -1
                            End If
                        Case "CHKSHOWVIEWMESSAGES"
                            muSettings.bShowViewMessages = CType(moChildren(X), UICheckBox).Value
                        Case "CHKSHOWCONNECTIONSTATUS"
                            muSettings.ShowConnectionStatus = CType(moChildren(X), UICheckBox).Value
                            Dim ofrmCS As frmConnectionStatus = CType(goUILib.GetWindow("frmConnectionStatus"), frmConnectionStatus)
                            If muSettings.ShowConnectionStatus = False Then
                                If ofrmCS Is Nothing = False Then
                                    MyBase.moUILib.RemoveWindow("frmConnectionStatus")
                                End If
                            ElseIf muSettings.ShowConnectionStatus = True Then
                                If ofrmCS Is Nothing = True Then ofrmCS = New frmConnectionStatus(goUILib)
                                ofrmCS.Visible = True
                                ofrmCS = Nothing
                            End If
                        Case "CHKSHOWMULTIHEALTHBARS"
                            muSettings.ShowMultiHealthBars = CType(moChildren(X), UICheckBox).Value
                    End Select

                End If
            Next X
        End Sub

        Private Sub ConfigureComboboxItemSelected(ByVal lIndex As Int32)
            If mbIgnoreEvents = True Then Return

            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X) Is Nothing = False AndAlso Me.moChildren(X).ControlName Is Nothing = False Then

                    Select Case Me.moChildren(X).ControlName.ToUpper
                        Case "CBOVERTEX"
                            With CType(moChildren(X), UIComboBox)
                                If .ListIndex <> -1 Then
                                    If muSettings.VertexProcessing <> .ItemData(.ListIndex) Then
                                        muSettings.VertexProcessing = .ItemData(.ListIndex)
                                        muSettings.SaveSettings()
                                        GFXEngine.gbDeviceLost = True
                                    End If
                                End If
                            End With
                        Case "CBOFULLSCREENRESOLUTION"
                            With CType(moChildren(X), UIComboBox)
                                If .ListIndex > -1 Then
                                    Dim sItem As String = .List(.ListIndex)
                                    'uDispMode.Width & " x " & uDispMode.Height & " at " & uDispMode.RefreshRate.ToString & " Hz"

                                    sItem = sItem.ToUpper.Replace(" ", "").Replace("AT", "X").Replace("HZ", "X")
                                    Dim sVals() As String = Split(sItem, "X")
                                    If sVals Is Nothing = False AndAlso sVals.GetUpperBound(0) >= 1 Then
                                        muSettings.FullScreenResX = CInt(sVals(0))
                                        muSettings.FullScreenResY = CInt(sVals(1))
                                        muSettings.FullScreenRefreshRate = .ItemData2(.ListIndex)
                                        muSettings.SaveSettings()
                                    End If
                                End If
                            End With
                        Case "CBOFOWRES"
                            Dim cboFOWRes As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboFOWRes.ListIndex > -1 Then
                                Dim lVal As Int32 = cboFOWRes.ItemData(cboFOWRes.ListIndex)
                                Dim bResetVisTex As Boolean = False
                                If lVal = 0 Then
                                    If muSettings.ShowFOWTerrainShading = True Then
                                        bResetVisTex = True
                                        muSettings.ShowFOWTerrainShading = False
                                    End If
                                Else
                                    If muSettings.ShowFOWTerrainShading = False OrElse muSettings.FOWTextureResolution <> lVal Then
                                        muSettings.ShowFOWTerrainShading = True
                                        muSettings.FOWTextureResolution = lVal
                                        bResetVisTex = True
                                    End If
                                End If
                                'This tells the engine to recreate the FOW texture
                                If bResetVisTex = True Then TerrainClass.ReleaseVisibleTexture()
                            End If
                        Case "CBOTERRAINTEXRES"
                            Dim cboTerrainTexRes As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboTerrainTexRes.ListIndex > -1 Then
                                Dim lVal As Int32 = cboTerrainTexRes.ItemData(cboTerrainTexRes.ListIndex)
                                If lVal <> muSettings.TerrainTextureResolution Then
                                    frmMain.Cursor = Cursors.WaitCursor
                                    muSettings.TerrainTextureResolution = lVal
                                    TerrainClass.bReloadVertexBuffer = True
                                    frmMain.Cursor = Cursors.Arrow
                                End If
                            End If
                        Case "CBORESOLUTION"
                            Dim cboTextureRes As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboTextureRes.ListIndex > -1 Then
                                Dim lVal As Int32 = cboTextureRes.ItemData(cboTextureRes.ListIndex)
                                If lVal <> muSettings.ModelTextureResolution Then
                                    frmMain.Cursor = Cursors.WaitCursor
                                    muSettings.ModelTextureResolution = CType(lVal, EngineSettings.eTextureResOptions)
                                    goResMgr.UpdateModelTextureResolution()
                                    frmMain.Cursor = Cursors.Arrow
                                End If
                            End If
                        Case "CBOPLANETFX"
                            Dim cboPlanetFX As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboPlanetFX.ListIndex > -1 Then
                                muSettings.PlanetFXParticles = cboPlanetFX.ItemData(cboPlanetFX.ListIndex)
                            End If
                        Case "CBOBURNFX"
                            Dim cboBurnFX As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboBurnFX.ListIndex > -1 Then
                                Dim lVal As Int32 = cboBurnFX.ItemData(cboBurnFX.ListIndex)
                                If lVal <> muSettings.BurnFXParticles Then
                                    muSettings.BurnFXParticles = lVal
                                    goPFXEngine32.ResetBurnFXPCnts()
                                End If
                            End If
                        Case "CBOENGINEFX"
                            Dim cboEngineFX As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboEngineFX.ListIndex > -1 Then
                                Dim lVal As Int32 = cboEngineFX.ItemData(cboEngineFX.ListIndex)
                                If lVal <> muSettings.EngineFXParticles Then
                                    muSettings.EngineFXParticles = lVal
                                    goPFXEngine32.ResetEngineFXPCnts()
                                End If
                            End If
                        Case "CBOILLUMMAP"
                            Dim cboIllumMap As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboIllumMap.ListIndex > -1 Then
                                Dim lVal As Int32 = cboIllumMap.ItemData(cboIllumMap.ListIndex)
                                If lVal <> muSettings.IlluminationMap Then
                                    frmMain.Cursor = Cursors.WaitCursor
                                    muSettings.IlluminationMap = lVal
                                    goResMgr.UpdateIllumTextureResolution()
                                    frmMain.Cursor = Cursors.Arrow
                                End If
                            End If
                        Case "CBOQUALITY"
                            Dim cboQuality As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboQuality.ListIndex > -1 Then
                                Dim lVal As Int32 = cboQuality.ItemData(cboQuality.ListIndex)
                                If lVal <> muSettings.LightQuality Then
                                    muSettings.LightQuality = CType(lVal, EngineSettings.LightQualitySetting)
                                End If
                            End If
                        Case "CBOTECHNIQUE"
                            Dim cboTechnique As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboTechnique.ListIndex > -1 Then
                                Dim lVal As Int32 = cboTechnique.ItemData(cboTechnique.ListIndex)
                                If lVal <> muSettings.WaterRenderMethod Then
                                    frmMain.Cursor = Cursors.WaitCursor
                                    muSettings.WaterRenderMethod = lVal
                                    TerrainClass.ForceWaterCreation()
                                    frmMain.Cursor = Cursors.Arrow
                                End If
                            End If
                        Case "CBOWATERRESOLUTION"
                            Dim cboWaterRes As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboWaterRes.ListIndex > -1 Then
                                Dim lVal As Int32 = cboWaterRes.ItemData(cboWaterRes.ListIndex)
                                If lVal <> muSettings.WaterTextureRes Then
                                    frmMain.Cursor = Cursors.WaitCursor
                                    muSettings.WaterTextureRes = CType(lVal, EngineSettings.eTextureResOptions)
                                    goResMgr.UpdateWaterTextureResolution()
                                    frmMain.Cursor = Cursors.Arrow
                                End If
                            End If
                        Case "CBOSCREENSHOTFORMAT"
                            Dim cboScreenshotFormat As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboScreenshotFormat.ListIndex > -1 Then
                                Dim lVal As Int32 = cboScreenshotFormat.ItemData(cboScreenshotFormat.ListIndex)
                                If lVal <> muSettings.ScreenshotFormat Then
                                    muSettings.ScreenshotFormat = lVal
                                End If
                            End If
                        Case "CBOEXPORTEDDATAFORMAT"
                            Dim cboExportedDataFormat As UIComboBox = CType(moChildren(X), UIComboBox)
                            If cboExportedDataFormat.ListIndex > -1 Then
                                Dim lVal As Int32 = cboExportedDataFormat.ItemData(cboExportedDataFormat.ListIndex)
                                If lVal <> muSettings.ExportedDataFormat Then
                                    muSettings.ExportedDataFormat = lVal
                                End If
                            End If
                    End Select
                End If
            Next X
        End Sub

        Private Sub ConfigureScrollBarScroll()
            If mbIgnoreEvents = True Then Return

            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X) Is Nothing = False AndAlso Me.moChildren(X).ControlName Is Nothing = False Then

                    Select Case Me.moChildren(X).ControlName.ToUpper
                        Case "HSCRENTITYCLIPPLANE"
                            muSettings.EntityClipPlane = (CType(moChildren(X), UIScrollBar).Value * 100)
                        Case "HSCRSPACEDENS"
                            muSettings.StarfieldParticlesSpace = (CType(moChildren(X), UIScrollBar).Value * 100)
                        Case "HSCRPLANETDENS"
                            muSettings.StarfieldParticlesPlanet = (CType(moChildren(X), UIScrollBar).Value * 10)
                        Case "HSCRGLOWFX"
                            muSettings.PostGlowAmt = CType(moChildren(X), UIScrollBar).Value / 10.0F
                        Case "HSCRAMBIENT"
                            muSettings.AmbientLevel = CType(moChildren(X), UIScrollBar).Value
                        Case "HSCRFLASHRATE"
                            muSettings.FlashRate = CType(moChildren(X), UIScrollBar).Value + 1
                        Case "HSCRDYKFREQ"
                            If muSettings.lDYK_Frequency <> -1 Then
                                muSettings.lDYK_Frequency = CType(moChildren(X), UIScrollBar).Value + 1
                            End If
                    End Select

                End If
            Next X
        End Sub

        Private Sub ResetInterfaceDefaults(ByVal sName As String)
            Dim clrPrev As System.Drawing.Color = muSettings.InterfaceBorderColor
            muSettings.InterfaceBorderColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            MyBase.moUILib.DirtyAllWindows(1, clrPrev)  'border

            clrPrev = muSettings.InterfaceFillColor
            muSettings.InterfaceFillColor = System.Drawing.Color.FromArgb(128, 32, 64, 128)
            MyBase.moUILib.DirtyAllWindows(2, clrPrev)  'fill

            clrPrev = muSettings.InterfaceTextBoxForeColor
            muSettings.InterfaceTextBoxForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            MyBase.moUILib.DirtyAllWindows(3, clrPrev)  'textfore

            clrPrev = muSettings.InterfaceTextBoxFillColor
            muSettings.InterfaceTextBoxFillColor = System.Drawing.Color.FromArgb(255, 32, 64, 92)
            MyBase.moUILib.DirtyAllWindows(4, clrPrev)  'textfill

            clrPrev = muSettings.InterfaceButtonColor
            muSettings.InterfaceButtonColor = System.Drawing.Color.FromArgb(255, 192, 220, 255)
            MyBase.moUILib.DirtyAllWindows(5, clrPrev)  'button

            ConfigureInterfaceColors()

        End Sub

#Region "  Color Configuration Routines  "
        Private Sub ConfigureColorButtonClick(ByVal sName As String)
            If moColorSelect Is Nothing Then
                moColorSelect = New fraColorSelect(MyBase.moUILib)
                Me.AddChild(CType(moColorSelect, UIControl))
                AddHandler moColorSelect.ColorSelectClosed, AddressOf moColorSelect_ColorSelectClosed
            End If
            moColorSelect.Visible = True
            moColorSelect.SetShowAlpha(False)
            moColorSelect.SetAlphaMin(0)

            Select Case sName.ToUpper
                Case "BTNSETBORDER"
                    moColorSelect.SetShowAlpha(True)
                    moColorSelect.SetAlphaMin(32)
                    moColorSelect.SetColor(muSettings.InterfaceBorderColor, sName)
                Case "BTNSETFILL"
                    moColorSelect.SetShowAlpha(True)
                    moColorSelect.SetColor(muSettings.InterfaceFillColor, sName & "INVALIDUSE")
                Case "BTNSETTEXTBOXFORE"
                    moColorSelect.SetShowAlpha(True)
                    moColorSelect.SetColor(muSettings.InterfaceTextBoxForeColor, sName)
                Case "BTNSETTEXTBOXBACK"
                    moColorSelect.SetShowAlpha(True)
                    moColorSelect.SetColor(muSettings.InterfaceTextBoxFillColor, sName)
                Case "BTNSETBUTTONCOLOR"
                    moColorSelect.SetShowAlpha(True)
                    moColorSelect.SetColor(muSettings.InterfaceButtonColor, sName)
                Case "BTNSETDEFAULT"
                    moColorSelect.SetColor(muSettings.DefaultChatColor, sName & "COLOR")
                Case "BTNSETERROR"
                    moColorSelect.SetColor(muSettings.AlertChatColor, "LBLALERTCOLOR")
                Case "BTNSETSTATUS"
                    moColorSelect.SetColor(muSettings.StatusChatColor, sName & "COLOR")
                Case "BTNSETLOCAL"
                    moColorSelect.SetColor(muSettings.LocalChatColor, sName & "COLOR")
                Case "BTNSETGUILD"
                    moColorSelect.SetColor(muSettings.GuildChatColor, sName & "COLOR")
                Case "BTNSETSENATE"
                    moColorSelect.SetColor(muSettings.SenateChatColor, sName & "COLOR")
                Case "BTNSETPM"
                    moColorSelect.SetColor(muSettings.PMChatColor, sName & "COLOR")
                Case "BTNSETCHANNEL"
                    moColorSelect.SetColor(muSettings.ChannelChatColor, sName & "COLOR")
                Case "BTNSETNEUTRAL"
                    moColorSelect.SetColor(ConvertVector4ToColor(muSettings.NeutralAssetColor), sName & "COLOR")
                Case "BTNSETMYCOLOR"
                    moColorSelect.SetColor(ConvertVector4ToColor(muSettings.MyAssetColor), sName)
                Case "BTNSETENEMY"
                    moColorSelect.SetColor(ConvertVector4ToColor(muSettings.EnemyAssetColor), sName & "COLOR")
                Case "BTNSETALLY"
                    moColorSelect.SetColor(ConvertVector4ToColor(muSettings.AllyAssetColor), sName & "COLOR")
                Case "BTNSETGUILDCOLOR"
                    moColorSelect.SetColor(ConvertVector4ToColor(muSettings.GuildAssetColor), sName & "ID")
                Case "BTNSETTACTICAL"
                    moColorSelect.SetColor(ConvertVector4ToColor(muSettings.TacticalAssetColor), sName)
                Case "BTNSETBUILDGHOSTACID"
                    moColorSelect.SetColor(muSettings.AcidBuildGhost, sName)
                Case "BTNSETBUILDGHOSTADAPTABLE"
                    moColorSelect.SetColor(muSettings.AdaptableBuildGhost, sName)
                Case "BTNSETBUILDGHOSTBARREN"
                    moColorSelect.SetColor(muSettings.BarrenBuildGhost, sName)
                Case "BTNSETBUILDGHOSTDESERT"
                    moColorSelect.SetColor(muSettings.DesertBuildGhost, sName)
                Case "BTNSETBUILDGHOSTICE"
                    moColorSelect.SetColor(muSettings.IceBuildGhost, sName)
                Case "BTNSETBUILDGHOSTLAVA"
                    moColorSelect.SetColor(muSettings.LavaBuildGhost, sName)
                Case "BTNSETBUILDGHOSTTERRAN"
                    moColorSelect.SetColor(muSettings.TerranBuildGhost, sName)
                Case "BTNSETBUILDGHOSTWATERWORLD"
                    moColorSelect.SetColor(muSettings.WaterworldBuildGhost, sName)
                Case "BTNSETMINERALCACHEACID"
                    moColorSelect.SetColor(muSettings.AcidMineralCache, sName)
                Case "BTNSETMINERALCACHEADAPTABLE"
                    moColorSelect.SetColor(muSettings.AdaptableMineralCache, sName)
                Case "BTNSETMINERALCACHEBARREN"
                    moColorSelect.SetColor(muSettings.BarrenMineralCache, sName)
                Case "BTNSETMINERALCACHEDESERT"
                    moColorSelect.SetColor(muSettings.DesertMineralCache, sName)
                Case "BTNSETMINERALCACHEICE"
                    moColorSelect.SetColor(muSettings.IceMineralCache, sName)
                Case "BTNSETMINERALCACHELAVA"
                    moColorSelect.SetColor(muSettings.LavaMineralCache, sName)
                Case "BTNSETMINERALCACHETERRAN"
                    moColorSelect.SetColor(muSettings.TerranMineralCache, sName)
                Case "BTNSETMINERALCACHEWATERWORLD"
                    moColorSelect.SetColor(muSettings.WaterworldMineralCache, sName)
                Case "BTNSETMINIMAPANGLEACID"
                    moColorSelect.SetColor(muSettings.AcidMinimapAngle, sName)
                Case "BTNSETMINIMAPANGLEADAPTABLE"
                    moColorSelect.SetColor(muSettings.AdaptableMinimapAngle, sName)
                Case "BTNSETMINIMAPANGLEBARREN"
                    moColorSelect.SetColor(muSettings.BarrenMinimapAngle, sName)
                Case "BTNSETMINIMAPANGLEDESERT"
                    moColorSelect.SetColor(muSettings.DesertMinimapAngle, sName)
                Case "BTNSETMINIMAPANGLEICE"
                    moColorSelect.SetColor(muSettings.IceMinimapAngle, sName)
                Case "BTNSETMINIMAPANGLELAVA"
                    moColorSelect.SetColor(muSettings.LavaMinimapAngle, sName)
                Case "BTNSETMINIMAPANGLETERRAN"
                    moColorSelect.SetColor(muSettings.TerranMinimapAngle, sName)
                Case "BTNSETMINIMAPANGLEWATERWORLD"
                    moColorSelect.SetColor(muSettings.WaterworldMinimapAngle, sName)
                Case Else
                    Return
            End Select

            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X) Is Nothing = False Then
                    If Me.moChildren(X).ControlName Is Nothing = False Then
                        If Me.moChildren(X).ControlName.ToUpper <> "FRACOLORSELECT" Then
                            Me.moChildren(X).Enabled = False
                        End If
                    Else
                        Me.moChildren(X).Enabled = False
                    End If
                End If
            Next X

            Me.IsDirty = True

        End Sub
        Private Function ConvertVector4ToColor(ByVal vec4 As Vector4) As System.Drawing.Color
            Return System.Drawing.Color.FromArgb(255, CInt(vec4.X * 255), CInt(vec4.Y * 255), CInt(vec4.Z * 255))
        End Function
        Private Function ConvertColorToVector4(ByVal clrVal As System.Drawing.Color) As Vector4
            Return New Vector4(clrVal.R / 255.0F, clrVal.G / 255.0F, clrVal.B / 255.0F, clrVal.A / 255.0F)
        End Function
        Private Sub moColorSelect_ColorSelectClosed(ByVal clrValue As System.Drawing.Color, ByVal sKeyName As String)
            Dim sLblName As String = sKeyName.ToUpper.Replace("BTNSET", "LBL") '& "COLOR"
            For X As Int32 = 0 To Me.ChildrenUB
                If Me.moChildren(X) Is Nothing = False Then
                    Me.moChildren(X).Enabled = True
                    If Me.moChildren(X).ControlName Is Nothing = False AndAlso moChildren(X).ControlName.ToUpper = sLblName Then
                        CType(Me.moChildren(X), UILabel).ForeColor = clrValue
                    End If
                End If
            Next X
            moColorSelect.Visible = False

            Select Case sKeyName.ToUpper
                Case "BTNSETBORDER"
                    Dim clrPrev As System.Drawing.Color = muSettings.InterfaceBorderColor
                    muSettings.InterfaceBorderColor = clrValue
                    MyBase.moUILib.DirtyAllWindows(1, clrPrev)
                Case "BTNSETFILLINVALIDUSE"
                    Dim clrPrev As System.Drawing.Color = muSettings.InterfaceFillColor
                    muSettings.InterfaceFillColor = clrValue
                    MyBase.moUILib.DirtyAllWindows(2, clrPrev)
                Case "BTNSETTEXTBOXFORE"
                    Dim clrPrev As System.Drawing.Color = muSettings.InterfaceTextBoxForeColor
                    muSettings.InterfaceTextBoxForeColor = clrValue
                    MyBase.moUILib.DirtyAllWindows(3, clrPrev)
                Case "BTNSETTEXTBOXBACK"
                    Dim clrPrev As System.Drawing.Color = muSettings.InterfaceTextBoxFillColor
                    muSettings.InterfaceTextBoxFillColor = clrValue
                    MyBase.moUILib.DirtyAllWindows(4, clrPrev)
                Case "BTNSETBUTTONCOLOR"
                    Dim clrPrev As System.Drawing.Color = muSettings.InterfaceButtonColor
                    muSettings.InterfaceButtonColor = clrValue
                    MyBase.moUILib.DirtyAllWindows(5, clrPrev)
                Case "BTNSETDEFAULTCOLOR"
                    muSettings.DefaultChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "LBLALERTCOLOR"
                    muSettings.AlertChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "BTNSETSTATUSCOLOR"
                    muSettings.StatusChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "BTNSETLOCALCOLOR"
                    muSettings.LocalChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "BTNSETGUILDCOLOR"
                    muSettings.GuildChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "BTNSETSENATECOLOR"
                    muSettings.SenateChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "BTNSETPMCOLOR"
                    muSettings.PMChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "BTNSETCHANNELCOLOR"
                    muSettings.ChannelChatColor = clrValue
                    MyBase.moUILib.DirtyChatWindow()
                Case "BTNSETNEUTRALCOLOR"
                    muSettings.NeutralAssetColor = ConvertColorToVector4(clrValue)
                Case "BTNSETMYCOLOR"
                    muSettings.MyAssetColor = ConvertColorToVector4(clrValue)
                Case "BTNSETENEMYCOLOR"
                    muSettings.EnemyAssetColor = ConvertColorToVector4(clrValue)
                Case "BTNSETALLYCOLOR"
                    muSettings.AllyAssetColor = ConvertColorToVector4(clrValue)
                Case "BTNSETGUILDCOLORID"
                    muSettings.GuildAssetColor = ConvertColorToVector4(clrValue)
                Case "BTNSETTACTICAL"
                    muSettings.TacticalAssetColor = ConvertColorToVector4(clrValue)
                Case "BTNSETBUILDGHOSTACID"
                    muSettings.AcidBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETBUILDGHOSTADAPTABLE"
                    muSettings.AdaptableBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETBUILDGHOSTBARREN"
                    muSettings.BarrenBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETBUILDGHOSTDESERT"
                    muSettings.DesertBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETBUILDGHOSTICE"
                    muSettings.IceBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETBUILDGHOSTLAVA"
                    muSettings.LavaBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETBUILDGHOSTTERRAN"
                    muSettings.TerranBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETBUILDGHOSTWATERWORLD"
                    muSettings.WaterworldBuildGhost = ConvertToBuildGhostColor(clrValue)
                Case "BTNSETMINERALCACHEACID"
                    muSettings.AcidMineralCache = clrValue
                Case "BTNSETMINERALCACHEADAPTABLE"
                    muSettings.AdaptableMineralCache = clrValue
                Case "BTNSETMINERALCACHEBARREN"
                    muSettings.BarrenMineralCache = clrValue
                Case "BTNSETMINERALCACHEDESERT"
                    muSettings.DesertMineralCache = clrValue
                Case "BTNSETMINERALCACHEICE"
                    muSettings.IceMineralCache = clrValue
                Case "BTNSETMINERALCACHELAVA"
                    muSettings.LavaMineralCache = clrValue
                Case "BTNSETMINERALCACHETERRAN"
                    muSettings.TerranMineralCache = clrValue
                Case "BTNSETMINERALCACHEWATERWORLD"
                    muSettings.WaterworldMineralCache = clrValue
                Case "BTNSETMINIMAPANGLEACID"
                    muSettings.AcidMinimapAngle = clrValue
                Case "BTNSETMINIMAPANGLEADAPTABLE"
                    muSettings.AdaptableMinimapAngle = clrValue
                Case "BTNSETMINIMAPANGLEBARREN"
                    muSettings.BarrenMinimapAngle = clrValue
                Case "BTNSETMINIMAPANGLEDESERT"
                    muSettings.DesertMinimapAngle = clrValue
                Case "BTNSETMINIMAPANGLEICE"
                    muSettings.IceMinimapAngle = clrValue
                Case "BTNSETMINIMAPANGLELAVA"
                    muSettings.LavaMinimapAngle = clrValue
                Case "BTNSETMINIMAPANGLETERRAN"
                    muSettings.TerranMinimapAngle = clrValue
                Case "BTNSETMINIMAPANGLEWATERWORLD"
                    muSettings.WaterworldMinimapAngle = clrValue
                Case Else
                    Return
            End Select


        End Sub
#End Region

#Region "  Gameplay Config  "
        Private Sub ConfigureGameplay()
            Dim chkBox As New UICheckBox(MyBase.moUILib)
            With chkBox
                'chkShowMinimap initial props
                .ControlName = "chkShowMinimap"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 2
                .Width = 122
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Show Mini Map"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.ShowMiniMap
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkDrawGrid"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 20
                .Width = 152
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Draw Grid in Space"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.DrawGrid
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkRenderCache"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 40
                .Width = 181
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Mineral Caches"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.RenderMineralCaches
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkShowIntro"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 60
                .Width = 90
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Show Intro"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.ShowIntro
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkShowTargetBoxes"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 80
                .Width = 152
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Show Target Boxes"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.ShowTargetBoxes
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, selected items will highlight their targets with colored boxes based on arc."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkFilterLanguage"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 100
                .Width = 135
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Filter Bad Words"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.FilterBadWords
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, text is filtered for general bad words."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkShowHPBars"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 120
                .Width = 235
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Show Hitpoint Bars Above Units"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.RenderHPBars
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, hitpoint bars are rendered above units." & vbCrLf & "Some machines may not handle this option correctly and may cause crashes."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkFlashSelection"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 140
                .Width = 228
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Flash Selected Units/Facilities"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.FlashSelections
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, units and facilities will flash while selected."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            Dim lblItem As UILabel = Nothing
            lblItem = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblFlashRate"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 160
                .Width = 80
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Flash Rate:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblItem, UIControl))

            Dim hscrBar As New UIScrollBar(MyBase.moUILib, False)
            With hscrBar
                .ControlName = "hscrFlashRate"
                .Left = lblItem.Left + lblItem.Width + 15
                .Top = lstOptionSet.Top + 160
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .MaxValue = 10
                .MinValue = 0
                .Value = muSettings.FlashRate
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrBar, UIControl))
            AddHandler hscrBar.ValueChange, AddressOf ConfigureScrollBarScroll


            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkDYK"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 180
                .Width = 228
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Show Did You Know Messages"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.lDYK_Frequency > 0
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, Did You Know messages will appear when changing environments."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            lblItem = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblDYKFreq"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 200
                .Width = 110
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "DYK Frequency:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "How often Did You Know messages appear."
            End With
            Me.AddChild(CType(lblItem, UIControl))

            hscrBar = New UIScrollBar(MyBase.moUILib, False)
            With hscrBar
                .ControlName = "hscrDYKFreq"
                .Left = lblItem.Left + lblItem.Width + 15
                .Top = lstOptionSet.Top + 200
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .MaxValue = 10
                .MinValue = 0
                .Value = Math.Max(0, muSettings.lDYK_Frequency - 1)
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = False
            End With
            Me.AddChild(CType(hscrBar, UIControl))

            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkShowViewMessages"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 220
                .Width = 225
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Show View Change Messages"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.bShowViewMessages
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, the name of the view will appear when you change views."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkShowMultiHealthBars"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 240
                .Width = 220
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Show MultiSelect Health Bars"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.ShowMultiHealthBars
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, multi-unit display will show progress-bar like unit health indicators."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkShowConnectionStatus"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 260
                .Width = 241
                .Height = 18
                .Enabled = True
                .Visible = isAdmin()
                .Caption = "Show Connection Status Window"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.ShowConnectionStatus
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, a connection status window will display above the chat bar."
            End With
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick
            Me.AddChild(CType(chkBox, UIControl))

            AddHandler hscrBar.ValueChange, AddressOf ConfigureScrollBarScroll

        End Sub
#End Region
#Region "  General Config  "
        Private Sub ConfigureGeneral()
            Dim chkBox As New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkWindowed"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 5
                .Width = 92
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Windowed"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.Windowed
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck

                'REMOVE WHEN FULLSCREEN MAKES IT IN
                '.ToolTipText = "Fullscreen mode has been disabled."
                '.Enabled = False
                '.Value = False
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkTripleBuffer"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 105
                .Width = 105 '0
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Triple Buffer"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.TripleBuffer
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Triple Buffering uses additional video memory to enhance performance." & vbCrLf & "Especially useful when Vertical Sync is turned on as performance can be greatly improved."
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkVSync"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 130
                .Width = 110 '0
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Vertical Sync"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.VSync
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Vertical Synchronization causes the video card to wait for the" & vbCrLf & "monitor to finish painting the screen before telling the monitor to paint a new screen." & vbCrLf & "This removes artifacts known as tearing which is one frame's image" & vbCrLf & "partially showing up with another frame's image."
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkChatTimeStamps"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 155
                .Width = 196 '0
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Chat Window Timestamps"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.ChatTimeStamps
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Show Timestamps in the chat window."
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            Dim lCboLeft As Int32

            Dim lblItem As UILabel = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblRes"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 30
                .Width = 157
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Fullscreen Resolution:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblItem, UIControl))

            lCboLeft = lblItem.Left + lblItem.Width + 5

            lblItem = Nothing
            lblItem = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblVertex"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 55
                .Width = 157
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Vertex Processing:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblItem, UIControl))

            lblItem = Nothing
            lblItem = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblEntityClipPlane"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 80
                .Width = 157
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Entity Clip Plane:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblItem, UIControl))

            Dim hscrBar As New UIScrollBar(MyBase.moUILib, False)
            With hscrBar
                .ControlName = "hscrEntityClipPlane"
                .Left = lblItem.Left + lblItem.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .MaxValue = 400
                .MinValue = 100
                .Value = muSettings.EntityClipPlane \ 100
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(hscrBar, UIControl))
            AddHandler hscrBar.ValueChange, AddressOf ConfigureScrollBarScroll

            Dim cboBox As UIComboBox
            cboBox = Nothing
            cboBox = New UIComboBox(MyBase.moUILib)
            With cboBox
                .ControlName = "cboVertex"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 55
                .Width = 155
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
            cboBox.Clear()
            cboBox.AddItem("Hardware")
            cboBox.ItemData(cboBox.NewIndex) = CreateFlags.HardwareVertexProcessing
            cboBox.AddItem("Mixed")
            cboBox.ItemData(cboBox.NewIndex) = CreateFlags.MixedVertexProcessing
            cboBox.AddItem("Software")
            cboBox.ItemData(cboBox.NewIndex) = CreateFlags.SoftwareVertexProcessing
            If MyBase.moUILib.oDevice.CreationParameters.Behavior.HardwareVertexProcessing = True Then
                cboBox.FindComboItemData(CreateFlags.HardwareVertexProcessing)
            ElseIf MyBase.moUILib.oDevice.CreationParameters.Behavior.SoftwareVertexProcessing = True Then
                cboBox.FindComboItemData(CreateFlags.SoftwareVertexProcessing)
            Else : cboBox.FindComboItemData(CreateFlags.MixedVertexProcessing)
            End If
            Me.AddChild(CType(cboBox, UIControl))
            AddHandler cboBox.ItemSelected, AddressOf ConfigureComboboxItemSelected

            cboBox = Nothing
            cboBox = New UIComboBox(MyBase.moUILib)
            With cboBox
                .ControlName = "cboFullScreenResolution"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 30
                .Width = 155
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

                'REMOVE WHEN FULLSCREEN MODE MAKES IT IN
                '.Enabled = False
                '.ToolTipText = "Fullscreen mode has been disabled."
            End With
            Me.AddChild(CType(cboBox, UIControl))

            For Each ai As AdapterInformation In Manager.Adapters
                If ai.Adapter = CType(Manager.Adapters.Default, AdapterInformation).Adapter Then
                    For Each uDispMode As DisplayMode In ai.SupportedDisplayModes
                        If uDispMode.Width >= 1024 AndAlso uDispMode.Height >= 768 Then

                            Dim lBPP As Int32 = GetBPP(uDispMode.Format)
                            If lBPP <> 32 Then Continue For

                            'ok, get our Resolution value...
                            Dim lResValue As Int32 = (uDispMode.Width * &H10000) Or uDispMode.Height
                            Dim lBPPValue As Int32 = uDispMode.RefreshRate

                            Dim bFound As Boolean = False
                            For X As Int32 = 0 To cboBox.ListCount - 1
                                If cboBox.ItemData(X) = lResValue AndAlso cboBox.ItemData2(X) = lBPPValue Then
                                    bFound = True
                                    Exit For
                                End If
                            Next X
                            If bFound = True Then Continue For

                            cboBox.AddItem(uDispMode.Width & " x " & uDispMode.Height & " at " & uDispMode.RefreshRate.ToString & " Hz")
                            cboBox.ItemData(cboBox.NewIndex) = lResValue
                            cboBox.ItemData2(cboBox.NewIndex) = lBPPValue
                        End If
                    Next
                End If
            Next ai

            'Now, find our itemdata stuff...
            If muSettings.FullScreenResX = -1 Then
                muSettings.FullScreenResX = Manager.Adapters.Default.CurrentDisplayMode.Width
            End If
            If muSettings.FullScreenResY = -1 Then
                muSettings.FullScreenResY = Manager.Adapters.Default.CurrentDisplayMode.Height
            End If
            If muSettings.FullScreenRefreshRate = -1 Then
                muSettings.FullScreenRefreshRate = Manager.Adapters.Default.CurrentDisplayMode.RefreshRate
            End If
            Dim lCurrResVal As Int32 = (muSettings.FullScreenResX * &H10000) Or muSettings.FullScreenResY
            Dim lCurrBPPVal As Int32 = muSettings.FullScreenRefreshRate
            For X As Int32 = 0 To cboBox.ListCount - 1
                If cboBox.ItemData(X) = lCurrResVal AndAlso cboBox.ItemData2(X) = lCurrBPPVal Then
                    cboBox.ListIndex = X
                    Exit For
                End If
            Next X
            'and THEN ad the event so we dont fire it prematurely
            AddHandler cboBox.ItemSelected, AddressOf ConfigureComboboxItemSelected

            ''Exported Data Format
            'lblItem = Nothing
            'lblItem = New UILabel(MyBase.moUILib)
            'With lblItem
            '    .ControlName = "lblExportedDataFormat"
            '    .Left = lstOptionSet.Left + lstOptionSet.Width + 10
            '    .Top = lstOptionSet.Top + 180
            '    .Width = 157
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Exported Data Format:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblItem, UIControl))

            'cboBox = Nothing
            'cboBox = New UIComboBox(MyBase.moUILib)
            'With cboBox
            '    .ControlName = "cboExportedDataFormat"
            '    .Left = lCboLeft
            '    .Top = lstOptionSet.Top + 180
            '    .Width = 155
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .FillColor = muSettings.InterfaceTextBoxFillColor
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            '    .DropDownListBorderColor = muSettings.InterfaceBorderColor
            '    .mbAcceptReprocessEvents = True
            'End With
            'cboBox.Clear()
            'cboBox.AddItem("Csv")
            'cboBox.ItemData(cboBox.NewIndex) = 1
            'cboBox.AddItem("Xml")
            'cboBox.ItemData(cboBox.NewIndex) = 2
            'cboBox.FindComboItemData(muSettings.ExportedDataFormat)
            'Me.AddChild(CType(cboBox, UIControl))
            'AddHandler cboBox.ItemSelected, AddressOf ConfigureComboboxItemSelected


            'Screenshot Format
            lblItem = Nothing
            lblItem = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblScreenshotFormat"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 180
                .Width = 157
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Screenshot Format:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblItem, UIControl))

            cboBox = Nothing
            cboBox = New UIComboBox(MyBase.moUILib)
            With cboBox
                .ControlName = "cboScreenshotFormat"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 180
                .Width = 155
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
            cboBox.Clear()
            cboBox.AddItem("Bmp")
            cboBox.ItemData(cboBox.NewIndex) = 0
            cboBox.AddItem("Jpg")
            cboBox.ItemData(cboBox.NewIndex) = 1
            cboBox.AddItem("Tga")
            cboBox.ItemData(cboBox.NewIndex) = 2
            cboBox.AddItem("Png")
            cboBox.ItemData(cboBox.NewIndex) = 3
            cboBox.FindComboItemData(muSettings.ScreenshotFormat)
            Me.AddChild(CType(cboBox, UIControl))
            AddHandler cboBox.ItemSelected, AddressOf ConfigureComboboxItemSelected
        End Sub

        Private Function GetBPP(ByVal uFmt As Format) As Int32
            Select Case uFmt
                Case Format.A1R5G5B5, Format.R5G6B5, Format.X1R5G5B5
                    Return 16
                Case Format.A8B8G8R8, Format.A8R8G8B8, Format.X8R8G8B8
                    Return 32
                Case Else
                    Return -1
            End Select
        End Function
#End Region
#Region "  Textures Config  "
        Private Sub ConfigureTexture()
            Dim lCboLeft As Int32

            Dim lblItem As UILabel = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblRes"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Model Texture Resolution:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblItem, UIControl))

            lCboLeft = lblItem.Left + lblItem.Width + 5

            lblItem = Nothing
            lblItem = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblTerrainRes"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 30
                .Width = 157
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Terrain Resolution:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Texture resolution for the terrain of planets." & vbCrLf & "Setting this value to low can increase performance on fill-rate limited video cards."
            End With
            Me.AddChild(CType(lblItem, UIControl))

            lblItem = Nothing
            lblItem = New UILabel(MyBase.moUILib)
            With lblItem
                .ControlName = "lblFOWRes"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 55
                .Width = 157
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Fog of War Resolution:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblItem, UIControl))

            Dim chkBox As New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkSmoothFOW"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 80
                .Width = 153
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Smooth Fog-of-War"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.SmoothFOW
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Click to toggle smoothing the fog-of-war texture layer."
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkHiResPlanets"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 105
                .Width = 180
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Hi-Res Planet Textures"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
                .Value = muSettings.HiResPlanetTexture
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If On, Planet Textures are rendered. If Off, the textures are calculated." & vbCrLf & "Turning this On will increase load times for NEW solar systems." & vbCrLf & "To see changes to this setting, you will need to reload the solar system."
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkRenderCosmos"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 130
                .Width = 173
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Deep Cosmos"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
                .Value = muSettings.RenderCosmos
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Check to render cosmo effects in space." & vbCrLf & "Turning this off will slightly reduce texture memory" & vbCrLf & "and may improve performance slightly."
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            chkBox = Nothing
            chkBox = New UICheckBox(MyBase.moUILib)
            With chkBox
                .ControlName = "chkRenderPlanetRings"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 155
                .Width = 185
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Planetary Rings"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
                .Value = muSettings.RenderPlanetRings
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Check to render planetary ring effects in space." & vbCrLf & "Turning this off will slightly reduce texture memory" & vbCrLf & "and may improve performance slightly."
            End With
            Me.AddChild(CType(chkBox, UIControl))
            AddHandler chkBox.Click, AddressOf ConfigureCheckboxClick

            Dim lMaxTexRes As Int32 = Math.Min(MyBase.moUILib.oDevice.DeviceCaps.MaxTextureWidth, MyBase.moUILib.oDevice.DeviceCaps.MaxTextureHeight)
            Dim cboBox As New UIComboBox(MyBase.moUILib)
            With cboBox
                .ControlName = "cboFOWRes"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 55
                .Width = 155
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
                .ToolTipText = "Indicates the texture resolution used to define viewable areas of terrain." & vbCrLf & "With Smooth Fog of War Enabled, setting this to low or medium is not as noticeable and will improve performance."
            End With
            cboBox.Clear()
            cboBox.AddItem("Off - No FOW shading")
            cboBox.ItemData(cboBox.NewIndex) = 0
            If lMaxTexRes >= 512 Then
                cboBox.AddItem("Normal Resolution")
                cboBox.ItemData(cboBox.NewIndex) = 512
            End If
            'If lMaxTexRes >= 1024 Then
            '    cboBox.AddItem("Medium Resolution")
            '    cboBox.ItemData(cboBox.NewIndex) = 1024
            'End If
            'If lMaxTexRes >= 2048 Then
            '    cboBox.AddItem("High Resolution")
            '    cboBox.ItemData(cboBox.NewIndex) = 2048
            'End If
            If muSettings.ShowFOWTerrainShading = True Then
                cboBox.FindComboItemData(muSettings.FOWTextureResolution)
            Else
                cboBox.FindComboItemData(0)
            End If

            Me.AddChild(CType(cboBox, UIControl))
            AddHandler cboBox.ItemSelected, AddressOf ConfigureComboboxItemSelected

            cboBox = Nothing
            cboBox = New UIComboBox(MyBase.moUILib)
            With cboBox
                .ControlName = "cboTerrainTexRes"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 30
                .Width = 155
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
                .ToolTipText = "Texture resolution for the terrain of planets." & vbCrLf & "Setting this value to low can increase performance on fill-rate limited video cards." 'lblTerrainTexRes.ToolTipText
            End With
            With cboBox
                .Clear()
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .FindComboItemData(muSettings.TerrainTextureResolution)
            End With
            Me.AddChild(CType(cboBox, UIControl))
            AddHandler cboBox.ItemSelected, AddressOf ConfigureComboboxItemSelected

            cboBox = Nothing
            cboBox = New UIComboBox(MyBase.moUILib)
            With cboBox
                .ControlName = "cboResolution"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 5
                .Width = 155
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
                .ToolTipText = "Sets the resolution of unit and facility models." & vbCrLf & _
                   "Very Low will work on some 64 MB video cards." & vbCrLf & _
                   "Low is recommended to have 128 MB of video memory." & vbCrLf & _
                   "Normal is recommended if you have 256 MB of video memory." & vbCrLf & _
                   "High is recommended for video cards with more than 256 MB of video memory."
            End With
            cboBox.Clear()
            cboBox.AddItem("Very Low Resolution")
            cboBox.ItemData(cboBox.NewIndex) = CInt(EngineSettings.eTextureResOptions.eVeryLowResTextures)
            cboBox.AddItem("Low Resolution")
            cboBox.ItemData(cboBox.NewIndex) = CInt(EngineSettings.eTextureResOptions.eLowResTextures)
            cboBox.AddItem("Normal Resolution")
            cboBox.ItemData(cboBox.NewIndex) = CInt(EngineSettings.eTextureResOptions.eNormResTextures)
            If lMaxTexRes >= 512 Then
                cboBox.AddItem("High Resolution")
                cboBox.ItemData(cboBox.NewIndex) = CInt(EngineSettings.eTextureResOptions.eHiResTextures)
            End If
            cboBox.FindComboItemData(muSettings.ModelTextureResolution)
            Me.AddChild(CType(cboBox, UIControl))
            AddHandler cboBox.ItemSelected, AddressOf ConfigureComboboxItemSelected



        End Sub
#End Region
#Region "  Particles Config  "
        Private Sub ConfigureParticles()
            'lblEngineFX initial props
            Dim lblEngineFX As UILabel = New UILabel(MyBase.moUILib)
            With lblEngineFX
                .ControlName = "lblEngineFX"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 66
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Engine FX:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Manages the Density of Engine FX Particles" & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."
            End With
            Me.AddChild(CType(lblEngineFX, UIControl))

            Dim lCboLeft As Int32 = lblEngineFX.Left + lblEngineFX.Width + 20

            'lblBurnFX initial props
            Dim lblBurnFX As UILabel = New UILabel(MyBase.moUILib)
            With lblBurnFX
                .ControlName = "lblBurnFX"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 30
                .Width = 66
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Burn FX:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Manages the Density of Burn (fire) FX particles." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."
            End With
            Me.AddChild(CType(lblBurnFX, UIControl))

            'lblPlanetFX initial props
            Dim lblPlanetFX As UILabel = New UILabel(MyBase.moUILib)
            With lblPlanetFX
                .ControlName = "lblPlanetFX"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 55
                .Width = 66
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Planet FX:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Manages the Density of Planet particle effects." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this will increase performance on most systems."
            End With
            Me.AddChild(CType(lblPlanetFX, UIControl))

            'chkShieldFX initial props
            Dim chkShieldFX As UICheckBox = New UICheckBox(MyBase.moUILib)
            With chkShieldFX
                .ControlName = "chkShieldFX"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 80
                .Width = 122
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Shield FX"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.RenderShieldFX
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If on, renders shield effects when an entity is hit." & vbCrLf & "Unchecking this value can increase performance in high combat situations."
            End With
            Me.AddChild(CType(chkShieldFX, UIControl))
            AddHandler chkShieldFX.Click, AddressOf ConfigureCheckboxClick

            'chkWormholeAura initial props
            Dim chkWormholeAura As UICheckBox = New UICheckBox(MyBase.moUILib)
            With chkWormholeAura
                .ControlName = "chkWormholeAura"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 155
                .Width = 160
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Wormhole Aura"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.WormholeAura
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If on, renders aura around celestial anomalies. Turning this off may improve performance on some machines."
            End With
            Me.AddChild(CType(chkWormholeAura, UIControl))
            AddHandler chkWormholeAura.Click, AddressOf ConfigureCheckboxClick

            Dim lblStarDensSpace As UILabel = New UILabel(MyBase.moUILib)
            With lblStarDensSpace
                .ControlName = "lblStarDensSpace"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 105
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Starfield Density (Space):"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Indicates how many stars appear in space."
            End With
            Me.AddChild(CType(lblStarDensSpace, UIControl))

            Dim hscrSpaceDens As UIScrollBar = New UIScrollBar(MyBase.moUILib, False)
            With hscrSpaceDens
                .ControlName = "hscrSpaceDens"
                .Left = lblStarDensSpace.Left + lblStarDensSpace.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .MaxValue = 320
                .MinValue = 10
                .Value = muSettings.StarfieldParticlesSpace \ 100
                .SmallChange = 10
                .LargeChange = 10
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
                .ToolTipText = lblStarDensSpace.ToolTipText
            End With
            Me.AddChild(CType(hscrSpaceDens, UIControl))
            AddHandler hscrSpaceDens.ValueChange, AddressOf ConfigureScrollBarScroll

            Dim lblStarDensPlanet As UILabel = New UILabel(MyBase.moUILib)
            With lblStarDensPlanet
                .ControlName = "lblStarDensPlanet"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 130
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Starfield Density (Planet):"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Indicates how many stars appear on planets."
            End With
            Me.AddChild(CType(lblStarDensPlanet, UIControl))

            Dim hscrPlanetDens As UIScrollBar = New UIScrollBar(MyBase.moUILib, False)
            With hscrPlanetDens
                .ControlName = "hscrPlanetDens"
                .Left = lblStarDensSpace.Left + lblStarDensSpace.Width + 5
                .Top = lstOptionSet.Top + 130
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .MaxValue = 400
                .MinValue = 10
                .Value = muSettings.StarfieldParticlesPlanet \ 10
                .SmallChange = 10
                .LargeChange = 10
                .ReverseDirection = False
                .mbAcceptReprocessEvents = True
                .ToolTipText = lblStarDensPlanet.ToolTipText
            End With
            Me.AddChild(CType(hscrPlanetDens, UIControl))
            AddHandler hscrPlanetDens.ValueChange, AddressOf ConfigureScrollBarScroll

            Dim chkPulseBolts As UICheckBox = New UICheckBox(MyBase.moUILib)
            With chkPulseBolts
                .ControlName = "chkPulseBolts"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 180
                .Width = 190
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Pulse Weapon Bolts"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.RenderPulseBolts
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If on, increases the visual appearance of pulse weapons by making them look more like bolts." & vbCrLf & "Turning this off may improve performance on some machines."
            End With
            Me.AddChild(CType(chkPulseBolts, UIControl))
            AddHandler chkPulseBolts.Click, AddressOf ConfigureCheckboxClick

            Dim chkExplosions As UICheckBox = New UICheckBox(MyBase.moUILib)
            With chkExplosions
                .ControlName = "chkExplosions"
                .Left = lblEngineFX.Left
                .Top = lstOptionSet.Top + 205
                .Width = 135
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Explosions"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.RenderExplosionFX
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                '.ToolTipText = "If on, increases the visual appearance of pulse weapons by making them look more like bolts." & vbCrLf & "Turning this off may improve performance on some machines."
            End With
            Me.AddChild(CType(chkExplosions, UIControl))
            AddHandler chkExplosions.Click, AddressOf ConfigureCheckboxClick


            Dim cboPlanetFX As UIComboBox = New UIComboBox(MyBase.moUILib)
            With cboPlanetFX
                .ControlName = "cboPlanetFX"
                .Left = lCboLeft
                .Top = lblPlanetFX.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = "Manages the Density of Planet particle effects." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this will increase performance on most systems."


                .Clear()
                .AddItem("Off")
                .ItemData(.NewIndex) = 0
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .AddItem("Full")
                .ItemData(.NewIndex) = 4
                .FindComboItemData(muSettings.PlanetFXParticles)
            End With
            Me.AddChild(CType(cboPlanetFX, UIControl))
            AddHandler cboPlanetFX.ItemSelected, AddressOf ConfigureComboboxItemSelected

            'cboBurnFX initial props
            Dim cboBurnFX As UIComboBox = New UIComboBox(MyBase.moUILib)
            With cboBurnFX
                .ControlName = "cboBurnFX"
                .Left = lCboLeft
                .Top = lblBurnFX.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = "Manages the Density of Burn (fire) FX particles." & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."

                .Clear()
                .AddItem("Off")
                .ItemData(.NewIndex) = 0
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .AddItem("Full")
                .ItemData(.NewIndex) = 4
                .FindComboItemData(muSettings.BurnFXParticles)
            End With
            Me.AddChild(CType(cboBurnFX, UIControl))
            AddHandler cboBurnFX.ItemSelected, AddressOf ConfigureComboboxItemSelected

            'cboEngineFX initial props
            Dim cboEngineFX As UIComboBox = New UIComboBox(MyBase.moUILib)
            With cboEngineFX
                .ControlName = "cboEngineFX"
                .Left = lCboLeft
                .Top = lblEngineFX.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = "Manages the Density of Engine FX Particles" & vbCrLf & "Full = 100%, High = 75%, Medium = 50%, Low = 25% affecting number of particles." & vbCrLf & "Reducing this may increase performance on some systems."

                .Clear()
                .AddItem("Off")
                .ItemData(.NewIndex) = 0
                .AddItem("Low")
                .ItemData(.NewIndex) = 1
                .AddItem("Medium")
                .ItemData(.NewIndex) = 2
                .AddItem("High")
                .ItemData(.NewIndex) = 3
                .AddItem("Full")
                .ItemData(.NewIndex) = 4
                .FindComboItemData(muSettings.EngineFXParticles)
            End With
            Me.AddChild(CType(cboEngineFX, UIControl))
            AddHandler cboEngineFX.ItemSelected, AddressOf ConfigureComboboxItemSelected

        End Sub
#End Region
#Region "  Lighting Config  "
        Private Sub ConfigureLighting()
            Dim lblQuality As New UILabel(MyBase.moUILib)
            With lblQuality
                .ControlName = "lblQuality"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Quality:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Indicates the quality of lighting and shading." & vbCrLf & "Low uses DirectX 9.0c standard lighting and will provide the highest performance." & vbCrLf & "Medium uses Per Pixel lighting and shading which requires 2.0 compatible shaders." & vbCrLf & "High uses Bump mapping in combination with per pixel shading."
            End With
            Me.AddChild(CType(lblQuality, UIControl))

            Dim lblIllumMaps As New UILabel(MyBase.moUILib)
            With lblIllumMaps
                .ControlName = "lblIllumMaps"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 30
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Illumination Maps:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Indicates whether to use illumination maps and at what texture resolution."
            End With
            Me.AddChild(CType(lblIllumMaps, UIControl))

            Dim lblGlowFX As New UILabel(MyBase.moUILib)
            With lblGlowFX
                .ControlName = "lblGlowFX"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 55
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Glow FX Intensity:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Adjusts the use of glow post-process effect." & vbCrLf & "Adjusting this all the way down (to the left) may improve performance."
            End With
            Me.AddChild(CType(lblGlowFX, UIControl))

            Dim hscrGlowFX As UIScrollBar = New UIScrollBar(MyBase.moUILib, False)
            With hscrGlowFX
                .ControlName = "hscrGlowFX"
                .Left = lblGlowFX.Left + lblGlowFX.Width + 5
                .Top = lblGlowFX.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .MaxValue = 10
                .MinValue = 0
                .Value = CInt(muSettings.PostGlowAmt * 10)
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = False
                .ToolTipText = lblGlowFX.ToolTipText
            End With
            Me.AddChild(CType(hscrGlowFX, UIControl))
            AddHandler hscrGlowFX.ValueChange, AddressOf ConfigureScrollBarScroll

            Dim lblAmbient As New UILabel(MyBase.moUILib)
            With lblAmbient
                .ControlName = "lblAmbient"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Gamma Adjust:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Adjusts the overall light level of the game." & vbCrLf & "Useful for when the game is too dark." & vbCrLf & "Only works if the Light Quality is not Low."
            End With
            Me.AddChild(CType(lblAmbient, UIControl))

            Dim hscrAmbient As UIScrollBar = New UIScrollBar(MyBase.moUILib, False)
            With hscrAmbient
                .ControlName = "hscrAmbient"
                .Left = lblAmbient.Left + lblAmbient.Width + 5
                .Top = lblAmbient.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .MaxValue = 50
                .MinValue = 0
                .Value = muSettings.AmbientLevel
                .SmallChange = 1
                .LargeChange = 4
                .ReverseDirection = False
                .mbAcceptReprocessEvents = False
                .ToolTipText = lblAmbient.ToolTipText
            End With
            Me.AddChild(CType(hscrAmbient, UIControl))
            AddHandler hscrAmbient.ValueChange, AddressOf ConfigureScrollBarScroll

            Dim chkSpecularMaps As New UICheckBox(MyBase.moUILib)
            With chkSpecularMaps
                .ControlName = "chkSpecularMaps"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 157
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Render Specular Maps"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.RenderSpecularMap
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "When checked, units and facilities are rendered with specular maps."
            End With
            Me.AddChild(CType(chkSpecularMaps, UIControl))
            AddHandler chkSpecularMaps.Click, AddressOf ConfigureCheckboxClick

            Dim chkBumpTerrain As New UICheckBox(MyBase.moUILib)
            With chkBumpTerrain
                .ControlName = "chkBumpTerrain"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 130
                .Width = 130
                .Height = 18
                .Enabled = MyBase.moUILib.oDevice.DeviceCaps.PixelShaderVersion.Major >= 2
                'If GFXEngine.bNVidiaCard = False Then .Enabled = False
                .Visible = True
                .Caption = "Bump Map Terrain"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.BumpMapTerrain
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Requires shader 2.0 and incurs a hefty performance penalty on some machines."
            End With
            Me.AddChild(CType(chkBumpTerrain, UIControl))
            AddHandler chkBumpTerrain.Click, AddressOf ConfigureCheckboxClick

            Dim chkBumpPlanetModels As New UICheckBox(MyBase.moUILib)
            With chkBumpPlanetModels
                .ControlName = "chkBumpPlanetModels"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 155
                .Width = 173
                .Height = 18
                .Enabled = MyBase.moUILib.oDevice.DeviceCaps.PixelShaderVersion.Major >= 2
                .Visible = True
                .Caption = "Bump Map Planet Models"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.BumpMapPlanetModel
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Requires shader 2.0 and incurs a hefty performance penalty on some machines."
            End With
            Me.AddChild(CType(chkBumpPlanetModels, UIControl))
            AddHandler chkBumpPlanetModels.Click, AddressOf ConfigureCheckboxClick

            Dim chkIllumTerrain As New UICheckBox(MyBase.moUILib)
            With chkIllumTerrain
                .ControlName = "chkIllumTerrain"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 180
                .Width = 162
                .Height = 18
                .Enabled = MyBase.moUILib.oDevice.DeviceCaps.PixelShaderVersion.Major >= 2
                'If GFXEngine.bNVidiaCard = False Then .Enabled = False
                .Visible = True
                .Caption = "Illuminate Planet Terrain"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = muSettings.IlluminationMapTerrain
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "Requires shader 2.0 and the Bump Terrain option on. When on, city creep is illuminated."
            End With
            Me.AddChild(CType(chkIllumTerrain, UIControl))
            AddHandler chkIllumTerrain.Click, AddressOf ConfigureCheckboxClick

            Dim cboIllumMap As New UIComboBox(MyBase.moUILib)
            With cboIllumMap
                .ControlName = "cboIllumMap"
                .Left = lblIllumMaps.Left + lblIllumMaps.Width + 5
                .Top = lblIllumMaps.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = lblIllumMaps.ToolTipText

                .Clear()
                .AddItem("Illumination Maps Off")
                .ItemData(.NewIndex) = -1
                .AddItem("Very Low Resolution")
                .ItemData(.NewIndex) = CInt(EngineSettings.eTextureResOptions.eVeryLowResTextures)
                .AddItem("Low Resolution")
                .ItemData(.NewIndex) = CInt(EngineSettings.eTextureResOptions.eLowResTextures)
                .AddItem("Normal Resolution")
                .ItemData(.NewIndex) = CInt(EngineSettings.eTextureResOptions.eNormResTextures)
                If Math.Min(MyBase.moUILib.oDevice.DeviceCaps.MaxTextureWidth, MyBase.moUILib.oDevice.DeviceCaps.MaxTextureHeight) >= 512 Then
                    .AddItem("High Resolution")
                    .ItemData(.NewIndex) = CInt(EngineSettings.eTextureResOptions.eHiResTextures)
                End If
                .FindComboItemData(muSettings.IlluminationMap)
            End With
            AddHandler cboIllumMap.ItemSelected, AddressOf ConfigureComboboxItemSelected
            Me.AddChild(CType(cboIllumMap, UIControl))

            Dim cboQuality As New UIComboBox(MyBase.moUILib)
            With cboQuality
                .ControlName = "cboQuality"
                .Left = lblQuality.Left + lblQuality.Width + 5
                .Top = lblQuality.Top
                .Width = 155
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = lblQuality.ToolTipText

                .Clear()
                .AddItem("Low Quality")
                .ItemData(.NewIndex) = EngineSettings.LightQualitySetting.VSPS1

                If MyBase.moUILib.oDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
                    .AddItem("Medium Quality")
                    .ItemData(.NewIndex) = EngineSettings.LightQualitySetting.PerPixel
                    .AddItem("High Quality")
                    .ItemData(.NewIndex) = EngineSettings.LightQualitySetting.BumpMap
                End If
                .FindComboItemData(muSettings.LightQuality)
            End With
            AddHandler cboQuality.ItemSelected, AddressOf ConfigureComboboxItemSelected
            Me.AddChild(CType(cboQuality, UIControl))

        End Sub
#End Region
#Region "  Interface Colors Config  "
        Private Sub ConfigureInterfaceColors()
            Dim lblBorder As New UILabel(MyBase.moUILib)
            With lblBorder
                .ControlName = "lblBorder"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Border Color"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBorder, UIControl))

            Dim btnSetBorder As New UIButton(MyBase.moUILib)
            With btnSetBorder
                .ControlName = "btnSetBorder"
                .Left = lblBorder.Left + lblBorder.Width + 5
                .Top = lblBorder.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBorder, UIControl))
            AddHandler btnSetBorder.Click, AddressOf ConfigureColorButtonClick

            Dim lblFill As New UILabel(MyBase.moUILib)
            With lblFill
                .ControlName = "lblFill"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 30
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Fill Color"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblFill, UIControl))

            Dim btnSetFill As New UIButton(MyBase.moUILib)
            With btnSetFill
                .ControlName = "btnSetFill"
                .Left = lblFill.Left + lblFill.Width + 5
                .Top = lblFill.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetFill, UIControl))
            AddHandler btnSetFill.Click, AddressOf ConfigureColorButtonClick

            Dim lblTextboxFore As New UILabel(MyBase.moUILib)
            With lblTextboxFore
                .ControlName = "lblTextboxFore"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 55
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Textbox Text Color"
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTextboxFore, UIControl))

            Dim btnSetTextboxFore As New UIButton(MyBase.moUILib)
            With btnSetTextboxFore
                .ControlName = "btnSetTextboxFore"
                .Left = lblTextboxFore.Left + lblTextboxFore.Width + 5
                .Top = lblTextboxFore.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetTextboxFore, UIControl))
            AddHandler btnSetTextboxFore.Click, AddressOf ConfigureColorButtonClick

            Dim lblTextboxBack As New UILabel(MyBase.moUILib)
            With lblTextboxBack
                .ControlName = "lblTextboxBack"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Textbox Back Color"
                .ForeColor = muSettings.InterfaceTextBoxFillColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTextboxBack, UIControl))

            Dim btnSetTextboxBack As New UIButton(MyBase.moUILib)
            With btnSetTextboxBack
                .ControlName = "btnSetTextboxBack"
                .Left = lblTextboxBack.Left + lblTextboxBack.Width + 5
                .Top = lblTextboxBack.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetTextboxBack, UIControl))
            AddHandler btnSetTextboxBack.Click, AddressOf ConfigureColorButtonClick

            Dim lblButtonColor As New UILabel(MyBase.moUILib)
            With lblButtonColor
                .ControlName = "lblButtonColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Button Color"
                .ForeColor = muSettings.InterfaceButtonColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblButtonColor, UIControl))

            Dim btnSetButtonColor As New UIButton(MyBase.moUILib)
            With btnSetButtonColor
                .ControlName = "btnSetButtonColor"
                .Left = lblButtonColor.Left + lblButtonColor.Width + 5
                .Top = lblButtonColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetButtonColor, UIControl))
            AddHandler btnSetButtonColor.Click, AddressOf ConfigureColorButtonClick

            Dim btnResetToDefault As New UIButton(MyBase.moUILib)
            With btnResetToDefault
                .ControlName = "btnResetToDefault"
                .Left = lblButtonColor.Left + lblButtonColor.Width + 5
                .Top = lblButtonColor.Top + lblButtonColor.Height + 50
                .Width = 140
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Reset to Default"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnResetToDefault, UIControl))
            AddHandler btnResetToDefault.Click, AddressOf ResetInterfaceDefaults
        End Sub
#End Region
#Region "  Chat Window Colors  "
        Private Sub ConfigureChatWindowColors()
            Dim lblDefaultColor As New UILabel(MyBase.moUILib)
            With lblDefaultColor
                .ControlName = "lblDefaultColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Default Color"
                .ForeColor = muSettings.DefaultChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblDefaultColor, UIControl))

            Dim btnSetDefault As New UIButton(MyBase.moUILib)
            With btnSetDefault
                .ControlName = "btnSetDefault"
                .Left = lblDefaultColor.Left + lblDefaultColor.Width + 5
                .Top = lblDefaultColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetDefault, UIControl))
            AddHandler btnSetDefault.Click, AddressOf ConfigureColorButtonClick

            Dim lblAlertColor As New UILabel(MyBase.moUILib)
            With lblAlertColor
                .ControlName = "lblAlertColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 30
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Alert Color"
                .ForeColor = muSettings.AlertChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblAlertColor, UIControl))

            Dim btnSetError As New UIButton(MyBase.moUILib)
            With btnSetError
                .ControlName = "btnSetError"
                .Left = lblAlertColor.Left + lblAlertColor.Width + 5
                .Top = lblAlertColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetError, UIControl))
            AddHandler btnSetError.Click, AddressOf ConfigureColorButtonClick

            Dim lblStatusColor As New UILabel(MyBase.moUILib)
            With lblStatusColor
                .ControlName = "lblStatusColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 55
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Status/MOTD Color"
                .ForeColor = muSettings.StatusChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblStatusColor, UIControl))

            Dim btnSetStatus As New UIButton(MyBase.moUILib)
            With btnSetStatus
                .ControlName = "btnSetStatus"
                .Left = lblStatusColor.Left + lblStatusColor.Width + 5
                .Top = lblStatusColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetStatus, UIControl))
            AddHandler btnSetStatus.Click, AddressOf ConfigureColorButtonClick

            Dim lblLocalColor As New UILabel(MyBase.moUILib)
            With lblLocalColor
                .ControlName = "lblLocalColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Local Color"
                .ForeColor = muSettings.LocalChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblLocalColor, UIControl))

            Dim btnSetLocal As New UIButton(MyBase.moUILib)
            With btnSetLocal
                .ControlName = "btnSetLocal"
                .Left = lblLocalColor.Left + lblLocalColor.Width + 5
                .Top = lblLocalColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetLocal, UIControl))
            AddHandler btnSetLocal.Click, AddressOf ConfigureColorButtonClick

            Dim lblGuildColor As New UILabel(MyBase.moUILib)
            With lblGuildColor
                .ControlName = "lblGuildColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Guild Color"
                .ForeColor = muSettings.GuildChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblGuildColor, UIControl))

            Dim btnSetGuild As New UIButton(MyBase.moUILib)
            With btnSetGuild
                .ControlName = "btnSetGuild"
                .Left = lblGuildColor.Left + lblGuildColor.Width + 5
                .Top = lblGuildColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetGuild, UIControl))
            AddHandler btnSetGuild.Click, AddressOf ConfigureColorButtonClick

            Dim lblSenateColor As New UILabel(MyBase.moUILib)
            With lblSenateColor
                .ControlName = "lblSenateColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 130
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Senate Color"
                .ForeColor = muSettings.SenateChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblSenateColor, UIControl))

            Dim btnSetSenate As New UIButton(MyBase.moUILib)
            With btnSetSenate
                .ControlName = "btnSetSenate"
                .Left = lblSenateColor.Left + lblSenateColor.Width + 5
                .Top = lblSenateColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetSenate, UIControl))
            AddHandler btnSetSenate.Click, AddressOf ConfigureColorButtonClick

            Dim lblPMColor As New UILabel(MyBase.moUILib)
            With lblPMColor
                .ControlName = "lblPMColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 155
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Private Message Color"
                .ForeColor = muSettings.PMChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPMColor, UIControl))

            Dim btnSetPM As New UIButton(MyBase.moUILib)
            With btnSetPM
                .ControlName = "btnSetPM"
                .Left = lblPMColor.Left + lblPMColor.Width + 5
                .Top = lblPMColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetPM, UIControl))
            AddHandler btnSetPM.Click, AddressOf ConfigureColorButtonClick

            Dim lblChannelColor As New UILabel(MyBase.moUILib)
            With lblChannelColor
                .ControlName = "lblChannelColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 180
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Channel Color"
                .ForeColor = muSettings.ChannelChatColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblChannelColor, UIControl))

            Dim btnSetChannel As New UIButton(MyBase.moUILib)
            With btnSetChannel
                .ControlName = "btnSetChannel"
                .Left = lblChannelColor.Left + lblChannelColor.Width + 5
                .Top = lblChannelColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetChannel, UIControl))
            AddHandler btnSetChannel.Click, AddressOf ConfigureColorButtonClick
        End Sub
#End Region
#Region "  Identification Colors  "
        Private Sub ConfigureIdentificationColors()
            Dim lblNeutralColor As New UILabel(MyBase.moUILib)
            With lblNeutralColor
                .ControlName = "lblNeutralColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Neutral Color"
                .ForeColor = ConvertVector4ToColor(muSettings.NeutralAssetColor)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNeutralColor, UIControl))

            Dim btnSetNeutral As New UIButton(MyBase.moUILib)
            With btnSetNeutral
                .ControlName = "btnSetNeutral"
                .Left = lblNeutralColor.Left + lblNeutralColor.Width + 5
                .Top = lblNeutralColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetNeutral, UIControl))
            AddHandler btnSetNeutral.Click, AddressOf ConfigureColorButtonClick

            Dim lblMyColor As New UILabel(MyBase.moUILib)
            With lblMyColor
                .ControlName = "lblMyColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 30
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Player Color"
                .ForeColor = ConvertVector4ToColor(muSettings.MyAssetColor)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMyColor, UIControl))

            Dim btnSetMyColor As New UIButton(MyBase.moUILib)
            With btnSetMyColor
                .ControlName = "btnSetMyColor"
                .Left = lblMyColor.Left + lblMyColor.Width + 5
                .Top = lblMyColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMyColor, UIControl))
            AddHandler btnSetMyColor.Click, AddressOf ConfigureColorButtonClick

            Dim lblEnemyColor As New UILabel(MyBase.moUILib)
            With lblEnemyColor
                .ControlName = "lblEnemyColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 55
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Enemy Color"
                .ForeColor = ConvertVector4ToColor(muSettings.EnemyAssetColor)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblEnemyColor, UIControl))

            Dim btnSetEnemy As New UIButton(MyBase.moUILib)
            With btnSetEnemy
                .ControlName = "btnSetEnemy"
                .Left = lblEnemyColor.Left + lblEnemyColor.Width + 5
                .Top = lblEnemyColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetEnemy, UIControl))
            AddHandler btnSetEnemy.Click, AddressOf ConfigureColorButtonClick

            Dim lblAllyColor As New UILabel(MyBase.moUILib)
            With lblAllyColor
                .ControlName = "lblAllyColor"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Ally Color"
                .ForeColor = ConvertVector4ToColor(muSettings.AllyAssetColor)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblAllyColor, UIControl))

            Dim btnSetAlly As New UIButton(MyBase.moUILib)
            With btnSetAlly
                .ControlName = "btnSetAlly"
                .Left = lblAllyColor.Left + lblAllyColor.Width + 5
                .Top = lblAllyColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetAlly, UIControl))
            AddHandler btnSetAlly.Click, AddressOf ConfigureColorButtonClick

            Dim lblGuildColor As New UILabel(MyBase.moUILib)
            With lblGuildColor
                .ControlName = "lblGuildColorID"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Guild Color"
                .ForeColor = ConvertVector4ToColor(muSettings.GuildAssetColor)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblGuildColor, UIControl))

            Dim btnSetGuild As New UIButton(MyBase.moUILib)
            With btnSetGuild
                .ControlName = "btnSetGuildColor"
                .Left = lblGuildColor.Left + lblGuildColor.Width + 5
                .Top = lblGuildColor.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetGuild, UIControl))
            AddHandler btnSetGuild.Click, AddressOf ConfigureColorButtonClick

            Dim lblTactical As New UILabel(MyBase.moUILib)
            With lblTactical
                .ControlName = "lblTactical"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 130
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Tactical Data"
                .ForeColor = ConvertVector4ToColor(muSettings.TacticalAssetColor)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTactical, UIControl))

            Dim btnSetTactical As New UIButton(MyBase.moUILib)
            With btnSetTactical
                .ControlName = "btnSetTactical"
                .Left = lblTactical.Left + lblTactical.Width + 5
                .Top = lblTactical.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetTactical, UIControl))
            AddHandler btnSetTactical.Click, AddressOf ConfigureColorButtonClick

            Dim lblNotice As New UILabel(MyBase.moUILib)
            With lblNotice
                .ControlName = "lblNotice"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 155
                .Width = Me.Width - .Left
                .Height = 32
                .Enabled = True
                .Visible = True
                .Caption = "NOTE: These settings are only applied if the Light Quality" & vbCrLf & "under the Lighting section is higher than Low Quality."
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
            End With
            Me.AddChild(CType(lblNotice, UIControl))
        End Sub
#End Region
#Region "  Build Ghost Colors "
        Private Function ConvertFromBuildGhostColor(ByVal clrVal As System.Drawing.Color) As System.Drawing.Color
            Return System.Drawing.Color.FromArgb(255, clrVal.R, clrVal.G, clrVal.B)
        End Function
        Private Function ConvertToBuildGhostColor(ByVal clrVal As System.Drawing.Color) As System.Drawing.Color
            Return System.Drawing.Color.FromArgb(64, clrVal.R, clrVal.G, clrVal.B)
        End Function
        Private Sub ConfigureBuildGhostColors()
            Dim lblBuildGhostAcid As New UILabel(MyBase.moUILib)
            With lblBuildGhostAcid
                .ControlName = "lblBuildGhostAcid"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Acid Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.AcidBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostAcid, UIControl))

            Dim btnSetBuildGhostAcid As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostAcid
                .ControlName = "btnSetBuildGhostAcid"
                .Left = lblBuildGhostAcid.Left + lblBuildGhostAcid.Width + 5
                .Top = lblBuildGhostAcid.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostAcid, UIControl))
            AddHandler btnSetBuildGhostAcid.Click, AddressOf ConfigureColorButtonClick

            Dim lblBuildGhostAdaptable As New UILabel(MyBase.moUILib)
            With lblBuildGhostAdaptable
                .ControlName = "lblBuildGhostAdaptable"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 30
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Adaptable Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.AdaptableBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostAdaptable, UIControl))

            Dim btnSetBuildGhostAdaptable As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostAdaptable
                .ControlName = "btnSetBuildGhostAdaptable"
                .Left = lblBuildGhostAdaptable.Left + lblBuildGhostAdaptable.Width + 5
                .Top = lblBuildGhostAdaptable.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostAdaptable, UIControl))
            AddHandler btnSetBuildGhostAdaptable.Click, AddressOf ConfigureColorButtonClick

            Dim lblBuildGhostBarren As New UILabel(MyBase.moUILib)
            With lblBuildGhostBarren
                .ControlName = "lblBuildGhostBarren"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 55
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Barren Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.BarrenBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostBarren, UIControl))

            Dim btnSetBuildGhostBarren As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostBarren
                .ControlName = "btnSetBuildGhostBarren"
                .Left = lblBuildGhostBarren.Left + lblBuildGhostBarren.Width + 5
                .Top = lblBuildGhostBarren.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostBarren, UIControl))
            AddHandler btnSetBuildGhostBarren.Click, AddressOf ConfigureColorButtonClick

            Dim lblBuildGhostDesert As New UILabel(MyBase.moUILib)
            With lblBuildGhostDesert
                .ControlName = "lblBuildGhostDesert"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Desert Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.DesertBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostDesert, UIControl))

            Dim btnSetBuildGhostDesert As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostDesert
                .ControlName = "btnSetBuildGhostDesert"
                .Left = lblBuildGhostDesert.Left + lblBuildGhostDesert.Width + 5
                .Top = lblBuildGhostDesert.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostDesert, UIControl))
            AddHandler btnSetBuildGhostDesert.Click, AddressOf ConfigureColorButtonClick

            Dim lblBuildGhostIce As New UILabel(MyBase.moUILib)
            With lblBuildGhostIce
                .ControlName = "lblBuildGhostIce"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Ice Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.IceBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostIce, UIControl))

            Dim btnSetBuildGhostIce As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostIce
                .ControlName = "btnSetBuildGhostIce"
                .Left = lblBuildGhostIce.Left + lblBuildGhostIce.Width + 5
                .Top = lblBuildGhostIce.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostIce, UIControl))
            AddHandler btnSetBuildGhostIce.Click, AddressOf ConfigureColorButtonClick

            Dim lblBuildGhostLava As New UILabel(MyBase.moUILib)
            With lblBuildGhostLava
                .ControlName = "lblBuildGhostLava"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 130
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Lava Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.LavaBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostLava, UIControl))

            Dim btnSetBuildGhostLava As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostLava
                .ControlName = "btnSetBuildGhostLava"
                .Left = lblBuildGhostLava.Left + lblBuildGhostLava.Width + 5
                .Top = lblBuildGhostLava.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostLava, UIControl))
            AddHandler btnSetBuildGhostLava.Click, AddressOf ConfigureColorButtonClick

            Dim lblBuildGhostTerran As New UILabel(MyBase.moUILib)
            With lblBuildGhostTerran
                .ControlName = "lblBuildGhostTerran"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 155
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Terran Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.TerranBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostTerran, UIControl))

            Dim btnSetBuildGhostTerran As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostTerran
                .ControlName = "btnSetBuildGhostTerran"
                .Left = lblBuildGhostTerran.Left + lblBuildGhostTerran.Width + 5
                .Top = lblBuildGhostTerran.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostTerran, UIControl))
            AddHandler btnSetBuildGhostTerran.Click, AddressOf ConfigureColorButtonClick

            Dim lblBuildGhostWaterworld As New UILabel(MyBase.moUILib)
            With lblBuildGhostWaterworld
                .ControlName = "lblBuildGhostWaterworld"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 180
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Waterworld Planets"
                .ForeColor = ConvertFromBuildGhostColor(muSettings.WaterworldBuildGhost)
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBuildGhostWaterworld, UIControl))

            Dim btnSetBuildGhostWaterworld As New UIButton(MyBase.moUILib)
            With btnSetBuildGhostWaterworld
                .ControlName = "btnSetBuildGhostWaterworld"
                .Left = lblBuildGhostWaterworld.Left + lblBuildGhostWaterworld.Width + 5
                .Top = lblBuildGhostWaterworld.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetBuildGhostWaterworld, UIControl))
            AddHandler btnSetBuildGhostWaterworld.Click, AddressOf ConfigureColorButtonClick
        End Sub
#End Region
#Region "  Mineral Cache Colors  "
        Private Sub ConfigureMineralCacheColors()
            Dim lblMineralCacheAcid As New UILabel(MyBase.moUILib)
            With lblMineralCacheAcid
                .ControlName = "lblMineralCacheAcid"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Acid Planets"
                .ForeColor = muSettings.AcidMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheAcid, UIControl))

            Dim btnSetMineralCacheAcid As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheAcid
                .ControlName = "btnSetMineralCacheAcid"
                .Left = lblMineralCacheAcid.Left + lblMineralCacheAcid.Width + 5
                .Top = lblMineralCacheAcid.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheAcid, UIControl))
            AddHandler btnSetMineralCacheAcid.Click, AddressOf ConfigureColorButtonClick

            Dim lblMineralCacheAdaptable As New UILabel(MyBase.moUILib)
            With lblMineralCacheAdaptable
                .ControlName = "lblMineralCacheAdaptable"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 30
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Adaptable Planets"
                .ForeColor = muSettings.AcidMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheAdaptable, UIControl))

            Dim btnSetMineralCacheAdaptable As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheAdaptable
                .ControlName = "btnSetMineralCacheAdaptable"
                .Left = lblMineralCacheAdaptable.Left + lblMineralCacheAdaptable.Width + 5
                .Top = lblMineralCacheAdaptable.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheAdaptable, UIControl))
            AddHandler btnSetMineralCacheAdaptable.Click, AddressOf ConfigureColorButtonClick

            Dim lblMineralCacheBarren As New UILabel(MyBase.moUILib)
            With lblMineralCacheBarren
                .ControlName = "lblMineralCacheBarren"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 55
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Barren Planets"
                .ForeColor = muSettings.BarrenMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheBarren, UIControl))

            Dim btnSetMineralCacheBarren As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheBarren
                .ControlName = "btnSetMineralCacheBarren"
                .Left = lblMineralCacheBarren.Left + lblMineralCacheBarren.Width + 5
                .Top = lblMineralCacheBarren.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheBarren, UIControl))
            AddHandler btnSetMineralCacheBarren.Click, AddressOf ConfigureColorButtonClick

            Dim lblMineralCacheDesert As New UILabel(MyBase.moUILib)
            With lblMineralCacheDesert
                .ControlName = "lblMineralCacheDesert"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Desert Planets"
                .ForeColor = muSettings.DesertMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheDesert, UIControl))

            Dim btnSetMineralCacheDesert As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheDesert
                .ControlName = "btnSetMineralCacheDesert"
                .Left = lblMineralCacheDesert.Left + lblMineralCacheDesert.Width + 5
                .Top = lblMineralCacheDesert.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheDesert, UIControl))
            AddHandler btnSetMineralCacheDesert.Click, AddressOf ConfigureColorButtonClick

            Dim lblMineralCacheIce As New UILabel(MyBase.moUILib)
            With lblMineralCacheIce
                .ControlName = "lblMineralCacheIce"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Ice Planets"
                .ForeColor = muSettings.IceMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheIce, UIControl))

            Dim btnSetMineralCacheIce As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheIce
                .ControlName = "btnSetMineralCacheIce"
                .Left = lblMineralCacheIce.Left + lblMineralCacheIce.Width + 5
                .Top = lblMineralCacheIce.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheIce, UIControl))
            AddHandler btnSetMineralCacheIce.Click, AddressOf ConfigureColorButtonClick

            Dim lblMineralCacheLava As New UILabel(MyBase.moUILib)
            With lblMineralCacheLava
                .ControlName = "lblMineralCacheLava"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 130
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Lava Planets"
                .ForeColor = muSettings.LavaMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheLava, UIControl))

            Dim btnSetMineralCacheLava As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheLava
                .ControlName = "btnSetMineralCacheLava"
                .Left = lblMineralCacheLava.Left + lblMineralCacheLava.Width + 5
                .Top = lblMineralCacheLava.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheLava, UIControl))
            AddHandler btnSetMineralCacheLava.Click, AddressOf ConfigureColorButtonClick

            Dim lblMineralCacheTerran As New UILabel(MyBase.moUILib)
            With lblMineralCacheTerran
                .ControlName = "lblMineralCacheTerran"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 155
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Terran Planets"
                .ForeColor = muSettings.TerranMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheTerran, UIControl))

            Dim btnSetMineralCacheTerran As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheTerran
                .ControlName = "btnSetMineralCacheTerran"
                .Left = lblMineralCacheTerran.Left + lblMineralCacheTerran.Width + 5
                .Top = lblMineralCacheTerran.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheTerran, UIControl))
            AddHandler btnSetMineralCacheTerran.Click, AddressOf ConfigureColorButtonClick

            Dim lblMineralCacheWaterworld As New UILabel(MyBase.moUILib)
            With lblMineralCacheWaterworld
                .ControlName = "lblMineralCacheWaterworld"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 180
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Waterworld Planets"
                .ForeColor = muSettings.WaterworldMineralCache
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineralCacheWaterworld, UIControl))

            Dim btnSetMineralCacheWaterworld As New UIButton(MyBase.moUILib)
            With btnSetMineralCacheWaterworld
                .ControlName = "btnSetMineralCacheWaterworld"
                .Left = lblMineralCacheWaterworld.Left + lblMineralCacheWaterworld.Width + 5
                .Top = lblMineralCacheWaterworld.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMineralCacheWaterworld, UIControl))
            AddHandler btnSetMineralCacheWaterworld.Click, AddressOf ConfigureColorButtonClick
        End Sub
#End Region
#Region "  Minimap Angle Colors  "
        Private Sub ConfigureMinimapAngleColors()
            Dim lblMinimapAngleAcid As New UILabel(MyBase.moUILib)
            With lblMinimapAngleAcid
                .ControlName = "lblMinimapAngleAcid"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Acid Planets"
                .ForeColor = muSettings.AcidMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleAcid, UIControl))

            Dim btnSetMinimapAngleAcid As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleAcid
                .ControlName = "btnSetMinimapAngleAcid"
                .Left = lblMinimapAngleAcid.Left + lblMinimapAngleAcid.Width + 5
                .Top = lblMinimapAngleAcid.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleAcid, UIControl))
            AddHandler btnSetMinimapAngleAcid.Click, AddressOf ConfigureColorButtonClick

            Dim lblMinimapAngleAdaptable As New UILabel(MyBase.moUILib)
            With lblMinimapAngleAdaptable
                .ControlName = "lblMinimapAngleAdaptable"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 30
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Adaptable Planets"
                .ForeColor = muSettings.AdaptableMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleAdaptable, UIControl))

            Dim btnSetMinimapAngleAdaptable As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleAdaptable
                .ControlName = "btnSetMinimapAngleAdaptable"
                .Left = lblMinimapAngleAdaptable.Left + lblMinimapAngleAdaptable.Width + 5
                .Top = lblMinimapAngleAdaptable.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleAdaptable, UIControl))
            AddHandler btnSetMinimapAngleAdaptable.Click, AddressOf ConfigureColorButtonClick

            Dim lblMinimapAngleBarren As New UILabel(MyBase.moUILib)
            With lblMinimapAngleBarren
                .ControlName = "lblMinimapAngleBarren"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 55
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Barren Planets"
                .ForeColor = muSettings.BarrenMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleBarren, UIControl))

            Dim btnSetMinimapAngleBarren As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleBarren
                .ControlName = "btnSetMinimapAngleBarren"
                .Left = lblMinimapAngleBarren.Left + lblMinimapAngleBarren.Width + 5
                .Top = lblMinimapAngleBarren.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleBarren, UIControl))
            AddHandler btnSetMinimapAngleBarren.Click, AddressOf ConfigureColorButtonClick

            Dim lblMinimapAngleDesert As New UILabel(MyBase.moUILib)
            With lblMinimapAngleDesert
                .ControlName = "lblMinimapAngleDesert"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 80
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Desert Planets"
                .ForeColor = muSettings.DesertMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleDesert, UIControl))

            Dim btnSetMinimapAngleDesert As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleDesert
                .ControlName = "btnSetMinimapAngleDesert"
                .Left = lblMinimapAngleDesert.Left + lblMinimapAngleDesert.Width + 5
                .Top = lblMinimapAngleDesert.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleDesert, UIControl))
            AddHandler btnSetMinimapAngleDesert.Click, AddressOf ConfigureColorButtonClick

            Dim lblMinimapAngleIce As New UILabel(MyBase.moUILib)
            With lblMinimapAngleIce
                .ControlName = "lblMinimapAngleIce"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 105
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Ice Planets"
                .ForeColor = muSettings.IceMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleIce, UIControl))

            Dim btnSetMinimapAngleIce As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleIce
                .ControlName = "btnSetMinimapAngleIce"
                .Left = lblMinimapAngleIce.Left + lblMinimapAngleIce.Width + 5
                .Top = lblMinimapAngleIce.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleIce, UIControl))
            AddHandler btnSetMinimapAngleIce.Click, AddressOf ConfigureColorButtonClick

            Dim lblMinimapAngleLava As New UILabel(MyBase.moUILib)
            With lblMinimapAngleLava
                .ControlName = "lblMinimapAngleLava"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 130
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Lava Planets"
                .ForeColor = muSettings.LavaMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleLava, UIControl))

            Dim btnSetMinimapAngleLava As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleLava
                .ControlName = "btnSetMinimapAngleLava"
                .Left = lblMinimapAngleLava.Left + lblMinimapAngleLava.Width + 5
                .Top = lblMinimapAngleLava.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleLava, UIControl))
            AddHandler btnSetMinimapAngleLava.Click, AddressOf ConfigureColorButtonClick

            Dim lblMinimapAngleTerran As New UILabel(MyBase.moUILib)
            With lblMinimapAngleTerran
                .ControlName = "lblMinimapAngleTerran"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 155
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Terran Planets"
                .ForeColor = muSettings.TerranMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleTerran, UIControl))

            Dim btnSetMinimapAngleTerran As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleTerran
                .ControlName = "btnSetMinimapAngleTerran"
                .Left = lblMinimapAngleTerran.Left + lblMinimapAngleTerran.Width + 5
                .Top = lblMinimapAngleTerran.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleTerran, UIControl))
            AddHandler btnSetMinimapAngleTerran.Click, AddressOf ConfigureColorButtonClick

            Dim lblMinimapAngleWaterworld As New UILabel(MyBase.moUILib)
            With lblMinimapAngleWaterworld
                .ControlName = "lblMinimapAngleWaterworld"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 5
                .Top = lstOptionSet.Top + 180
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Waterworld Planets"
                .ForeColor = muSettings.WaterworldMinimapAngle
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMinimapAngleWaterworld, UIControl))

            Dim btnSetMinimapAngleWaterworld As New UIButton(MyBase.moUILib)
            With btnSetMinimapAngleWaterworld
                .ControlName = "btnSetMinimapAngleWaterworld"
                .Left = lblMinimapAngleWaterworld.Left + lblMinimapAngleWaterworld.Width + 5
                .Top = lblMinimapAngleWaterworld.Top
                .Width = 50
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Set"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSetMinimapAngleWaterworld, UIControl))
            AddHandler btnSetMinimapAngleWaterworld.Click, AddressOf ConfigureColorButtonClick
        End Sub
#End Region
#Region "  Water Options  "
        Private Sub ConfigureWater()
            Dim lCboLeft As Int32

            Dim lblTechnique As UILabel = New UILabel(MyBase.moUILib)
            With lblTechnique
                .ControlName = "lblTechnique"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 5
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Water Render Method:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTechnique, UIControl))

            lCboLeft = lblTechnique.Left + lblTechnique.Width + 5

            Dim lblResolution As UILabel = New UILabel(MyBase.moUILib)
            With lblResolution
                .ControlName = "lblResolution"
                .Left = lstOptionSet.Left + lstOptionSet.Width + 10
                .Top = lstOptionSet.Top + 30
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Water Texture Resolution:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblResolution, UIControl))

            Dim cboWaterResolution As New UIComboBox(MyBase.moUILib)
            With cboWaterResolution
                .ControlName = "cboWaterResolution"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 30
                .Width = 155
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
                .ToolTipText = "Sets the resolution of the textures used for 2.0 shader water." & vbCrLf
            End With
            cboWaterResolution.Clear()
            cboWaterResolution.AddItem("Low Resolution")
            cboWaterResolution.ItemData(cboWaterResolution.NewIndex) = CInt(EngineSettings.eTextureResOptions.eVeryLowResTextures)
            cboWaterResolution.AddItem("Medium Resolution")
            cboWaterResolution.ItemData(cboWaterResolution.NewIndex) = CInt(EngineSettings.eTextureResOptions.eLowResTextures)
            cboWaterResolution.AddItem("High Resolution")
            cboWaterResolution.ItemData(cboWaterResolution.NewIndex) = CInt(EngineSettings.eTextureResOptions.eNormResTextures)

            cboWaterResolution.FindComboItemData(muSettings.WaterTextureRes)
            Me.AddChild(CType(cboWaterResolution, UIControl))
            AddHandler cboWaterResolution.ItemSelected, AddressOf ConfigureComboboxItemSelected

            Dim cboTechnique As New UIComboBox(MyBase.moUILib)
            With cboTechnique
                .ControlName = "cboTechnique"
                .Left = lCboLeft
                .Top = lstOptionSet.Top + 5
                .Width = 155
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
                .ToolTipText = "Sets the technique for rendering the water, lava and acid effects." & vbCrLf & _
                 "Plain will greatly improve performance on planets with water, lava and acid at a high cost of quality." & vbCrLf & _
                 "Shader 2.0 requires pixel shader 2.0 in order to use (recommended)."
            End With
            cboTechnique.Clear()
            cboTechnique.AddItem("Plain Water")
            cboTechnique.ItemData(cboTechnique.NewIndex) = 0
            If MyBase.moUILib.oDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
                cboTechnique.AddItem("Shader 2.0")
                cboTechnique.ItemData(cboTechnique.NewIndex) = 1
            End If
            cboTechnique.FindComboItemData(muSettings.WaterRenderMethod)
            Me.AddChild(CType(cboTechnique, UIControl))
            AddHandler cboTechnique.ItemSelected, AddressOf ConfigureComboboxItemSelected
        End Sub
#End Region
    End Class
#End Region

	Private Sub btnAliases_Click(ByVal sName As String) Handles btnAliases.Click
		If goCurrentPlayer Is Nothing Then Return
		Dim ofrm As frmAliasing = CType(goUILib.GetWindow("frmAliasing"), frmAliasing)
		If ofrm Is Nothing Then ofrm = New frmAliasing(goUILib)
		ofrm.Visible = True
	End Sub

    Private Sub SelfDestructNoticeResult(ByVal lRes As MsgBoxResult)
        If lRes = MsgBoxResult.Yes Then
            btnSelfDestruct.Caption = "Confirm"
        End If
    End Sub

	Private Sub btnSelfDestruct_Click(ByVal sName As String) Handles btnSelfDestruct.Click
		If goCurrentPlayer Is Nothing Then Return
		If gbAliased = True Then Return
        If btnSelfDestruct.Caption.ToUpper = "SELF DESTRUCT" Then
            Dim oMsgBox As New frmMsgBox(goUILib, "If you are in a Spawn System (indicated by (S) in the name), then you are under new spawn protection. If you respawn, you will not respawn into an (S) system and will lose the spawn protection." & vbCrLf & vbCrLf & "This cannot be undone. Do you wish to proceed?", MsgBoxStyle.YesNo, "Confirm")
            AddHandler oMsgBox.DialogClosed, AddressOf SelfDestructNoticeResult
            oMsgBox.Visible = True
        ElseIf btnSelfDestruct.Caption.ToUpper = "CANCEL SELF DESTRUCT" Then
            frmDeath.mbSelfDestruct = False
            Dim yMsg(2) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerRestart).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = 0
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            btnSelfDestruct.Caption = "Self Destruct"
        Else
            'do it
            frmDeath.mbSelfDestruct = True
            Dim yMsg(2) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerRestart).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = 1
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            btnSelfDestruct.Caption = "Cancel Self Destruct"
        End If
	End Sub

    Private Sub btnRestartTutorial_Click(ByVal sName As String) Handles btnRestartTutorial.Click
        If btnRestartTutorial.Caption.ToUpper = "CONFIRM" Then

            If goCurrentPlayer Is Nothing = False Then
                Dim sPName As String = goCurrentPlayer.PlayerName
                Try
                    If Exists(sPName & ".tut") = True Then Kill(sPName & ".tut")
                Catch
                End Try
            End If

            Dim yMsg(1) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRestartTutorial).CopyTo(yMsg, 0)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            frmMain.mbfrmConfirmHandled = True
            frmMain.Close()

            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            sPath &= "UpdaterClient.exe"
            If Exists(sPath) = True Then Shell(sPath, AppWinStyle.NormalFocus, False, -1)
            Application.Exit()
            End
        Else
            btnRestartTutorial.Caption = "Confirm"
        End If
    End Sub

    Private Sub btnClaim_Click(ByVal sName As String) Handles btnClaim.Click
        If goCurrentPlayer Is Nothing = False Then
            If goCurrentPlayer.yPlayerPhase = 255 Then
                Dim oFrm As frmClaim = CType(MyBase.moUILib.GetWindow("frmClaim"), frmClaim)
                If oFrm Is Nothing Then oFrm = New frmClaim(MyBase.moUILib)
                oFrm.Visible = True
                Return
            End If
        End If
        MyBase.moUILib.AddNotification("You must be in live in order to check your claimables.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub btnCredits_Click(ByVal sName As String) Handles btnCredits.Click
        Dim ofrm As New frmCredits(goUILib)
        ofrm.Visible = True
    End Sub

    Private Sub chkRaiseFullInvul_Click() Handles chkRaiseFullInvul.Click
        Dim yMsg(10) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eSetIronCurtain).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 6)

        If chkRaiseFullInvul.Value = True Then
            yMsg(10) = 1
            goCurrentPlayer.lStatusFlags = goCurrentPlayer.lStatusFlags Or elPlayerStatusFlag.AlwaysRaiseFullInvul
        Else
            yMsg(10) = 0
            If (goCurrentPlayer.lStatusFlags And elPlayerStatusFlag.AlwaysRaiseFullInvul) <> 0 Then
                goCurrentPlayer.lStatusFlags = goCurrentPlayer.lStatusFlags Xor elPlayerStatusFlag.AlwaysRaiseFullInvul
            End If
        End If

        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub frmOptions_WindowMoved() Handles Me.WindowMoved
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmAliasing")
        If ofrm Is Nothing = True Then
            Dim oofrm As UIWindow = MyBase.moUILib.GetWindow("frmAliasLogins")
            If oofrm Is Nothing = False Then
                oofrm.Left = Me.Left - oofrm.Width
                oofrm.Top = Me.Top
            End If
            oofrm = Nothing
        End If
        ofrm = Nothing

    End Sub
End Class

'Interface created from Interface Builder
Public Class fraColorSelect
	Inherits UIWindow

	Private lblRed As UILabel
	Private lblGreen As UILabel
    Private lblBlue As UILabel
    Private lblAlpha As UILabel
	Private WithEvents hscrRed As UIScrollBar
	Private txtSample As UITextBox
	Private WithEvents hscrGreen As UIScrollBar
    Private WithEvents hscrBlue As UIScrollBar
    Private WithEvents hscrAlpha As UIScrollBar
	Private WithEvents txtRed As UITextBox
	Private WithEvents txtGreen As UITextBox
    Private WithEvents txtBlue As UITextBox
    Private WithEvents txtAlpha As UITextBox
	Private WithEvents btnClose As UIButton

	Private mbIgnoreEvents As Boolean = True
	Private msKeyName As String
	Public Event ColorSelectClosed(ByVal clrValue As System.Drawing.Color, ByVal sKeyName As String)

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'fraColorSelect initial props
		With Me
			.ControlName = "fraColorSelect"
			.Left = 276
			.Top = 92
			.Width = 256
            .Height = 153 '128
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.BorderLineWidth = 2
			.Moveable = True
			.mbAcceptReprocessEvents = True
		End With

		'lblRed initial props
		lblRed = New UILabel(oUILib)
		With lblRed
			.ControlName = "lblRed"
			.Left = 10
			.Top = 10
			.Width = 25
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Red"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblRed, UIControl))

		'lblGreen initial props
		lblGreen = New UILabel(oUILib)
		With lblGreen
			.ControlName = "lblGreen"
			.Left = 10
			.Top = 35
			.Width = 40
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Green"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblGreen, UIControl))

		'lblBlue initial props
		lblBlue = New UILabel(oUILib)
		With lblBlue
			.ControlName = "lblBlue"
			.Left = 10
			.Top = 60
			.Width = 30
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Blue"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
        Me.AddChild(CType(lblBlue, UIControl))

        'lblAlpha initial props
        lblAlpha = New UILabel(oUILib)
        With lblAlpha
            .ControlName = "lblAlpha"
            .Left = 10
            .Top = 85
            .Width = 49
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Opacity"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAlpha, UIControl))

		'hscrRed initial props
		hscrRed = New UIScrollBar(oUILib, False)
		With hscrRed
			.ControlName = "hscrRed"
			.Left = 60
			.Top = 10
			.Width = 135
			.Height = 18
			.Enabled = True
			.Visible = True
			.Value = 0
			.MaxValue = 255
			.MinValue = 0
			.SmallChange = 1
			.LargeChange = 1
			.ReverseDirection = False
			.mbAcceptReprocessEvents = True
		End With
		Me.AddChild(CType(hscrRed, UIControl))

		'txtSample initial props
		txtSample = New UITextBox(oUILib)
		With txtSample
			.ControlName = "txtSample"
			.Left = 10
            .Top = 110
			.Width = 100
			.Height = 32
			.Enabled = False
			.Visible = True
			.Caption = "Sample"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.BackColorEnabled = System.Drawing.Color.FromArgb(255, 0, 0, 0)
			.BackColorDisabled = .BackColorEnabled
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.Locked = False
		End With
		Me.AddChild(CType(txtSample, UIControl))

		'hscrGreen initial props
		hscrGreen = New UIScrollBar(oUILib, False)
		With hscrGreen
			.ControlName = "hscrGreen"
			.Left = 60
			.Top = 35
			.Width = 135
			.Height = 18
			.Enabled = True
			.Visible = True
			.Value = 0
			.MaxValue = 255
			.MinValue = 0
			.SmallChange = 1
			.LargeChange = 1
			.ReverseDirection = False
			.mbAcceptReprocessEvents = True
		End With
		Me.AddChild(CType(hscrGreen, UIControl))

		'hscrBlue initial props
		hscrBlue = New UIScrollBar(oUILib, False)
		With hscrBlue
			.ControlName = "hscrBlue"
			.Left = 60
			.Top = 60
			.Width = 135
			.Height = 18
			.Enabled = True
			.Visible = True
			.Value = 0
			.MaxValue = 255
			.MinValue = 0
			.SmallChange = 1
			.LargeChange = 1
			.ReverseDirection = False
			.mbAcceptReprocessEvents = True
		End With
        Me.AddChild(CType(hscrBlue, UIControl))

        'hscrAlpha initial props
        hscrAlpha = New UIScrollBar(oUILib, False)
        With hscrAlpha
            .ControlName = "hscrAlpha"
            .Left = 60
            .Top = 85
            .Width = 135
            .Height = 18
            .Enabled = True
            .Visible = False
            .MaxValue = 255
            .MinValue = 0
            .Value = 255
            .SmallChange = 1
            .LargeChange = 1
            .ReverseDirection = False
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(hscrAlpha, UIControl))

		'txtRed initial props
		txtRed = New UITextBox(oUILib)
		With txtRed
			.ControlName = "txtRed"
			.Left = 205
			.Top = 9
			.Width = 40
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
		End With
		Me.AddChild(CType(txtRed, UIControl))

		'txtGreen initial props
		txtGreen = New UITextBox(oUILib)
		With txtGreen
			.ControlName = "txtGreen"
			.Left = 205
			.Top = 34
			.Width = 40
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
		End With
		Me.AddChild(CType(txtGreen, UIControl))

		'txtBlue initial props
		txtBlue = New UITextBox(oUILib)
		With txtBlue
			.ControlName = "txtBlue"
			.Left = 205
			.Top = 59
			.Width = 40
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
		End With
        Me.AddChild(CType(txtBlue, UIControl))

        'txtAlpha initial props
        txtAlpha = New UITextBox(oUILib)
        With txtAlpha
            .ControlName = "txtAlpha"
            .Left = 205
            .Top = 84
            .Width = 40
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "255"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
        End With
        Me.AddChild(CType(txtAlpha, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 145
            .Top = 111 '86
			.Width = 100
			.Height = 32
			.Enabled = True
			.Visible = True
			.Caption = "Close"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClose, UIControl))

		mbIgnoreEvents = False
	End Sub

	Public Sub SetColor(ByVal clrVal As System.Drawing.Color, ByVal sKeyName As String)
		mbIgnoreEvents = True
        msKeyName = sKeyName

		hscrBlue.Value = clrVal.B
		hscrGreen.Value = clrVal.G
        hscrRed.Value = clrVal.R
        hscrAlpha.Value = clrVal.A

		txtBlue.Caption = clrVal.B.ToString
		txtGreen.Caption = clrVal.G.ToString
        txtRed.Caption = clrVal.R.ToString
        txtAlpha.Caption = clrVal.A.ToString

        txtSample.ForeColor = clrVal

		mbIgnoreEvents = False
    End Sub

    Public Sub SetShowAlpha(ByVal bValue As Boolean)
        hscrAlpha.Visible = bValue
        txtAlpha.Visible = bValue
        lblAlpha.Visible = bValue
    End Sub

    Public Sub SetAlphaMin(ByVal lVal As Int32)
        hscrAlpha.MinValue = lVal
    End Sub


	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        Dim clrResultValue As System.Drawing.Color = System.Drawing.Color.FromArgb(hscrAlpha.Value, hscrRed.Value, hscrGreen.Value, hscrBlue.Value)
		Me.Visible = False
		RaiseEvent ColorSelectClosed(clrResultValue, msKeyName)
	End Sub

    Private Sub UpdateColorValue()
        Dim lA As Int32 = hscrAlpha.Value
        Dim lR As Int32 = hscrRed.Value
        Dim lG As Int32 = hscrGreen.Value
        Dim lB As Int32 = hscrBlue.Value

        txtSample.ForeColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
    End Sub

	Private Sub hscrBlue_ValueChange() Handles hscrBlue.ValueChange
		If mbIgnoreEvents = True Then Return
		mbIgnoreEvents = True
		txtBlue.Caption = hscrBlue.Value.ToString
		UpdateColorValue()
		mbIgnoreEvents = False
	End Sub

	Private Sub hscrGreen_ValueChange() Handles hscrGreen.ValueChange
		If mbIgnoreEvents = True Then Return
		mbIgnoreEvents = True
		txtGreen.Caption = hscrGreen.Value.ToString
		UpdateColorValue()
		mbIgnoreEvents = False
	End Sub

	Private Sub hscrRed_ValueChange() Handles hscrRed.ValueChange
		If mbIgnoreEvents = True Then Return
		mbIgnoreEvents = True
		txtRed.Caption = hscrRed.Value.ToString
		UpdateColorValue()
		mbIgnoreEvents = False
	End Sub

	Private Sub txtBlue_TextChanged() Handles txtBlue.TextChanged
		If mbIgnoreEvents = True Then Return
		mbIgnoreEvents = True
		Dim lTemp As Int32 = CInt(Val(txtBlue.Caption))
		If lTemp < 0 Then lTemp = 0
		If lTemp > 255 Then lTemp = 255
		txtBlue.Caption = lTemp.ToString
		hscrBlue.Value = lTemp
		UpdateColorValue()
		mbIgnoreEvents = False
	End Sub

	Private Sub txtGreen_TextChanged() Handles txtGreen.TextChanged
		If mbIgnoreEvents = True Then Return
		mbIgnoreEvents = True
		Dim lTemp As Int32 = CInt(Val(txtGreen.Caption))
		If lTemp < 0 Then lTemp = 0
		If lTemp > 255 Then lTemp = 255
		txtGreen.Caption = lTemp.ToString
		hscrGreen.Value = lTemp
		UpdateColorValue()
		mbIgnoreEvents = False
	End Sub

	Private Sub txtRed_TextChanged() Handles txtRed.TextChanged
		If mbIgnoreEvents = True Then Return
		mbIgnoreEvents = True
		Dim lTemp As Int32 = CInt(Val(txtRed.Caption))
		If lTemp < 0 Then lTemp = 0
		If lTemp > 255 Then lTemp = 255
		txtRed.Caption = lTemp.ToString
		hscrRed.Value = lTemp
		UpdateColorValue()
		mbIgnoreEvents = False
	End Sub

    Private Sub hscrAlpha_ValueChange() Handles hscrAlpha.ValueChange
        If mbIgnoreEvents = True Then Return
        mbIgnoreEvents = True
        txtAlpha.Caption = hscrAlpha.Value.ToString
        UpdateColorValue()
        mbIgnoreEvents = False
    End Sub

    Private Sub txtAlpha_TextChanged() Handles txtAlpha.TextChanged
        If mbIgnoreEvents = True Then Return
        mbIgnoreEvents = True
        Dim lTemp As Int32 = CInt(Val(txtAlpha.Caption))
        If lTemp < 0 Then lTemp = 0
        If lTemp > 255 Then lTemp = 255
        txtAlpha.Caption = lTemp.ToString
        hscrAlpha.Value = lTemp
        UpdateColorValue()
        mbIgnoreEvents = False
    End Sub
End Class