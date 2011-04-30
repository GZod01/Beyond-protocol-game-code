Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTutorialStep
    Inherits UIWindow

    Private txtDisplayText As UITextBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnNext As UIButton

    Private WithEvents btnTakeMeThere As UIButton
    Private WithEvents btnINeedCredits As UIButton
    Private WithEvents btnReplay As UIButton
    Private WithEvents btnImStuck As UIButton

    Private WithEvents chkAutoTrigger As UICheckBox
    Private WithEvents btnPause As UIButton

    Private Shared mlSoundIdx As Int32 = -1     'should only be called once

    Private mbPaused As Boolean = False

    Private msSoundFile As String = ""

    Private mbIgnoreResize As Boolean = True

    'Private Shared mlMoverUsedWindowX As Int32 = -1
    'Private Shared mlMoverUserWindowY As Int32 = -1

    Private msAlert As String = ""

    Private mlCurrentStepID As Int32 = -1

    Private moSW As Stopwatch = Nothing

    Private mbSoundStarted As Boolean = False
    Private mlSoundStartCycle As Int32 = 0

    Private mlImStuckButtonCycle As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmTutorialStep initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTutorialStep
            .ControlName = "frmTutorialStep"
            .Left = 489
            .Top = 151
            .Width = 250
            .Height = 380
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 1
        End With

        'txtDisplayText initial props
        txtDisplayText = New UITextBox(oUILib)
        With txtDisplayText
            .ControlName = "txtDisplayText"
            .Left = 5
            .Top = 5
            .Width = 240
            .Height = 300
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
        Me.AddChild(CType(txtDisplayText, UIControl))

        ''btnClose initial props
        'btnClose = New UIButton(oUILib)
        'With btnClose
        '    .ControlName = "btnClose"
        '    .Left = 147
        '    .Top = 331
        '    .Width = 100
        '    .Height = 22
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Close"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = CType(5, DrawTextFormat)
        '    .ControlImageRect = New Rectangle(0, 0, 120, 32)
        'End With
        'Me.AddChild(CType(btnClose, UIControl))

        'btnPrevious initial props
        btnINeedCredits = New UIButton(oUILib)
        With btnINeedCredits
            .ControlName = "btnINeedCredits"
            .Left = 5
            .Top = 307
            .Width = 110
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Need Credits"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Requests a shot of 1,000,000 credits from the server." & vbCrLf & _
                           "This is only permitted during Tutorial Phase 1. The" & vbCrLf & _
                           "number of requests are counted and displayed to any guilds" & vbCrLf & _
                           "you apply to for membership as part of your overall score."
        End With
        Me.AddChild(CType(btnINeedCredits, UIControl))

        'btnNext initial props
        btnNext = New UIButton(oUILib)
        With btnNext
            .ControlName = "btnNext"
            .Left = 147
            .Top = 307
            .Width = 100
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Next"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnNext, UIControl))

        'btnTakeMeThere initial props
        btnTakeMeThere = New UIButton(oUILib)
        With btnTakeMeThere
            .ControlName = "btnTakeMeThere"
            .Left = 5
            .Top = 331
            .Width = 110
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Take Me There"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnTakeMeThere, UIControl))

        'btnImStuck initial props
        btnImStuck = New UIButton(oUILib)
        With btnImStuck
            .ControlName = "btnImStuck"
            .Left = 5
            .Top = 331
            .Width = 110
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "I'm Stuck"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnImStuck, UIControl))

        'chkAutoTrigger initial props
        chkAutoTrigger = New UICheckBox(oUILib)
        With chkAutoTrigger
            .ControlName = "chkAutoTrigger"
            .Left = 10
            .Top = btnTakeMeThere.Top + btnTakeMeThere.Height + 5
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Auto-Trigger"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "Check to allow the tutorial to progressively" & vbCrLf & _
                           "move through the tutorial as events occur in" & vbCrLf & _
                           "the game. Uncheck to cause the tutorial to" & vbCrLf & _
                           "wait for you to click Unpause to continue."
        End With
        Me.AddChild(CType(chkAutoTrigger, UIControl))

        'btnPause initial props
        btnPause = New UIButton(oUILib)
        With btnPause
            .ControlName = "btnPause"
            .Left = btnNext.Left
            .Top = btnTakeMeThere.Top
            .Width = 100
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Pause"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Stops/Resumes progression of the tutorial from moving to the next step."
        End With
        Me.AddChild(CType(btnPause, UIControl))

        ''btnReplay initial props
        'btnReplay = New UIButton(oUILib)
        'With btnReplay
        '    .ControlName = "btnReplay"
        '    .Left = 147
        '    .Top = 358
        '    .Width = 100
        '    .Height = 22
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Replay"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = CType(5, DrawTextFormat)
        '    .ControlImageRect = New Rectangle(0, 0, 120, 32)
        '    .ToolTipText = "Click to replay the audio for this tutorial step."
        'End With
        'Me.AddChild(CType(btnReplay, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mbIgnoreResize = False
    End Sub

    Public Sub SetFromStepNew(ByRef oStep As NewTutorialManager.ScriptStep)
        If oStep Is Nothing Then Return
        mbIgnoreResize = True

        mbPaused = True
        If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
        Try
            With oStep

                If .lImStuckDelay > 0 Then mlImStuckButtonCycle = .lImStuckDelay + glCurrentCycle Else mlImStuckButtonCycle = -1
                btnImStuck.Visible = False

                mlCurrentStepID = .StepID
                msAlert = .AlertControl

                Dim lX As Int32 = .DisplayLocX
                If lX = -1 Then
                    '-1 indicates center
                    lX = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
                ElseIf lX = -2 Then
                    '-2 indicates right align
                    lX = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
                End If
                Dim lY As Int32 = .DisplayLocY
                If lY = -1 Then
                    '-1 indicates center
                    lY = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
                ElseIf lY = -2 Then
                    '2 indicates bottom align
                    lY = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
                End If

                'If mlMoverUsedWindowX = -1 Then Me.Left = lX Else Me.Left = mlMoverUsedWindowX
                'If mlMoverUserWindowY = -1 Then Me.Top = lY Else Me.Top = mlMoverUserWindowY
                Me.Left = lX
                Me.Top = lY

                'txtDisplayText.Caption = "Step " & .StepID & ": " & .DisplayText
                If NewTutorialManager.lStepGroupNumber > 0 Then
                    txtDisplayText.Caption = NewTutorialManager.lStepGroupNumber.ToString & "-" & (.StepID - NewTutorialManager.lStepGroupNumberStartedOnStepID + 1).ToString & ": " & .DisplayText
                Else
                    txtDisplayText.Caption = .DisplayText
                End If
                txtDisplayText.ForceStartVisible()

                If goSound Is Nothing = False AndAlso .WAVFile <> "" Then
                    msSoundFile = .WAVFile
                    'mlSoundIdx = goSound.StartSound(.WAVFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)
                    mbSoundStarted = False
                    mlSoundStartCycle = glCurrentCycle + 10
                Else
                    moSW = Stopwatch.StartNew
                End If
            End With
        Catch
        End Try

        mbPaused = Not chkAutoTrigger.Value
        If mbPaused = True Then
            btnPause.Caption = "Unpause"
        Else
            btnPause.Caption = "Pause"
        End If
        mbIgnoreResize = False
    End Sub

    'Public Sub SetFromStep(ByVal uStep As TutorialManager.TutorialStep)
    '    mbIgnoreResize = True

    '    If mlSoundIdx <> -1 Then goSound.StopSound(mlSoundIdx)

    '    With uStep

    '        Dim lX As Int32 = .lDisplayLocX
    '        If lX = -1 Then
    '            '-1 indicates center
    '            lX = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
    '        ElseIf lX = -2 Then
    '            '-2 indicates right align
    '            lX = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
    '        End If
    '        Dim lY As Int32 = .lDisplayLocY
    '        If lY = -1 Then
    '            '-1 indicates center
    '            lY = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
    '        ElseIf lY = -2 Then
    '            '2 indicates bottom align
    '            lY = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
    '        End If

    '        If mlMoverUsedWindowX = -1 Then Me.Left = lX Else Me.Left = mlMoverUsedWindowX
    '        If mlMoverUserWindowY = -1 Then Me.Top = lY Else Me.Top = mlMoverUserWindowY

    '        txtDisplayText.Caption = .sDisplayText

    '        If goSound Is Nothing = False AndAlso .sWAVFile <> "" Then
    '            msSoundFile = .sWAVFile
    '            mlSoundIdx = goSound.StartSound(.sWAVFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)
    '        End If
    '    End With

    '    mbIgnoreResize = False
    'End Sub

    'Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
    '    Me.Visible = False
    '    mbPaused = True
    '    Try
    '        If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
    '    Catch
    '    End Try
    '    goTutorial.CancelCustomStepType()
    '    mbPaused = False
    'End Sub

    Private Sub frmTutorialStep_OnNewFrame() Handles Me.OnNewFrame
        If mbSoundStarted = False Then
            If mlSoundStartCycle <= glCurrentCycle Then
                mbSoundStarted = True
                If goSound Is Nothing = False Then mlSoundIdx = goSound.StartSound(msSoundFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)
            Else
                Return
            End If
        End If

        If mlImStuckButtonCycle > -1 AndAlso glCurrentCycle > mlImStuckButtonCycle Then
            If btnImStuck.Visible = False Then btnImStuck.Visible = True
        End If

        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
            If btnINeedCredits.Enabled = False Then btnINeedCredits.Enabled = False
        ElseIf btnINeedCredits.Enabled = True Then
            btnINeedCredits.Enabled = False
        End If

        If NewTutorialManager.bStepHasCenterScreen <> btnTakeMeThere.Enabled Then btnTakeMeThere.Enabled = NewTutorialManager.bStepHasCenterScreen

        If mlSoundIdx <> -1 AndAlso mbPaused = False AndAlso goSound Is Nothing = False Then
            If goSound.IsSoundStopped(mlSoundIdx) = True Then
                mlSoundIdx = -1
                If NewTutorialManager.GetTutorialStepID = 317 Then
                    MyBase.moUILib.RemoveWindow(Me.ControlName)
                    Return
                End If
                NewTutorialManager.TutorialStepFinished()
            End If
        ElseIf moSW Is Nothing = False AndAlso moSW.ElapsedMilliseconds > 60000 Then
            moSW = Nothing
            If NewTutorialManager.GetTutorialStepID = 317 Then
                MyBase.moUILib.RemoveWindow(Me.ControlName)
                Return
            End If
            NewTutorialManager.TutorialStepFinished()
        End If

        'If mbPaused = True Then
        '    If btnPause.Caption <> "Unpause" Then btnPause.Caption = "Unpause"
        'Else
        '    If btnPause.Caption <> "Pause" Then btnPause.Caption = "Pause"
        'End If

        If (goTutorial Is Nothing OrElse goTutorial.TutorialOn = False) AndAlso (NewTutorialManager.TutorialOn = False) Then MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnNext_Click(ByVal sName As String) Handles btnNext.Click
        If NewTutorialManager.TutorialOn = True Then

            If NewTutorialManager.GetTutorialStepID = 317 Then
                MyBase.moUILib.RemoveWindow(Me.ControlName)
                Return
            End If

            If mlCurrentStepID = 94 Then
                Dim oWin As frmArmorBuilder = CType(MyBase.moUILib.GetWindow("frmArmorBuilder"), frmArmorBuilder)
                If oWin Is Nothing = False Then
                    If oWin.ValidateTutorialConfig() = False Then
                        MyBase.moUILib.AddNotification("Some entries are not valid. Please correct and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                End If
            End If

            If mlSoundIdx > -1 Then
                If goSound Is Nothing = False Then
                    Try
                        goSound.StopSound(mlSoundIdx)
                    Catch
                    End Try
                End If
            End If

            If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then
                Dim lStepID As Int32 = NewTutorialManager.GetTutorialStepID()
                NewTutorialManager.FindAndExecuteStepID(lStepID + 1)
                Return
            End If

            If NewTutorialManager.TutorialStepFinished() = False Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eTutorialStepNextClick, -1, -1, -1, "")
            End If
        Else
            If goTutorial Is Nothing = False Then goTutorial.ForcefullyMoveNext()
        End If

        'mbPaused = False
    End Sub

    Private Sub btnPause_Click(ByVal sName As String) Handles btnPause.Click
        mbPaused = Not mbPaused
        If mbPaused = True Then
            btnPause.Caption = "Unpause"
            Try
                If goSound Is Nothing = False AndAlso mlSoundIdx <> -1 Then goSound.StopSound(mlSoundIdx)
            Catch
            End Try
        Else
            btnPause.Caption = "Pause"

            'If goSound Is Nothing = False AndAlso msSoundFile <> "" Then
            '    mlSoundIdx = goSound.StartSound(msSoundFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)
            'End If
        End If
    End Sub

    Private Sub btnINeedCredits_Click(ByVal sName As String) Handles btnINeedCredits.Click
        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eTutorialGiveCredits).CopyTo(yMsg, 0)
        MyBase.moUILib.SendMsgToPrimary(yMsg)
        MyBase.moUILib.AddNotification("Requesting 1,000,000 credits from server.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    'Private Sub chkAutoTrigger_Click() Handles chkAutoTrigger.Click
    '	goTutorial.bAutoTrigger = chkAutoTrigger.Value
    'End Sub

	Private Sub btnReplay_Click(ByVal sName As String) Handles btnReplay.Click
		If msAlert Is Nothing = False AndAlso msAlert <> "" Then
			NewTutorialManager.AlertControl(msAlert)
		End If
		If msSoundFile <> "" Then
            'Dim bOriginalPause As Boolean = mbPaused
            'mbPaused = True

			If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then
				Try
					goSound.StopSound(mlSoundIdx)
				Catch
				End Try
			End If
            mlSoundIdx = goSound.StartSound(msSoundFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)

            'mbPaused = bOriginalPause
		End If
	End Sub

    Protected Overrides Sub Finalize()
		Try
			If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
		Catch
		End Try
		MyBase.Finalize()
    End Sub

    Private Sub btnTakeMeThere_Click(ByVal sName As String) Handles btnTakeMeThere.Click
        NewTutorialManager.ExecuteCenterScreenCmd()
    End Sub

    Private Sub chkAutoTrigger_Click() Handles chkAutoTrigger.Click
        If chkAutoTrigger.Value = False Then
            mbPaused = True
            btnPause.Caption = "Unpause"
        End If
    End Sub

    'Private Sub frmTutorialStep_WindowMoved() Handles Me.WindowMoved
    '    If mbIgnoreResize = True Then Return
    '    mlMoverUsedWindowX = Me.Left
    '    mlMoverUserWindowY = Me.Top
    'End Sub

    Private Sub btnImStuck_Click(ByVal sName As String) Handles btnImStuck.Click
        If NewTutorialManager.TutorialOn = True Then
            Dim lStepID As Int32 = NewTutorialManager.GetTutorialStepID()
            NewTutorialManager.FindAndExecuteStepID(lStepID + 1)
        End If
    End Sub
End Class