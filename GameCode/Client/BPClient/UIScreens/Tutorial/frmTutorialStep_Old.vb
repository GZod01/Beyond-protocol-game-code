Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTutorialStep_Old
    Inherits UIWindow

    Private txtDisplayText As UITextBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnNext As UIButton

    Private WithEvents btnPause As UIButton
    Private WithEvents btnPrevious As UIButton
    Private WithEvents btnReplay As UIButton

    Private WithEvents chkAutoTrigger As UICheckBox

    Private Shared mlSoundIdx As Int32 = -1     'should only be called once

    Private mbPaused As Boolean = False

    Private msSoundFile As String = ""

    Private mbIgnoreResize As Boolean = True

    Private Shared mlMoverUsedWindowX As Int32 = -1
    Private Shared mlMoverUserWindowY As Int32 = -1

    Private msAlert As String = ""

    Private mlCurrentStepID As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmTutorialStep initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTutorialOld
            .ControlName = "frmTutorialStep_Old"
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

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 147
            .Top = 331
            .Width = 100
            .Height = 22
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

        'btnPrevious initial props
        btnPrevious = New UIButton(oUILib)
        With btnPrevious
            .ControlName = "btnPrevious"
            .Left = 5
            .Top = 307
            .Width = 100
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Previous"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnPrevious, UIControl))

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

        'btnPause initial props
        btnPause = New UIButton(oUILib)
        With btnPause
            .ControlName = "btnPause"
            .Left = 5
            .Top = 331
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
        End With
        Me.AddChild(CType(btnPause, UIControl))

        'chkAutoTrigger initial props
        chkAutoTrigger = New UICheckBox(oUILib)
        With chkAutoTrigger
            .ControlName = "chkAutoTrigger"
            .Left = 10
            .Top = 360
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
                           "wait for you to click NEXT to continue."
        End With
        Me.AddChild(CType(chkAutoTrigger, UIControl))

        'btnReplay initial props
        btnReplay = New UIButton(oUILib)
        With btnReplay
            .ControlName = "btnReplay"
            .Left = 147
            .Top = 358
            .Width = 100
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Replay"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to replay the audio for this tutorial step."
        End With
        Me.AddChild(CType(btnReplay, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mbIgnoreResize = False
    End Sub

    Public Sub SetFromStep(ByVal uStep As TutorialManager.TutorialStep)
        mbIgnoreResize = True

        If mlSoundIdx <> -1 Then goSound.StopSound(mlSoundIdx)

        With uStep

            Dim lX As Int32 = .lDisplayLocX
            If lX = -1 Then
                '-1 indicates center
                lX = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
            ElseIf lX = -2 Then
                '-2 indicates right align
                lX = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
            End If
            Dim lY As Int32 = .lDisplayLocY
            If lY = -1 Then
                '-1 indicates center
                lY = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
            ElseIf lY = -2 Then
                '2 indicates bottom align
                lY = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
            End If

            If mlMoverUsedWindowX = -1 Then Me.Left = lX Else Me.Left = mlMoverUsedWindowX
            If mlMoverUserWindowY = -1 Then Me.Top = lY Else Me.Top = mlMoverUserWindowY

            txtDisplayText.Caption = .sDisplayText

            If goSound Is Nothing = False AndAlso .sWAVFile <> "" Then
                msSoundFile = .sWAVFile
                mlSoundIdx = goSound.StartSound(.sWAVFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)
            End If
        End With

        mbIgnoreResize = False
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        Me.Visible = False
        mbPaused = True
        Try
            If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
        Catch
        End Try
        goTutorial.CancelCustomStepType()
        mbPaused = False
        msSoundFile = ""
    End Sub

    Private Sub frmTutorialStep_OnNewFrame() Handles Me.OnNewFrame
        If mlSoundIdx <> -1 AndAlso mbPaused = False AndAlso chkAutoTrigger.Value = True Then
            If goSound Is Nothing = False Then
                If goSound.IsSoundStopped(mlSoundIdx) = True Then
                    mlSoundIdx = -1
                    goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.TutorialStepEnded)
                    'NewTutorialManager.TutorialStepFinished()
                End If
            End If
        End If

        If mbPaused = True Then
            If btnPause.Caption <> "Unpause" Then btnPause.Caption = "Unpause"
        Else
            If btnPause.Caption <> "Pause" Then btnPause.Caption = "Pause"
        End If

        If (goTutorial Is Nothing OrElse goTutorial.TutorialOn = False) Then MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnNext_Click(ByVal sName As String) Handles btnNext.Click
        If goTutorial Is Nothing = False Then goTutorial.ForcefullyMoveNext()
        mbPaused = False
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

            If goSound Is Nothing = False AndAlso msSoundFile <> "" Then
                Try
                    If goSound Is Nothing = False AndAlso mlSoundIdx <> -1 Then goSound.StopSound(mlSoundIdx)
                Catch
                End Try
                mlSoundIdx = goSound.StartSound(msSoundFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub btnPrevious_Click(ByVal sName As String) Handles btnPrevious.Click
        If goTutorial Is Nothing = False Then goTutorial.MovePrevious()
    End Sub

    Private Sub frmTutorialStep_OnResize() Handles Me.OnResize
        If mbIgnoreResize = True Then Return
        mlMoverUsedWindowX = Me.Left
        mlMoverUserWindowY = Me.Top
    End Sub

    Private Sub chkAutoTrigger_Click() Handles chkAutoTrigger.Click
        goTutorial.bAutoTrigger = chkAutoTrigger.Value
    End Sub

    Private Sub btnReplay_Click(ByVal sName As String) Handles btnReplay.Click
        If msSoundFile <> "" Then
            Dim bOriginalPause As Boolean = mbPaused
            mbPaused = True

            If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then
                Try
                    goSound.StopSound(mlSoundIdx)
                Catch
                End Try
            End If
            mlSoundIdx = goSound.StartSound(msSoundFile, False, SoundMgr.SoundUsage.eTutorialVoice, Nothing, Nothing)

            mbPaused = bOriginalPause
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Try
            If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
        Catch
        End Try
        MyBase.Finalize()
    End Sub
End Class

'Interface created from Interface Builder
Public Class frmHelpItem
    Inherits UIWindow

    Private txtDisplayText As UITextBox

    Private WithEvents btnClose As UIButton
    Private WithEvents btnNext As UIButton
    Private WithEvents btnPrevious As UIButton

    Private mbIgnoreResize As Boolean = True

    Private Shared mlMoverUsedWindowX As Int32 = -1
    Private Shared mlMoverUserWindowY As Int32 = -1

    Private moTopic As TOCList.TOCItem = Nothing
    Private mlPage As Int32 = -1
 
    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmTutorialStep initial props
        With Me
            .ControlName = "frmHelpItem"
            .Left = 489
            .Top = 151
            .Width = 250
            .Height = 511 '380
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 1
        End With

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 147
            .Top = Me.Height - 25 '331
            .Width = 100
            .Height = 25
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

        'btnPrevious initial props
        btnPrevious = New UIButton(oUILib)
        With btnPrevious
            .ControlName = "btnPrevious"
            .Left = 5

            .Width = 100
            .Height = 25
            .Top = btnClose.Top - .Height - 1 '307
            .Enabled = True
            .Visible = True
            .Caption = "Previous"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnPrevious, UIControl))

        'btnNext initial props
        btnNext = New UIButton(oUILib)
        With btnNext
            .ControlName = "btnNext"
            .Left = 147
            .Top = 438 '307
            .Width = 100
            .Height = 25
            .Top = btnClose.Top - .Height - 1 '307
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



        'txtDisplayText initial props
        txtDisplayText = New UITextBox(oUILib)
        With txtDisplayText
            .ControlName = "txtDisplayText"
            .Left = 5
            .Top = 5
            .Width = 240
            .Height = btnNext.Top - .Top - 1 - 1 '300
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

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mbIgnoreResize = False
    End Sub

    Public Sub SetFromTopicNode(ByVal oTopic As TOCList.TOCItem)
        mbIgnoreResize = True
        moTopic = oTopic
        mlPage = -1
        'If mlMoverUsedWindowX = -1 Then Me.Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - Me.Width Else Me.Left = mlMoverUsedWindowX
        'If mlMoverUserWindowY = -1 Then Me.Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - Me.Height Else Me.Top = mlMoverUserWindowY
        mbIgnoreResize = False

        btnNext_Click("btnNext")
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        Me.Visible = False
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Public Sub ChangePage(ByVal lDir As Int32)
        If moTopic Is Nothing OrElse moTopic.oItems Is Nothing Then Return

        Try
            With moTopic
                mlPage += lDir
                If mlPage > .oItems.GetUpperBound(0) Then mlPage = .oItems.GetUpperBound(0)
                If mlPage < 0 Then mlPage = 0

                btnNext.Enabled = mlPage < .oItems.GetUpperBound(0)
                btnPrevious.Enabled = mlPage > 0

                txtDisplayText.Caption = .oItems(mlPage).sText

                If lDir > 0 Then
                    txtDisplayText.ForceStartVisible()
                Else
                    txtDisplayText.ForceEndVisible()
                End If

            End With
        Catch
        End Try
    End Sub

    Private Sub btnNext_Click(ByVal sName As String) Handles btnNext.Click
        ChangePage(1)
    End Sub

    Private Sub btnPrevious_Click(ByVal sName As String) Handles btnPrevious.Click
        ChangePage(-1)
    End Sub

    Private Sub frmHelpItem_WindowMoved() Handles Me.WindowMoved
        mlMoverUsedWindowX = Me.Left
        mlMoverUserWindowY = Me.Top

        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmTutorialTOC")
        If ofrm Is Nothing = False Then
            ofrm.Top = Me.Top
            ofrm.Left = Me.Left - ofrm.Width
        End If
    End Sub
End Class