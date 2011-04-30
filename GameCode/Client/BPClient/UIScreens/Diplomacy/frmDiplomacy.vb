Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmDiplomacy
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnExport As UIButton
    Private lnDiv1 As UILine
    Private mfraList As fraPlayerRelList
    Private mfraDetails As fraPlayerDetails
    Private mfraMyDetails As fraMyDipDetails

    Private mlLastUpdate As Int32 = 0

    Private mbFirstTime As Boolean = True
    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenDiplomacyWindow)

        'frmDiplomacy initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eDiplomacy
            .ControlName = "frmDiplomacy"
            '.Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 370
            '.Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.DiplomacyX
                lTop = muSettings.DiplomacyY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 370
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 310
            If lLeft + 745 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 745
            If lTop + 620 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 620

            .Left = lLeft
            .Top = lTop

            .Width = 745
            .Height = 620 '600
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 1
            .mbAcceptReprocessEvents = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 0
            .Width = 600
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Foreign Policy for Magistrate Enoch Dagor"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            '.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            '.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnExport
        btnExport = New UIButton(oUILib)
        With btnExport
            .ControlName = "btnExport"
            .Left = btnHelp.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "E"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            '.ToolTipText = "You can change the format of exported data from CSV to XML " & vbCrLf & "in the general settings window." & vbCrLf & "Exported data will go into your GameDir\ExportedData"
            .ToolTipText = "Click to export data." & vbCrLf & "Exported data will go into your GameDir\ExportedData"
        End With
        Me.AddChild(CType(btnExport, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 0
            .Top = 25
            .Width = 745
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

		'fraDiplomacyList initial props
		mfraList = New fraPlayerRelList(oUILib)
		With mfraList
			.Left = 5
			.Top = 30
			.Width = 735
            .Height = 295 '280
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.FillWindow = False
		End With
		Me.AddChild(CType(mfraList, UIControl))

		'fraPlayerDetails initial props
		mfraDetails = New fraPlayerDetails(moUILib)
		With mfraDetails
			.Left = 5
            .Top = 335 '315
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
		End With
		Me.AddChild(CType(mfraDetails, UIControl))

		'mfraMyDetails initial props
		mfraMyDetails = New fraMyDipDetails(moUILib)
		With mfraMyDetails
            .Left = mfraDetails.Left + mfraDetails.Width + 5
            .Top = 335 '315
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
		End With
		Me.AddChild(CType(mfraMyDetails, UIControl))

        AddHandler mfraList.ListItemClicked, AddressOf ItemClicked
        AddHandler mfraList.ItemDataChanged, AddressOf ItemRelChanged

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewAgents Or AliasingRights.eViewDiplomacy) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            goUILib.AddNotification("You lack the rights to view the diplomatic relations interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
		End If

		Dim oTmpWin As frmQuickBar = CType(MyBase.moUILib.GetWindow("frmQuickBar"), frmQuickBar)
		If oTmpWin Is Nothing = False Then
			oTmpWin.ClearFlashState(4)
		End If
        oTmpWin = Nothing

        mbLoading = False
		'glCurrentEnvirView = CurrentView.eDiplomacyScreen
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow("frmIconChooser")
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		'ReturnToPreviousView()
	End Sub

	Public Sub SetFromCurrentPlayer()
		lblTitle.Caption = "Foreign Policy for " & Player.GetPlayerNameWithTitle(goCurrentPlayer.yPlayerTitle, goCurrentPlayer.PlayerName, goCurrentPlayer.bIsMale)
		mfraList.SetFromCurrentPlayer()
	End Sub

    Private Sub ItemClicked(ByVal lPlayerID As Int32)

        If mfraDetails.GetCurrentRelID <> lPlayerID Then
            mfraList.ClearAllItemsTempScore()
        End If

        Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lPlayerID)
        mfraDetails.SetData(oRel)
    End Sub

    Private Sub ItemRelChanged(ByVal lPlayerID As Int32, ByVal lRelVal As Int32)
        If mfraDetails.GetCurrentRelID() <> lPlayerID Then
            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lPlayerID)
            mfraDetails.SetData(oRel)
        End If
        mfraDetails.SetExternalScore(lRelVal)
    End Sub

	Private Sub frmDiplomacy_OnNewFrame() Handles Me.OnNewFrame
		If glCurrentCycle - mlLastUpdate > 300 Then
			Dim yMsg(5) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eGetPlayerScores).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
			mlLastUpdate = glCurrentCycle
		End If

		If mbFirstTime = True AndAlso glCurrentCycle - mlLastUpdate > 60 Then
			mbFirstTime = False
			Me.IsDirty = True
		End If

		mfraList.fraPlayerRelList_OnNewFrame()
		mfraMyDetails.frmMyDipDetails_OnNewFrame()
		mfraDetails.fraPlayerDetails_OnNewFrame()
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eDiplomacy)
	End Sub

    Private Sub frmDiplomacy_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.DiplomacyX = Me.Left
            muSettings.DiplomacyY = Me.Top
        End If
    End Sub

    Private Sub btnExport_Click(ByVal sName As String) Handles btnExport.Click
        mfraList.Export_DiplomacyInfo()
    End Sub
End Class