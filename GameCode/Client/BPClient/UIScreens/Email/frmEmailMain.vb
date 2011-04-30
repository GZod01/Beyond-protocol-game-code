Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmEmailMain
    Inherits UIWindow

    Public Shared lCurrentUnreadMessages As Int32 = 0

    Private WithEvents btnCompose As UIButton
    Private WithEvents btnContacts As UIButton
    Private WithEvents btnOptions As UIButton 
    Private lblFolders As UILabel
    Private WithEvents btnAddFolder As UIButton
    Private WithEvents lstFolders As UIListBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents lstEmails As UIListBox
    Private WithEvents mfraDetails As fraEmailDetail
    'Private WithEvents btnEmailIcon As UIButton
	Private WithEvents btnPurge As UIButton

    Private mlOriginalX As Int32
    Private mlOriginalY As Int32
    Private mlOriginalH As Int32
    Private mlOriginalW As Int32

    'Private mbStateShown As Boolean = False

    'Private mbFlickerState As Boolean = True
    'Private mlLastFlickerUpdate As Int32
    'Private mbHasNewEmail As Boolean = False

    Private mlLastDeleteHit As Int32 = 0
	Private mlLastUpdate As Int32 = 0 
	Private mlLastMsgUpdate As Int32 = 0
    Private Shared mlEmailsRequested As Int32 = -1

    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenEmailWindow)

        'frmEmailMain initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eEmailMain
            .ControlName = "frmEmailMain"
            '.Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 350
            '.Top = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 280

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.EmailMainX
                lTop = muSettings.EmailMainY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 350
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 280
            If lLeft + 700 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 700
            If lTop + 565 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 565

            .Left = lLeft
            .Top = lTop

            .Width = 700
            .Height = 565
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 1
            .Moveable = True
            .bRoundedBorder = False

            mlOriginalX = .Left
            mlOriginalY = .Top
            mlOriginalH = .Height
            mlOriginalW = .Width
        End With

        ''btnEmailIcon initial props
        'btnEmailIcon = New UIButton(oUILib)
        'With btnEmailIcon
        '    .ControlName = "btnEmailIcon"
        '    .Left = 0
        '    .Top = 0
        '    .Width = 32
        '    .Height = 16
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = CType(5, DrawTextFormat)
        '    .ControlImageRect = New Rectangle(161, 240, 31, 15)
        '    .ControlImageRect_Disabled = .ControlImageRect
        '    .ControlImageRect_Normal = .ControlImageRect
        '    .ControlImageRect_Pressed = .ControlImageRect
        'End With
        'Me.AddChild(CType(btnEmailIcon, UIControl))

        'btnCompose initial props
        btnCompose = New UIButton(oUILib)
        With btnCompose
            .ControlName = "btnCompose"
            .Left = 10
            .Top = 10
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Compose"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCompose, UIControl))

        'btnContacts initial props
        btnContacts = New UIButton(oUILib)
        With btnContacts
            .ControlName = "btnContacts"
            .Left = 120
            .Top = 10
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Contacts"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnContacts, UIControl))

        'btnOptions initial props
        btnOptions = New UIButton(oUILib)
        With btnOptions
            .ControlName = "btnOptions"
            .Left = 230
            .Top = 10
            .Width = 100
            .Height = 24
            .Enabled = Not gbAliased
            .Visible = True
            .Caption = "Options"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnOptions, UIControl))

        'lblFolders initial props
        lblFolders = New UILabel(oUILib)
        With lblFolders
            .ControlName = "lblFolders"
            .Left = 10
            .Top = 40
            .Width = 54
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Folders"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFolders, UIControl))

        'btnAddFolder initial props
        btnAddFolder = New UIButton(oUILib)
        With btnAddFolder
            .ControlName = "btnAddFolder"
            .Left = 90
            .Top = 39
            .Width = 60
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Add"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAddFolder, UIControl))

        'NewControl3 initial props
        lstFolders = New UIListBox(oUILib)
        With lstFolders
            .ControlName = "lstFolders"
            .Left = 10
            .Top = 65
            .Width = 140
			.Height = 130
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
		Me.AddChild(CType(lstFolders, UIControl))

		'btnPurge initial props
		btnPurge = New UIButton(oUILib)
		With btnPurge
			.ControlName = "btnPurge"
			.Left = 10
			.Top = lstFolders.Top + lstFolders.Height + 5
			.Width = 100
			.Height = 24
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Purge"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
			.ToolTipText = "Deletes all emails (read and unread) from the currently selected folder."
		End With
		Me.AddChild(CType(btnPurge, UIControl))

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
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'lstEmails initial props
        lstEmails = New UIListBox(oUILib)
        With lstEmails
            .ControlName = "lstEmails"
            .Left = 160
            .Top = 40
            .Width = 530
            .Height = 175
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstEmails, UIControl))

        'mfraDetail initial props
        mfraDetails = New fraEmailDetail(oUILib)
        With mfraDetails
            .Left = 10
            .Top = 225
            .Visible = True
        End With
        Me.AddChild(CType(mfraDetails, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewEmail) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to view the email interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
		End If
        Dim bHasMail As Boolean = False
        If mlEmailsRequested = glPlayerID Then
            For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
                If goCurrentPlayer.moEmailFolders(X).PlayerMsgs Is Nothing = False Then
                    bHasMail = True
                    Exit For
                End If
            Next
        End If

        If mlEmailsRequested <> glPlayerID OrElse bHasMail = False Then
            mlEmailsRequested = glPlayerID
            Dim yMsg(1) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestEmailSummarys).CopyTo(yMsg, 0)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End If

        Dim oTmpWin As frmQuickBar = CType(MyBase.moUILib.GetWindow("frmQuickBar"), frmQuickBar)
        If oTmpWin Is Nothing = False Then
            oTmpWin.ClearFlashState(1)
        End If
        oTmpWin = Nothing

        mbLoading = False
    End Sub

    Private Sub FillFolderList()
        Dim lCurrentPCF_ID As Int32 = -1

        If lstFolders.ListIndex > -1 Then
            lCurrentPCF_ID = lstFolders.ItemData(lstFolders.ListIndex)
        End If

        lstFolders.Clear()
        lstEmails.Clear()
        mfraDetails.cboMoveToFolder.Clear()

        For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
            If goCurrentPlayer.mlEmailFolderIdx(X) <> -1 Then
                goCurrentPlayer.moEmailFolders(X).bHasBeenSorted = False
                lstFolders.AddItem(goCurrentPlayer.moEmailFolders(X).FolderName)
                lstFolders.ItemData(lstFolders.NewIndex) = goCurrentPlayer.moEmailFolders(X).PCF_ID

                With mfraDetails.cboMoveToFolder
                    .AddItem(goCurrentPlayer.moEmailFolders(X).FolderName)
                    .ItemData(.NewIndex) = goCurrentPlayer.moEmailFolders(X).PCF_ID
                End With
            End If
        Next X

        If lCurrentPCF_ID <> -1 Then
            For X As Int32 = 0 To lstFolders.ListCount - 1
                If lstFolders.ItemData(X) = lCurrentPCF_ID Then
                    lstFolders.ListIndex = X
                    Exit For
                End If
            Next X
        Else
            For X As Int32 = 0 To lstFolders.ListCount - 1
                If lstFolders.ItemData(X) = PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF Then
                    lstFolders.ListIndex = X
                    Exit For
                End If
            Next
        End If
    End Sub

	Private Sub FillEmailList()
		btnPurge.Caption = "Purge"
		lstEmails.Clear()

		mfraDetails.ClearValues()

		If lstFolders.ListIndex = -1 Then Return

		Dim lPCF_ID As Int32 = lstFolders.ItemData(lstFolders.ListIndex)
		Dim oFolder As PlayerCommFolder = Nothing

		For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
			If goCurrentPlayer.mlEmailFolderIdx(X) = lPCF_ID Then
				oFolder = goCurrentPlayer.moEmailFolders(X)
				Exit For
			End If
		Next X

        If oFolder Is Nothing Then Return
        oFolder.SortEmails()

		For X As Int32 = 0 To oFolder.PlayerMsgUB
			If oFolder.PlayerMsgsIdx(X) <> -1 Then
				Dim sLine As String = oFolder.PlayerMsgs(X).sSender.PadRight(25, " "c)
				sLine &= oFolder.PlayerMsgs(X).MsgTitle
				lstEmails.AddItem(sLine, Not oFolder.PlayerMsgs(X).bMsgRead)
				lstEmails.ItemData(lstEmails.NewIndex) = oFolder.PlayerMsgs(X).ObjectID
			End If
		Next X

		If lstEmails.ListCount > 0 Then
			lstEmails.ListIndex = 0
		End If

	End Sub

	'Private Sub btnEmailIcon_Click(ByVal sName As String) Handles btnEmailIcon.Click
	'    SetState(True)
	'End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnCompose_Click(ByVal sName As String) Handles btnCompose.Click
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to create emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If

		Dim ofrm As New frmNewEmail(MyBase.moUILib)
		ofrm = Nothing
		If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.EmailComposeButtonClick)
	End Sub

	Private Sub btnAddFolder_Click(ByVal sName As String) Handles btnAddFolder.Click
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to view the create folders.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		Dim ofrmInputBox As New frmInputBox(goUILib)
		AddHandler ofrmInputBox.InputBoxClosed, AddressOf InputBoxResponse
	End Sub

	Private Sub btnContacts_Click(ByVal sName As String) Handles btnContacts.Click
		If HasAliasedRights(AliasingRights.eViewDiplomacy) = False Then
			MyBase.moUILib.AddNotification("You lack rights to view the contacts list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		Dim oAddrBook As New frmAddressBook(goUILib)
		oAddrBook = Nothing
		If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.EmailContactsButtonClick)
	End Sub

	Private Sub frmEmailMain_OnNewFrame() Handles Me.OnNewFrame

		'If we are presently minimized...
		'If mbStateShown = False Then
		'    If goCurrentPlayer Is Nothing = False Then

		'        If mbHasNewEmail = True Then
		'            If glCurrentCycle - mlLastFlickerUpdate > 15 Then
		'                mlLastFlickerUpdate = glCurrentCycle
		'                mbFlickerState = Not mbFlickerState

		'                If mbFlickerState = True Then
		'                    btnEmailIcon.ControlImageRect_Normal = New Rectangle(160, 240, 1, 1)
		'                Else
		'                    btnEmailIcon.ControlImageRect_Normal = New Rectangle(161, 240, 31, 15)
		'                End If
		'                btnEmailIcon.IsDirty = True
		'            End If
		'        Else
		'            mbFlickerState = False
		'            btnEmailIcon.ControlImageRect_Normal = New Rectangle(161, 240, 31, 15)
		'        End If
		'    End If
		'Else
		'We are maximized, check if there are any new items to render...
		If glCurrentCycle - mlLastUpdate > 30 Then
			mlLastUpdate = glCurrentCycle
			Dim lItemCnt As Int32 = 0
			For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
				If goCurrentPlayer.mlEmailFolderIdx(X) <> -1 Then lItemCnt += 1
			Next X
			If lItemCnt <> lstFolders.ListCount Then
				FillFolderList()
			Else
				Dim lFolderID As Int32 = -1
				If lstFolders.ListIndex > -1 Then lFolderID = lstFolders.ItemData(lstFolders.ListIndex)
				lItemCnt = 0
                If lFolderID <> -1 Then
                    For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
                        If goCurrentPlayer.mlEmailFolderIdx(X) = lFolderID Then
                            If goCurrentPlayer.moEmailFolders(X).bHasBeenSorted = False Then lItemCnt = -100000
                            goCurrentPlayer.moEmailFolders(X).SortEmails()
                            For Y As Int32 = 0 To goCurrentPlayer.moEmailFolders(X).PlayerMsgUB
                                If goCurrentPlayer.moEmailFolders(X).PlayerMsgsIdx(Y) <> -1 Then lItemCnt += 1
                            Next Y
                            Exit For
                        End If
                    Next X

                    If lItemCnt <> lstEmails.ListCount Then
                        Dim lEmailIdx As Int32 = lstEmails.ListIndex
                        FillEmailList()
                        If lEmailIdx <> -1 AndAlso lEmailIdx < lstEmails.ListCount Then
                            lstEmails.ListIndex = lEmailIdx
                            lstEmails.EnsureItemVisible(lEmailIdx)
                        End If
                    End If
                End If
			End If

			Dim oMsg As PlayerComm = GetCurrentEmailMsg()
			If oMsg Is Nothing = False Then
				If oMsg.lLastMsgUpdate <> mlLastMsgUpdate Then
                    lstEmails_ItemClick(lstEmails.ListIndex)
                    mlLastMsgUpdate = oMsg.lLastMsgUpdate
				End If
			End If
		End If
		'End If

	End Sub

	Private Sub InputBoxResponse(ByVal bCancelled As Boolean, ByVal sValue As String)
		If bCancelled = False AndAlso sValue <> "" Then

			'check the hardcodes
			Select Case sValue.Trim.ToUpper
				Case "INBOX", "OUTBOX", "DRAFTS", "DELETED"
					Return
				Case Else
					Dim yMsg(29) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eAddPlayerCommFolder).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 2)
					System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 6)
					If sValue.Length > 20 Then sValue = sValue.Substring(0, 20)
					System.Text.ASCIIEncoding.ASCII.GetBytes(sValue).CopyTo(yMsg, 10)
					MyBase.moUILib.SendMsgToPrimary(yMsg)
			End Select
		End If
	End Sub

	Private Function GetCurrentEmailMsg() As PlayerComm
		If lstEmails.ListIndex = -1 Then Return Nothing
		If lstFolders.ListIndex = -1 Then Return Nothing

		Dim oFolder As PlayerCommFolder = Nothing

		For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
			If goCurrentPlayer.mlEmailFolderIdx(X) = lstFolders.ItemData(lstFolders.ListIndex) Then
				oFolder = goCurrentPlayer.moEmailFolders(X)
				Exit For
			End If
		Next X

		If oFolder Is Nothing Then Return Nothing

		Dim oMsg As PlayerComm = Nothing
		For X As Int32 = 0 To oFolder.PlayerMsgUB
			If oFolder.PlayerMsgsIdx(X) = lstEmails.ItemData(lstEmails.ListIndex) Then
				oMsg = oFolder.PlayerMsgs(X)
				Exit For
			End If
		Next X

		Return oMsg
	End Function

	Private Sub lstEmails_ItemClick(ByVal lIndex As Integer) Handles lstEmails.ItemClick
		mfraDetails.ClearValues()

		Dim oMsg As PlayerComm = GetCurrentEmailMsg()
		If oMsg Is Nothing Then Return

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eEmailSelected, -1, -1, -1, "")
		End If

		If oMsg.bMsgRead = False Then
			oMsg.bMsgRead = True
			Dim yMsg(10) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eMarkEmailReadStatus).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(oMsg.ObjectID).CopyTo(yMsg, 2)
			System.BitConverter.GetBytes(oMsg.PCF_ID).CopyTo(yMsg, 6)
			yMsg(10) = 255
			MyBase.moUILib.SendMsgToPrimary(yMsg)
            lstEmails.ItemBold(lIndex) = False

            lCurrentUnreadMessages -= 1
            If lCurrentUnreadMessages < 0 Then lCurrentUnreadMessages = 0
		End If
		If oMsg.bRequestedDetails = False Then
			oMsg.bRequestedDetails = True
			Dim yMsg(9) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eRequestEmail).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(oMsg.PCF_ID).CopyTo(yMsg, 2)
			System.BitConverter.GetBytes(oMsg.ObjectID).CopyTo(yMsg, 6)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		End If

		mfraDetails.SetFromMsg(oMsg)
	End Sub

	Private Sub lstEmails_ItemDblClick(ByVal lIndex As Integer) Handles lstEmails.ItemDblClick
		Dim oMsg As PlayerComm = GetCurrentEmailMsg()
		If oMsg Is Nothing Then Return

		If lstFolders.ItemData(lstFolders.ListIndex) = PlayerCommFolder.ePCF_ID_HARDCODES.eDrafts_PCF Then
			Dim oFrmNewEmail As New frmNewEmail(goUILib)
			oFrmNewEmail.SetFromCurrentMsg(oMsg, frmNewEmail.eCreateFromMsgType.eEditDraft)
			oFrmNewEmail = Nothing
		End If
	End Sub

	Private Sub lstEmails_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstEmails.OnKeyDown
		Select Case e.KeyCode
            Case Keys.Delete
                BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eEmailDeleted, lstEmails.ItemData(lstEmails.ListIndex), ObjectType.ePlayerComm)

                If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
                    MyBase.moUILib.AddNotification("You lack rights to delete emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                    Return
                End If
                'delete the item...
                Dim yMsg(11) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eDeleteEmailItem).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(lstEmails.ItemData(lstEmails.ListIndex)).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(ObjectType.ePlayerComm).CopyTo(yMsg, 6)
                System.BitConverter.GetBytes(lstFolders.ItemData(lstFolders.ListIndex)).CopyTo(yMsg, 8)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
        End Select
	End Sub

	Private Sub btnOptions_Click(ByVal sName As String) Handles btnOptions.Click
		'TODO: Will probably want to expand this further in the future, but for now...
		If gbAliased = True Then
			goUILib.AddNotification("You cannot access the external email options of an aliased account.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		Dim ofrm As frmEmailSetup = CType(goUILib.GetWindow("frmEmailSetup"), frmEmailSetup)
		If ofrm Is Nothing Then ofrm = New frmEmailSetup(goUILib)
		ofrm.Visible = True
		ofrm = Nothing
		If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.EmailOptionsButtonClick)
	End Sub

	Private Sub lstFolders_ItemClick(ByVal lIndex As Integer) Handles lstFolders.ItemClick
		FillEmailList()
	End Sub

	Private Sub lstFolders_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstFolders.OnKeyDown

        Select Case e.KeyCode
            Case Keys.Delete
                If lstFolders.ListIndex <= 3 Then Return
                If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
                    MyBase.moUILib.AddNotification("You lack rights to delete folders.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                    Return
                End If
                If glCurrentCycle - mlLastDeleteHit > 90 Then   '3 seconds
                    goUILib.AddNotification("Press Delete again to confirm Folder Deletion. This action cannot be undone.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    mlLastDeleteHit = glCurrentCycle
                Else
                    Dim yMsg(7) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eDeleteEmailItem).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(lstFolders.ItemData(lstFolders.ListIndex)).CopyTo(yMsg, 2)
                    System.BitConverter.GetBytes(-1S).CopyTo(yMsg, 6)
                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                End If
        End Select
	End Sub

#Region "  Details Pane Events  "
	Private Sub mfraDetails_DeleteMsg() Handles mfraDetails.DeleteMsg
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to delete emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If

		Dim oMsg As PlayerComm = GetCurrentEmailMsg()
		If oMsg Is Nothing Then Return

		Dim yMsg(11) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eDeleteEmailItem).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(lstEmails.ItemData(lstEmails.ListIndex)).CopyTo(yMsg, 2)
		System.BitConverter.GetBytes(ObjectType.ePlayerComm).CopyTo(yMsg, 6)
		System.BitConverter.GetBytes(lstFolders.ItemData(lstFolders.ListIndex)).CopyTo(yMsg, 8)
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub mfraDetails_ForwardMsg() Handles mfraDetails.ForwardMsg
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to create emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		Dim oMsg As PlayerComm = GetCurrentEmailMsg()
		If oMsg Is Nothing Then Return
		Dim ofrmNewEmail As New frmNewEmail(goUILib)
		ofrmNewEmail.SetFromCurrentMsg(oMsg, frmNewEmail.eCreateFromMsgType.eForward)
		ofrmNewEmail = Nothing
	End Sub

	Private Sub mfraDetails_MoveToFolder(ByVal lPCF_ID As Integer) Handles mfraDetails.MoveToFolder
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to move emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		Dim oMsg As PlayerComm = GetCurrentEmailMsg()
		If oMsg Is Nothing Then Return

		Dim yMsg(15) As Byte

		System.BitConverter.GetBytes(GlobalMessageCode.eMoveEmailToFolder).CopyTo(yMsg, 0)
		oMsg.GetGUIDAsString.CopyTo(yMsg, 2)
		System.BitConverter.GetBytes(lstFolders.ItemData(lstFolders.ListIndex)).CopyTo(yMsg, 8)
		System.BitConverter.GetBytes(lPCF_ID).CopyTo(yMsg, 12)

		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub mfraDetails_Reply() Handles mfraDetails.Reply
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to create emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		Dim oMsg As PlayerComm = GetCurrentEmailMsg()
		If oMsg Is Nothing Then Return
		Dim ofrmNewEmail As New frmNewEmail(goUILib)
		ofrmNewEmail.SetFromCurrentMsg(oMsg, frmNewEmail.eCreateFromMsgType.eReply)
		ofrmNewEmail = Nothing
	End Sub

	Private Sub mfraDetails_ReplyToAll() Handles mfraDetails.ReplyToAll
		If HasAliasedRights(AliasingRights.eAlterEmail) = False Then
			MyBase.moUILib.AddNotification("You lack rights to create emails.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		Dim oMsg As PlayerComm = GetCurrentEmailMsg()
		If oMsg Is Nothing Then Return
		Dim ofrmNewEmail As New frmNewEmail(goUILib)
		ofrmNewEmail.SetFromCurrentMsg(oMsg, frmNewEmail.eCreateFromMsgType.eReplyToAll)
		ofrmNewEmail = Nothing
	End Sub
#End Region

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eEmail)
	End Sub

	Private Sub btnPurge_Click(ByVal sName As String) Handles btnPurge.Click
		If lstFolders.ListIndex > -1 Then
			If btnPurge.Caption.ToUpper = "CONFIRM" Then
				Dim yMsg(11) As Byte
				Dim lPos As Int32 = 0

				System.BitConverter.GetBytes(GlobalMessageCode.eDeleteEmailItem).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(ObjectType.ePlayerComm).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(lstFolders.ItemData(lstFolders.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4

                If lstFolders.ItemData(lstFolders.ListIndex) = PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF Then
                    frmEmailMain.lCurrentUnreadMessages = 0
                End If

				MyBase.moUILib.SendMsgToPrimary(yMsg)

				btnPurge.Caption = "Purge"
			Else : btnPurge.Caption = "Confirm"
			End If
		End If
	End Sub

    Private Sub frmEmailMain_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.EmailMainX = Me.Left
            muSettings.EmailMainY = Me.Top
        End If
    End Sub
End Class