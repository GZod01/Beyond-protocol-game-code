Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Enum eyChatRoomCommandType As Byte
    AddNewChatRoom = 1
    LeaveChatRoom = 2
    ToggleAdminRights = 4
    JoinChannel = 8
    SetChannelPassword = 16
    KickPlayer = 32
    InvitePlayer = 64
    SetChannelPublic = 128
End Enum

'Chat Window
Public Class frmChat
    Inherits UIWindow

    Private Structure OriginalMsg
        Public sText As String
        Public lPlayerID As Int32
        Public yType As ChatMessageType
        Public dtDate As Date
    End Structure
    Private Shared muOriginalTexts() As OriginalMsg

    Private txtTabs() As UITextBox
    Private moTabs() As ChatTab
    Private mlTabUB As Int32 = -1

    Private WithEvents txtNew As UITextBox
    Private WithEvents scrText As UIScrollBar
	Private lblText() As UILabel
	'Private txtText As UITextBox
    Private lnDiv As UILine

    Private WithEvents btnAdd As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnToggle As UIButton

    Private mlCurrentTabIndex As Int32 = 0

    Private mlPreviousTabClick As Int32 = 0
    Private mlPreviousTabClickIdx As Int32 = 0

    Private mlLastTabFlash As Int32 = 0
    Private mbFlashColor As Boolean = True
    Private mbHasFlash As Boolean = False

    Private mbForceNextRefresh As Boolean = False
	Private mlForcefulRefresh As Int32

	Private msIgnoreList() As String
    Private mlIgnoreListUB As Int32 = -1

    Private clrFlashOn As System.Drawing.Color
    Private clrSelected As System.Drawing.Color
    Private clrFlashOff As System.Drawing.Color

    Private mbMinimized As Boolean = False
    Private mlNormalHeight As Int32
    Private mbIgnoreResize As Boolean = True

    Private mlLastMsgCycle As Int32
    Private mlLastRenderCycle As Int32

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmChat initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eChatWindow
            .ControlName = "frmChat"
            .Left = 590 'Should be right next to frmContents
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - 156     'bottom align
            If NewTutorialManager.TutorialOn = False Then
                If muSettings.ChatWindowLocX <> -1 Then .Left = muSettings.ChatWindowLocX
                If muSettings.ChatWindowLocY <> -1 Then .Top = muSettings.ChatWindowLocY
            End If

            If muSettings.ChatWindowWidth < 128 Then .Width = 384 Else .Width = muSettings.ChatWindowWidth
            If muSettings.ChatWindowHeight < 64 Then .Height = 156 Else .Height = muSettings.ChatWindowHeight
            '.Width = 384
            '.Height = 156
            .Enabled = True
            .Visible = True

            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height

            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(192, 0, 32, 64)
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True  'always render the chat interface
            .mbAcceptReprocessEvents = True
            .BorderLineWidth = 1
            .bRoundedBorder = False
            .Resizable = True
            .ResizeLimits = New System.Drawing.Rectangle(128, 64, 1024, 1024)
        End With

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = Me.Width - 19
            .Top = 1
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            '.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Remove the current chat tab."
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'btnAdd initial props
        btnAdd = New UIButton(oUILib)
        With btnAdd
            .ControlName = "btnAdd"
            .Left = btnDelete.Left - 19
            .Top = 1
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "+"
            .ForeColor = muSettings.InterfaceBorderColor
            '.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Adds a new tab to the chat window."
        End With
        Me.AddChild(CType(btnAdd, UIControl))

        'btnToggle initial props
        btnToggle = New UIButton(oUILib)
        With btnToggle
            .ControlName = "btnToggle"
            .Left = btnAdd.Left - 38
            .Top = btnAdd.Top
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
            .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled)
            .ControlImageRect_Normal = .ControlImageRect
            .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eDownArrow_Button_Down)
            .ToolTipText = "Click to minimize window or maximize window."
        End With
        Me.AddChild(CType(btnToggle, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnAdd.Left - 19
            .Top = btnAdd.Top
            .Width = 18
            .Height = 18
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



        'txtNew initial props
        txtNew = New UITextBox(oUILib)
        With txtNew
            .ControlName = "txtNew"
            .Left = 0
            .Top = 137
            .Width = 381
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            '.ForeColor = System.Drawing.Color.FromArgb(-16777216)
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            '.BackColorEnabled = System.Drawing.Color.FromArgb(-1)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 0
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtNew, UIControl))

        'lblText initial props
        Dim X As Int32
        ReDim lblText(5)
        For X = 0 To lblText.Length - 1
            lblText(X) = New UILabel(oUILib)
            With lblText(X)
                .ControlName = "lblText(" & X & ")"
                .Left = 3
                .Top = txtNew.Top - ((X + 1) * 18)
                .Width = 365
                .Height = 17
                .Enabled = True
                .Visible = True
                .Caption = ""
                '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.Bottom Or DrawTextFormat.Left
            End With
            Me.AddChild(CType(lblText(X), UIControl))
        Next X

        'lnDiv initial props
        lnDiv = New UILine(oUILib)
        With lnDiv
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .ControlName = "lnDiv"
            .Enabled = True
            .Height = 0
            .Width = Me.Width
            .Left = 0
            .Top = 20
            .Visible = True
        End With
        Me.AddChild(CType(lnDiv, UIControl))

        'scrText initial props
        scrText = New UIScrollBar(oUILib, True)
        With scrText
            .ControlName = "scrText"
            .Left = 366
            .Top = lnDiv.Top + 3
            .Width = 18
            .Height = txtNew.Top - lnDiv.Top - 3
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
        Me.AddChild(CType(scrText, UIControl))

        scrText.MaxValue = ChatTab.mlLineUB
        If muOriginalTexts Is Nothing = True Then
            ReDim muOriginalTexts(49)
            For X = 0 To muOriginalTexts.GetUpperBound(0)
                With muOriginalTexts(X)
                    .sText = ""
                    .lPlayerID = -1
                    .yType = ChatMessageType.eLocalMessage
                    .dtDate = Nothing
                End With
            Next X
        End If
        LoadTabs()

        If muOriginalTexts Is Nothing = False Then
            For X = 0 To mlTabUB
                If moTabs(X) Is Nothing = False Then moTabs(X).ClearLines()
            Next X
            For X = muOriginalTexts.GetUpperBound(0) To 0 Step -1
                With muOriginalTexts(X)
                    If .sText Is Nothing OrElse .sText = "" Then Continue For
                    AddChatMessage(.lPlayerID, .sText, .yType, .dtDate, False)
                End With
            Next X
        End If



        mbIgnoreResize = False

        frmChat_OnResize()

        If muSettings.ChatWindowState = 0 Then
            btnToggle_Click(btnToggle.ControlName)
        End If

    End Sub

    Private Sub txtNew_OnGotFocus() Handles txtNew.OnGotFocus
        If mlCurrentTabIndex > -1 AndAlso moTabs(mlCurrentTabIndex).sMessagePrefix.ToUpper = "/R" Then
            If moTabs(mlCurrentTabIndex).PreviousPMSenders Is Nothing OrElse moTabs(mlCurrentTabIndex).PreviousPMSenders(0) = "" Then Return
            txtNew.Caption = "/pm " & moTabs(mlCurrentTabIndex).PreviousPMSenders(0) & " "
            txtNew.SelStart = txtNew.Caption.Length
            txtNew.CursorPos = txtNew.SelStart
        ElseIf mbSetFromReplyTab = True Then
            mbSetFromReplyTab = False
            txtNew.SelStart = txtNew.Caption.Length
            txtNew.CursorPos = txtNew.SelStart
            txtNew.SelLength = 0
        End If
    End Sub

    'Public Sub SetDefaultChatTabPrefix(ByVal sPrefix As String)
    '    If moTabs Is Nothing = False AndAlso moTabs.GetUpperBound(0) > -1 AndAlso moTabs(0) Is Nothing = False Then
    '        moTabs(0).sMessagePrefix = sPrefix
    '        moTabs(0).SaveTab()
    '    End If
    'End Sub
    Private mbSetFromReplyTab As Boolean = False
    Private Sub txtNew_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNew.OnKeyDown
        If e.KeyCode = Keys.Up Then
            If mlCurrentTabIndex > -1 Then
                txtNew.Caption = moTabs(mlCurrentTabIndex).GetPreviousChatCache(1)
                txtNew.SelStart = txtNew.Caption.Length
            End If
        ElseIf e.KeyCode = Keys.Down Then
            If mlCurrentTabIndex > -1 Then
                txtNew.Caption = moTabs(mlCurrentTabIndex).GetPreviousChatCache(-1)
                txtNew.SelStart = txtNew.Caption.Length
            End If
        ElseIf e.KeyCode = Keys.Tab Then
            If mlCurrentTabIndex > -1 Then
                Dim lCnt As Int32 = 6
                While lCnt > 0
                    moTabs(mlCurrentTabIndex).PreviousPMSenderIdx += 1
                    If moTabs(mlCurrentTabIndex).PreviousPMSenderIdx > moTabs(mlCurrentTabIndex).PreviousPMSenders.GetUpperBound(0) Then moTabs(mlCurrentTabIndex).PreviousPMSenderIdx = 0
                    If moTabs(mlCurrentTabIndex).PreviousPMSenders(moTabs(mlCurrentTabIndex).PreviousPMSenderIdx) <> "" Then Exit While
                    lCnt -= 1
                End While
                If moTabs(mlCurrentTabIndex).PreviousPMSenders(moTabs(mlCurrentTabIndex).PreviousPMSenderIdx) = "" Then Return
                txtNew.Caption = "/pm " & moTabs(mlCurrentTabIndex).PreviousPMSenders(moTabs(mlCurrentTabIndex).PreviousPMSenderIdx) & " "
                txtNew.SelStart = txtNew.Caption.Length
                txtNew.CursorPos = txtNew.SelStart
                mbSetFromReplyTab = True
            End If
        ElseIf e.KeyCode = Keys.Enter Then
            Dim sNewVal As String = Trim$(txtNew.Caption)
            txtNew.Caption = ""

            If sNewVal = "" Then Return

            moTabs(mlCurrentTabIndex).AddPrevChatCache(sNewVal)

            If sNewVal.ToLower.StartsWith("/getid") = True Then
                If goCurrentEnvir Is Nothing = False Then
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) > -1 Then
                            If goCurrentEnvir.oEntity(X).bSelected = True Then
                                AddChatMessage(-1, "ID: " & goCurrentEnvir.oEntity(X).ObjectID, ChatMessageType.eSysAdminMessage, Date.MinValue, False)
                                Return
                            End If
                        End If
                    Next X
                End If
            End If

            If sNewVal.ToLower.StartsWith("/dotechdebugging") = True Then
                'If glPlayerID = 1 OrElse glPlayerID = 2 OrElse glPlayerID = 6 Then
                Dim ofrm As New frmCoeff(goUILib)
                ofrm.Visible = True
                TechBuilderComputer.bShowDebug = Not TechBuilderComputer.bShowDebug
                AddChatMessage(-1, "Tech Debug Toggled", ChatMessageType.eSysAdminMessage, Date.MinValue, False)
                Return
                'End If
            End If

            If sNewVal.ToLower.StartsWith("/observe") OrElse sNewVal.ToLower.StartsWith("/fps") Then
                If sNewVal.ToLower.StartsWith("/fps") = True Then sNewVal = "/observe fps"

                sNewVal = sNewVal.Substring(8).Trim
                If sNewVal.ToLower.StartsWith("fps") = True Then sNewVal = "-254"

                Dim lID As Int32 = CInt(Val(sNewVal))

                If lID = -255 OrElse lID = -254 OrElse lID = -253 Then
                    Dim ofrmObs As frmObserve = New frmObserve(goUILib, lID, 0)
                    ofrmObs.Visible = True
                    ofrmObs = Nothing
                    Return
                End If

                Dim iTypeID As Int16 = 0
                If lID < 1 Then
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then 'AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                            lID = goCurrentEnvir.oEntity(X).ObjectID
                            iTypeID = goCurrentEnvir.oEntity(X).ObjTypeID
                            If lID > 0 Then
                                Dim ofrmObs As frmObserve = New frmObserve(goUILib, lID, iTypeID)
                                ofrmObs.Visible = True
                                ofrmObs = Nothing
                            End If

                            'Exit For
                        End If
                    Next X
                End If



                Return
            ElseIf sNewVal.ToLower.StartsWith("/ignoreadd") = True Then
                Dim sIgnore As String = sNewVal.Substring(10).Trim
                If AddToIgnoreList(sIgnore) = True Then
                    MyBase.moUILib.AddNotification(sIgnore & " has been added to your ignore list.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    ForceSaveTabs()
                End If
                Return
            ElseIf sNewVal.ToLower.StartsWith("/ignoreremove") = True Then
                Dim sIgnore As String = sNewVal.Substring(13).Trim
                If RemoveFromIgnoreList(sIgnore) = True Then
                    MyBase.moUILib.AddNotification(sIgnore & " has been removed from your ignore list.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    ForceSaveTabs()
                Else
                    MyBase.moUILib.AddNotification(sIgnore & " is not being ignored", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
                Return
            ElseIf sNewVal.ToLower.StartsWith("/ignore") = True Then
                AddChatMessage(-1, "You are currently ignoring:", ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                For X As Int32 = 0 To mlIgnoreListUB
                    If msIgnoreList(X) Is Nothing = False AndAlso msIgnoreList(X) <> "" Then
                        AddChatMessage(-1, "  " & msIgnoreList(X), ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                    End If
                Next X
                Return
            ElseIf sNewVal.ToLower.StartsWith("/forcenext") = True Then
                If NewTutorialManager.TutorialOn = True Then
                    Dim lStepID As Int32 = NewTutorialManager.GetTutorialStepID()
                    NewTutorialManager.FindAndExecuteStepID(lStepID + 1)
                    Return
                End If
            End If

            If sNewVal.ToLower.StartsWith("/clientver") = True Then
                AddChatMessage(-1, "Client Version: " & gl_CLIENT_VERSION.ToString, ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                Return
            End If

            If sNewVal.ToLower.StartsWith("/help") = True OrElse (sNewVal.ToLower.StartsWith("/") AndAlso sNewVal.Length = 1) Then
                'ok, display help
                Dim sText() As String = Split(GetHelpCmdText(), vbCrLf)
                AddChatMessage(-1, "Chat Help Command List:", ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                For X As Int32 = 0 To sText.GetUpperBound(0)
                    AddChatMessage(-1, sText(X), ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                Next X
            ElseIf sNewVal.ToLower.StartsWith("/keymap") = True Then
                Dim sText() As String = Split(GetKeyMapText, vbCrLf)
                AddChatMessage(-1, "Keyboard Command List:", ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                For X As Int32 = 0 To sText.GetUpperBound(0)
                    AddChatMessage(-1, sText(X), ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                Next X
            ElseIf sNewVal.ToLower.StartsWith("/findentity") = True Then ' AndAlso (glPlayerID = 1 OrElse glPlayerID = 2 OrElse glPlayerID = 6 OrElse glPlayerID = 7 OrElse glPlayerID = 9) Then
                Dim sTmp As String = Replace$(sNewVal.ToLower, "/findentity", "")
                Dim sVals() As String = Split(sTmp.Trim, ",")
                If sVals.GetUpperBound(0) = 1 Then
                    Dim lID As Int32 = CInt(Val(sVals(0).Trim))
                    Dim iTypeID As Int16 = CShort(Val(sVals(1).Trim))

                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) = lID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iTypeID Then
                            goCamera.mlCameraAtX = CInt(goCurrentEnvir.oEntity(X).LocX)
                            goCamera.mlCameraAtZ = CInt(goCurrentEnvir.oEntity(X).LocZ)
                            goCamera.mlCameraAtY = 0
                            goCamera.mlCameraX = goCamera.mlCameraAtX
                            goCamera.mlCameraZ = goCamera.mlCameraAtZ + 1000
                            goCamera.mlCameraAtY = CInt(goCurrentEnvir.oEntity(X).LocY) + 1000
                            AddChatMessage(-1, "Entity Found", ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                            Return
                        End If
                    Next X

                    AddChatMessage(-1, "Entity not found", ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                End If
            Else
                Dim yMsg() As Byte

                If sNewVal.ToLower.StartsWith("/warpoints") = True Then
                    Dim bFoundItem As Boolean = False
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                            sNewVal = "/warpoints " & goCurrentEnvir.oEntity(X).ObjectID.ToString & "," & goCurrentEnvir.oEntity(X).ObjTypeID.ToString
                            bFoundItem = True
                            Exit For
                        End If
                    Next X
                    If bFoundItem = False Then
                        AddChatMessage(-1, "You must select a unit or facility first.", ChatMessageType.eSysAdminMessage, Date.MinValue, False)
                        Return
                    End If
                End If

                '????
                sNewVal = sNewVal.Trim

                If sNewVal.StartsWith("/") = False Then
                    'Append the message only if it does not have a slash. A slash indicates the player wants to do something else
                    If moTabs(mlCurrentTabIndex).sMessagePrefix <> "" Then
                        If moTabs(mlCurrentTabIndex).sMessagePrefix.EndsWith(" ") = False Then
                            sNewVal = moTabs(mlCurrentTabIndex).sMessagePrefix & " " & sNewVal
                        Else : sNewVal = moTabs(mlCurrentTabIndex).sMessagePrefix & sNewVal
                        End If
                    ElseIf goCurrentPlayer.yPlayerPhase = 0 AndAlso mlCurrentTabIndex = 0 Then
                        moTabs(mlCurrentTabIndex).sMessagePrefix = "/GENERAL "
                        sNewVal = "/GENERAL " & sNewVal
                        moTabs(mlCurrentTabIndex).SaveTab()
                    End If

                    '	'Now, if we still do not start with a slash..
                    '	If sNewVal.StartsWith("/") = False Then
                    '		'append the player's name for a Local message
                    '		sNewVal = goCurrentPlayer.PlayerName & ": " & sNewVal
                    '	End If
                    'ElseIf sNewVal.ToLower.StartsWith("/local") = True Then
                    '	sNewVal = goCurrentPlayer.PlayerName & ": " & sNewVal.Substring(6).Trim
                ElseIf sNewVal.ToLower.StartsWith("/r ") = True Then
                    If moTabs(mlCurrentTabIndex).PreviousPMSenders Is Nothing OrElse moTabs(mlCurrentTabIndex).PreviousPMSenders(0) = "" Then
                        AddChatMessage(-1, "No one has sent you any messages yet.", ChatMessageType.ePrivateMessage, Date.MinValue, True)
                        Return
                    Else
                        sNewVal = "/pm " & moTabs(mlCurrentTabIndex).PreviousPMSenders(0) & " " & sNewVal.Substring(2)
                    End If
                End If

                If sNewVal.ToUpper.StartsWith("/GENERAL") = True Then 'OrElse sNewVal.ToUpper.StartsWith("/FAQ_CSR") = True Then
                    If NewTutorialManager.TutorialOn = True Then
                        If NewTutorialManager.GetTutorialStepID = 319 Then
                            If sNewVal.ToUpper.Contains("PLANET") = True AndAlso sNewVal.ToUpper.Contains("FALL") = True Then
                                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eFirstChatMsgSent, -1, -1, -1, "")
                            End If
                        End If
                    End If
                End If

                'TODO: if we are to do any pre-send processing, now is the time
                If sNewVal.ToLower.StartsWith("/gm") = True AndAlso sNewVal.Length > 5 Then
                    AddChatMessage(-1, goCurrentPlayer.PlayerName & ": " & sNewVal.Substring(4), ChatMessageType.eSysAdminMessage, Date.MinValue, True)
                End If

                BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eChatMessageSent, sNewVal.Length, 0)

                'Prepare our message
                ReDim yMsg(sNewVal.Length + 5)
                System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(sNewVal.Length).CopyTo(yMsg, 2)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(yMsg, 6)

                ''determine who to send it to
                'If sNewVal.StartsWith("/") = True Then
                '	'ok, send to the primary server
                MyBase.moUILib.SendMsgToPrimary(yMsg)
                'Else
                '	'ok, send to the region server
                '	MyBase.moUILib.SendMsgToRegion(yMsg)
                'End If
                End If

                MyBase.moUILib.FocusedControl = Nothing
                txtNew.HasFocus = False

                mbForceNextRefresh = True

        End If

    End Sub

    Private Sub StoreOriginalText(ByVal sLine As String, ByVal lPlayerID As Int32, ByVal yType As ChatMessageType)
        For X As Int32 = muOriginalTexts.GetUpperBound(0) To 1 Step -1
            muOriginalTexts(X) = muOriginalTexts(X - 1)
        Next X
        With muOriginalTexts(0)
            .yType = yType
            .lPlayerID = lPlayerID
            .sText = sLine
            .dtDate = Now
        End With
    End Sub

    Public Sub AddChatMessage(ByVal lPlayerID As Int32, ByVal sText As String, ByVal yType As ChatMessageType, ByVal dtDate As Date, ByVal bStoreMsg As Boolean)
        mlLastMsgCycle = glCurrentCycle

        'Determine our matching filter
        If bStoreMsg = True Then StoreOriginalText(sText, lPlayerID, yType)

        Dim lFilter As Int32 = CInt(Math.Pow(2, yType))
        Dim sChannel As String = ""
        If lFilter = ChatTab.ChatFilter.eChannelMessages Then
            'Find our channel
            Dim lTemp As Int32 = sText.IndexOf("(")
            If lTemp <> -1 Then
                sChannel = sText.Substring(lTemp + 1, sText.IndexOf(")", lTemp) - lTemp - 1)
            End If
        End If
        If sChannel = "Emperor" Then yType = ChatMessageType.eSenateMessage
        'If yType = ChatMessageType.ePrivateMessage Then
        If IgnoreMessage(sText) = True Then Return
        'End If

        If muSettings.ChatTimeStamps = True AndAlso dtDate <> Date.MinValue AndAlso yType <> ChatMessageType.eNotificationMessage AndAlso yType <> ChatMessageType.eSysAdminMessage Then
            sText = dtDate.ToString("HH:mm") & " " & sText
        End If

        'first, go thru and shift our lines
        If lblText.GetUpperBound(0) = -1 Then Return
        Dim rcTest As Rectangle = lblText(0).GetTextDimensions(sText)
        While rcTest.Width > lblText(0).Width
            Dim lStrLen As Int32 = sText.Length
            Dim lNewLen As Int32 = CInt(Math.Floor(((lblText(0).Width - 5) / rcTest.Width) * lStrLen)) - 1
            Dim lSplit As Int32 = InStrRev(sText, " ", lNewLen)

            If lSplit < 1 Then
                lSplit = CInt(lStrLen * (lblText(0).Width / (rcTest.Width + 1)))
                If lSplit < 1 Then lSplit = lStrLen
                If lSplit > lStrLen Then Exit While
            End If

            Dim sLeftVal As String = Mid$(sText, 1, lSplit - 1)
            'AddChatMessage(lPlayerID, sLeftVal, yType)
            AddParsedMessage(lPlayerID, sLeftVal, lFilter, sChannel, yType)
            sText = Mid$(sText, lSplit + 1)

            rcTest = lblText(0).GetTextDimensions(sText)
        End While

        'moTabs(mlCurrentTabIndex).AddLine(sText, lPlayerID)
        For X As Int32 = 0 To mlTabUB
            If moTabs(X).AcceptsLine(lFilter, sChannel) = True Then
                moTabs(X).AddLine(sText, lPlayerID, yType)
            End If
        Next X

        mbHasFlash = False
        For X As Int32 = 0 To mlTabUB
            If X = mlCurrentTabIndex AndAlso mbMinimized = False Then
                moTabs(X).bHasNewMessage = False
            ElseIf moTabs(X).bHasNewMessage = True AndAlso mbFlashColor = True Then
                txtTabs(X).BackColorEnabled = clrFlashOn 'System.Drawing.Color.FromArgb(255, 175, 192, 80)
            Else
                txtTabs(X).BackColorEnabled = clrFlashOff 'muSettings.InterfaceTextBoxFillColor
            End If
            mbHasFlash = mbHasFlash OrElse moTabs(X).bHasNewMessage
        Next X

        RefreshView()
    End Sub
    
	Public Sub SetNewMsgText(ByVal sText As String)
		txtNew.Caption = sText
		txtNew.SelStart = txtNew.Caption.Length
	End Sub

    Private Sub AddParsedMessage(ByVal lPlayerID As Int32, ByVal sText As String, ByVal lFilter As Int32, ByVal sChannel As String, ByVal yType As ChatMessageType)
        'first, go thru and shift our lines
        Dim rcTest As Rectangle = lblText(0).GetTextDimensions(sText)
        While rcTest.Width > lblText(0).Width
            Dim lStrLen As Int32 = sText.Length
            Dim lNewLen As Int32 = CInt(Math.Floor(((lblText(0).Width - 5) / rcTest.Width) * lStrLen)) - 1
            Dim lSplit As Int32 = InStrRev(sText, " ", lNewLen)

            If lSplit < 1 Then
                lSplit = CInt(lStrLen * (rcTest.Width / (lblText(0).Width + 1)))
                If lSplit < 1 Then lSplit = lStrLen
                If lSplit > lStrLen Then Exit While
            End If

            Dim sLeftVal As String = Mid$(sText, 1, lSplit - 1)
            AddParsedMessage(lPlayerID, sLeftVal, lFilter, sChannel, yType)
            sText = Mid$(sText, lSplit + 1)

            rcTest = lblText(0).GetTextDimensions(sText)
        End While

        For X As Int32 = 0 To mlTabUB
            If moTabs(X).AcceptsLine(lFilter, sChannel) = True Then
                moTabs(X).AddLine(sText, lPlayerID, yType)
            End If
        Next X

    End Sub

    Private Sub RefreshView()
        Dim X As Int32
        Dim lIdx As Int32

        'Now, fill our labels
        For X = 0 To lblText.Length - 1
            lIdx = scrText.Value + X

            With moTabs(mlCurrentTabIndex)
                If lIdx > ChatTab.mlLineUB Then
                    lblText(X).Caption = ""
                Else
					lblText(X).Caption = .msLines(lIdx)

					If .mlSenderID(lIdx) = glPlayerID OrElse .myMessageType(lIdx) > ChatTab.ChatFilterCount - 1 Then
						lblText(X).ForeColor = muSettings.DefaultChatColor	'System.Drawing.Color.FromArgb(255, 255, 255, 255)
					Else
						Select Case .myMessageType(lIdx)
							Case ChatMessageType.eAllianceMessage
								lblText(X).ForeColor = muSettings.GuildChatColor
							Case ChatMessageType.eChannelMessage
								'TODO: Determine the channel ID
								lblText(X).ForeColor = muSettings.ChannelChatColor
							Case ChatMessageType.eLocalMessage
								lblText(X).ForeColor = muSettings.LocalChatColor
							Case ChatMessageType.eNotificationMessage
								lblText(X).ForeColor = muSettings.AlertChatColor
							Case ChatMessageType.ePrivateMessage
								lblText(X).ForeColor = muSettings.PMChatColor
							Case ChatMessageType.eSenateMessage
								lblText(X).ForeColor = muSettings.SenateChatColor
							Case ChatMessageType.eSysAdminMessage
								lblText(X).ForeColor = muSettings.StatusChatColor
							Case Else
								lblText(X).ForeColor = muSettings.DefaultChatColor
						End Select
					End If
                End If
            End With
        Next X
        Me.IsDirty = True
    End Sub

    Private Sub scrText_ValueChange() Handles scrText.ValueChange
        RefreshView()
        Me.IsDirty = True
    End Sub

    Private Sub LoadTabs()
        Try
            Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sFile.EndsWith("\") = False Then sFile &= "\"
            sFile &= goCurrentPlayer.PlayerName & ".cht"
            Dim oINI As InitFile = New InitFile(sFile)

            Dim lSeqNbr As Int32 = 0

            Dim sHdr As String = "ChatTab" & lSeqNbr

            Dim sName As String = oINI.GetString(sHdr, "TabName", "Chat")
            While sName <> ""
                mlTabUB += 1
                ReDim Preserve moTabs(mlTabUB)
                moTabs(mlTabUB) = New ChatTab()
                With moTabs(mlTabUB)
                    .sTabName = sName
                    .sChannel = oINI.GetString(sHdr, "Channel", "")
                    .lSequenceNumber = lSeqNbr

                    If mlTabUB = 0 Then
                        Dim lTemp As ChatTab.ChatFilter
                        lTemp = ChatTab.ChatFilter.eAllianceMessages Or ChatTab.ChatFilter.eChannelMessages Or ChatTab.ChatFilter.eLocalMessages Or ChatTab.ChatFilter.ePMs Or ChatTab.ChatFilter.eSenateMessages Or ChatTab.ChatFilter.eSysAdminMessages Or ChatTab.ChatFilter.eAliasChatMessage
                        .lFilter = CInt(Val(oINI.GetString(sHdr, "Filter", CInt(lTemp).ToString)))
                    Else : .lFilter = CInt(Val(oINI.GetString(sHdr, "Filter", "0")))
                    End If

                    Dim sDefMsgPrfx As String = ""
                    If mlTabUB = 0 Then
                        sDefMsgPrfx = "/GENERAL "
                    End If
                    .sMessagePrefix = oINI.GetString(sHdr, "MsgPrefix", sDefMsgPrfx)

                    If .sMessagePrefix Is Nothing = False AndAlso .sMessagePrefix.ToUpper.StartsWith("/GENERAL") = True Then
                        If goCurrentPlayer Is Nothing = False Then
                            If goCurrentPlayer.yPlayerPhase = 255 Then
                                .sMessagePrefix = "/GENERAL "
                                oINI.WriteString(sHdr, "MsgPrefix", .sMessagePrefix)
                            End If
                        End If
                    End If
                    'For X As Int32 = 0 To .clrMessages.GetUpperBound(0)
                    '    Dim lR As Int32 = CInt(Val(oINI.GetString(sHdr, "Color" & X & "_R", .clrMessages(X).R.ToString)))
                    '    Dim lG As Int32 = CInt(Val(oINI.GetString(sHdr, "Color" & X & "_G", .clrMessages(X).G.ToString)))
                    '    Dim lB As Int32 = CInt(Val(oINI.GetString(sHdr, "Color" & X & "_B", .clrMessages(X).B.ToString)))
                    '    .clrMessages(X) = System.Drawing.Color.FromArgb(255, lR, lG, lB)
                    'Next X
                End With

                lSeqNbr += 1
                sHdr = "ChatTab" & lSeqNbr
                sName = oINI.GetString(sHdr, "TabName", "")
            End While

            sHdr = "IgnoreList"
            Dim lIgnoreIdx As Int32 = 0
            sName = oINI.GetString(sHdr, "Ignore_" & lIgnoreIdx, "")
            While sName <> ""
                mlIgnoreListUB += 1
                ReDim Preserve msIgnoreList(mlIgnoreListUB)
                msIgnoreList(mlIgnoreListUB) = sName

                lIgnoreIdx += 1
                sName = oINI.GetString(sHdr, "Ignore_" & lIgnoreIdx, "")
            End While

            oINI = Nothing
        Catch
        End Try

        RefreshTabs()
        SetTabIndex(0)
        If moTabs(0).sMessagePrefix.ToUpper.StartsWith("/GENERAL") = False Then
            Me.AddChatMessage(-1, "Warning.  Your primary/first Chat Tab is not set to /GENERAL.  If this is unintentional you can double click the tab, and enter /GENERAL in the message prefix section.", ChatMessageType.eSysAdminMessage, Nothing, True)
        End If
    End Sub

    Private Sub RefreshTabs()
        If txtTabs Is Nothing Then ReDim txtTabs(mlTabUB)
        If txtTabs.GetUpperBound(0) <> mlTabUB Then ReDim Preserve txtTabs(mlTabUB)

        With muSettings.InterfaceTextBoxFillColor
            clrFlashOff = System.Drawing.Color.FromArgb(255, .R \ 2, .G \ 2, .B \ 2)
            clrFlashOn = System.Drawing.Color.FromArgb(255, 255 - .R, 255 - .G, 255 - .B)
        End With
        clrSelected = muSettings.InterfaceTextBoxFillColor


        For X As Int32 = 0 To mlTabUB
            Dim bAdded As Boolean = False

            If txtTabs(X) Is Nothing Then
                txtTabs(X) = New UITextBox(goUILib)
                bAdded = True
            End If

            With txtTabs(X)
                .ControlName = "txtTabs(" & X & ")"
                .Left = 1 + (X * 75)
                .Top = 1
                .Width = 75
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = moTabs(X).sTabName
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
                If X = mlCurrentTabIndex Then
                    '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 92, 128, 160)
                    .BackColorEnabled = clrSelected
                Else : .BackColorEnabled = clrFlashOff 'muSettings.InterfaceTextBoxFillColor
                End If
                .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
                .MaxLength = 0
                .Locked = True
                .BorderColor = muSettings.InterfaceBorderColor
                .BorderLineWidth = 1
                .bRoundedBorder = True
            End With

            If bAdded = True Then
                Me.AddChild(CType(txtTabs(X), UITextBox))
                AddHandler txtTabs(X).OnMouseDown, AddressOf TabClick
            End If
        Next X
    End Sub

    Public Sub RefreshChatText()
        If muOriginalTexts Is Nothing = False Then
            For X As Int32 = 0 To mlTabUB
                If moTabs(X) Is Nothing = False Then moTabs(X).ClearLines()
            Next X
            For X As Int32 = muOriginalTexts.GetUpperBound(0) To 0 Step -1
                With muOriginalTexts(X)
                    If .sText Is Nothing OrElse .sText = "" Then Continue For
                    AddChatMessage(.lPlayerID, .sText, .yType, .dtDate, False)
                End With
            Next X
        End If
    End Sub

    Private Sub SetTabIndex(ByVal lIdx As Int32)
        txtTabs(mlCurrentTabIndex).BackColorEnabled = clrFlashOff 'muSettings.InterfaceTextBoxFillColor
        txtTabs(lIdx).BackColorEnabled = clrSelected
        'txtTabs(lIdx).BackColorEnabled = System.Drawing.Color.FromArgb(255, 92, 128, 160)
        mlCurrentTabIndex = lIdx
        RefreshView()
    End Sub

    Private Sub TabClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)

        If mbMinimized = True Then btnToggle_Click("btnToggle")

        Dim lTabIndex As Int32 = -1

        For X As Int32 = 0 To mlTabUB
            Dim oPos As Point = txtTabs(X).GetAbsolutePosition()
            If lMouseX > oPos.X AndAlso lMouseX < oPos.X + txtTabs(X).Width Then
                lTabIndex = X
                Exit For
            End If
        Next X

        If lTabIndex <> -1 Then
            If glCurrentCycle - mlPreviousTabClick < 10 AndAlso lTabIndex = mlPreviousTabClickIdx Then
                Dim ofrmProps As New frmChatTabProps(goUILib)
                ofrmProps.SetFromTab(moTabs(lTabIndex))
                ofrmProps = Nothing
            Else
                SetTabIndex(lTabIndex)
                mlPreviousTabClick = glCurrentCycle
            End If
            mlPreviousTabClickIdx = lTabIndex
        End If

        If goUILib.FocusedControl Is Nothing = False Then
            goUILib.FocusedControl.HasFocus = False
            goUILib.FocusedControl = Nothing
        End If

    End Sub

	Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click

		If mlTabUB > 3 Then
			MyBase.moUILib.AddNotification("There is currently a hard limit of 5 chat tabs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		mlTabUB += 1
		ReDim Preserve moTabs(mlTabUB)
		moTabs(mlTabUB) = New ChatTab()
		With moTabs(mlTabUB)
			.lFilter = 0
			.lSequenceNumber = mlTabUB
			.sChannel = ""
			.sMessagePrefix = ""
			.sTabName = "New Tab"
		End With
		RefreshTabs()

		Dim ofrmTabProps As frmChatTabProps = New frmChatTabProps(goUILib)
		ofrmTabProps.SetFromTab(moTabs(mlTabUB))
		ofrmTabProps = Nothing
		ForceSaveTabs()
	End Sub

	Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
		If mlCurrentTabIndex <> 0 Then
			For X As Int32 = mlCurrentTabIndex To mlTabUB - 1
				moTabs(X) = moTabs(X + 1)
			Next X
			mlTabUB -= 1
            ReDim Preserve moTabs(mlTabUB)
 

			For X As Int32 = 0 To Me.ChildrenUB
				If Me.moChildren(X).ControlName = txtTabs(mlTabUB + 1).ControlName Then
					Me.RemoveChild(X)
					Exit For
				End If
			Next X
			ReDim Preserve txtTabs(mlTabUB)
			mlCurrentTabIndex = 0
            SetTabIndex(0)
            Me.IsDirty = True

            ForceSaveTabs()

            RefreshTabs()
		End If
	End Sub

	Public Sub TabPropsClosed()
		RefreshTabs()
	End Sub

	Private Sub ForceSaveTabs()
		Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
		If sFile.EndsWith("\") = False Then sFile &= "\"
		sFile &= goCurrentPlayer.PlayerName & ".cht"
		If Exists(sFile) = True Then Kill(sFile)

		For X As Int32 = 0 To mlTabUB
			moTabs(X).SaveTab()
		Next X

		Dim oINI As InitFile = New InitFile(sFile)
		Dim lIdx As Int32 = 0
		For X As Int32 = 0 To mlIgnoreListUB
			If msIgnoreList(X) <> "" Then
				oINI.WriteString("IgnoreList", "Ignore_" & lIdx, msIgnoreList(X))
				lIdx += 1
			End If
		Next X
	End Sub

	Private Sub frmChat_OnNewFrame() Handles Me.OnNewFrame
		If mbHasFlash = True AndAlso glCurrentCycle - mlLastTabFlash > 30 Then
			mbFlashColor = Not mbFlashColor
			mlLastTabFlash = glCurrentCycle

			'Set our tabs
			mbHasFlash = False
			For X As Int32 = 0 To mlTabUB
				If X = mlCurrentTabIndex Then
					moTabs(X).bHasNewMessage = False
				ElseIf moTabs(X).bHasNewMessage = True AndAlso mbFlashColor = True Then
                    txtTabs(X).BackColorEnabled = clrFlashOn 'System.Drawing.Color.FromArgb(255, 175, 192, 80)
				Else
                    txtTabs(X).BackColorEnabled = clrFlashOff ' muSettings.InterfaceTextBoxFillColor
				End If
				mbHasFlash = mbHasFlash OrElse moTabs(X).bHasNewMessage
			Next X
			Me.IsDirty = True
		End If

        If mbForceNextRefresh = True AndAlso glCurrentCycle - mlForcefulRefresh > 45 Then
            mlForcefulRefresh = glCurrentCycle
            Me.IsDirty = True
            mbForceNextRefresh = False
        End If

        Me.IsDirty = Me.IsDirty OrElse (mlLastMsgCycle >= mlLastRenderCycle)
	End Sub

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click

        If NewTutorialManager.TutorialOn = True Then
            MyBase.moUILib.AddNotification("The Chat tutorial will be available at another time.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eChat)
    End Sub

    Private Sub frmChat_OnRender() Handles Me.OnRender
        mlLastRenderCycle = glCurrentCycle
    End Sub

    Private Sub frmChat_OnResize() Handles Me.OnResize
        If mbIgnoreResize = True Then Return

        If btnDelete Is Nothing = False Then
            btnDelete.Left = Me.Width - btnDelete.Width - 2
            If btnAdd Is Nothing = False Then
                btnAdd.Left = btnDelete.Left - btnAdd.Width - 1
                If btnHelp Is Nothing = False Then
                    btnHelp.Left = btnAdd.Left - btnHelp.Width - 1
                    If btnToggle Is Nothing = False Then
                        btnToggle.Left = btnHelp.Left - btnToggle.Width + 1 '+1 not -1 as the image has a shadow on the right edge.
                    End If
                End If
            End If
        End If
        If scrText Is Nothing = False Then
            scrText.Left = Me.Width - scrText.Width - 2
            scrText.Height = Me.Height - scrText.Top - txtNew.Height - 2
        End If
        If lblText Is Nothing = False Then
            For X As Int32 = 0 To lblText.GetUpperBound(0)
                If lblText(X) Is Nothing = False Then
                    lblText(X).Width = Me.Width - scrText.Width
                End If
            Next X
        End If
        If txtNew Is Nothing = False Then
            txtNew.Width = Me.Width - (txtNew.Left * 2) - 1
            txtNew.Top = Me.Height - txtNew.Height - 1
        End If
        If lnDiv Is Nothing = False Then lnDiv.Width = Me.Width - 1

        'Ok, determine what our thing is...
        If lblText Is Nothing = False Then
            Dim lPrevUB As Int32 = lblText.GetUpperBound(0)
            Dim lNewUB As Int32 = ((txtNew.Top - lnDiv.Top) \ 18) - 1

            If lNewUB < lPrevUB Then
                For X As Int32 = lPrevUB To lNewUB + 1 Step -1
                    For Y As Int32 = 0 To Me.ChildrenUB
                        If Object.Equals(Me.moChildren(Y), lblText(X)) = True Then
                            RemoveChild(Y)
                            Exit For
                        End If
                    Next Y
                Next X
                ReDim Preserve lblText(lNewUB)
            ElseIf lNewUB > lPrevUB Then
                ReDim Preserve lblText(lNewUB)
                'ok, now, go thru and add them
                For X As Int32 = lPrevUB + 1 To lNewUB
                    lblText(X) = New UILabel(MyBase.moUILib)
                    With lblText(X)
                        .ControlName = "lblText(" & X & ")"
                        .Left = 3
                        .Top = txtNew.Top - ((X + 1) * 18)
                        .Width = Me.Width - 19
                        .Height = 17
                        .Enabled = True
                        .Visible = True
                        .Caption = ""
                        '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
                        .ForeColor = muSettings.InterfaceBorderColor
                        .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
                        .DrawBackImage = False
                        .FontFormat = DrawTextFormat.Bottom Or DrawTextFormat.Left
                    End With
                    Me.AddChild(CType(lblText(X), UIControl))
                Next X
            End If

            For X As Int32 = 0 To lblText.GetUpperBound(0)
                lblText(X).Width = Me.Width - 19
                lblText(X).Top = txtNew.Top - ((X + 1) * 18)
            Next X
        End If

        RefreshChatText()

        Dim ofrmCS As frmConnectionStatus = CType(goUILib.GetWindow("frmConnectionStatus"), frmConnectionStatus)
        If ofrmCS Is Nothing = False Then
            ofrmCS.Left = Me.Left + Me.Width - ofrmCS.Width
            ofrmCS.Top = Me.Top - ofrmCS.Height - 2
            ofrmCS.Visible = True
        End If

        muSettings.ChatWindowHeight = Me.Height
        muSettings.ChatWindowWidth = Me.Width

    End Sub

    Private Sub frmChat_WindowMoved() Handles Me.WindowMoved
        muSettings.ChatWindowLocX = Me.Left
        muSettings.ChatWindowLocY = Me.Top
        Dim ofrmCS As frmConnectionStatus = CType(goUILib.GetWindow("frmConnectionStatus"), frmConnectionStatus)
        If ofrmCS Is Nothing = False Then
            ofrmCS.Left = Me.Left + Me.Width - ofrmCS.Width
            ofrmCS.Top = Me.Top - ofrmCS.Height - 2
            ofrmCS.Visible = True
        End If
    End Sub

    Public Shared Function GetKeyMapText() As String
        Dim oSB As New System.Text.StringBuilder()
        oSB.AppendLine("Escape - Opens/Closes options window")
        oSB.AppendLine("Backspace - Goes up one environment level")
        oSB.AppendLine("Home - Returns to the current environment")
        oSB.AppendLine("End - Returns camera to it's default direction")
        oSB.AppendLine("Tab - Expands the notification window to show history")
        oSB.AppendLine("")
        oSB.AppendLine("A - Toggle Auto-Launch for the selected building")
        oSB.AppendLine("B - Opens the build window for the selection")
        oSB.AppendLine("C - Opens the contents window for the selection")
        oSB.AppendLine("G - Opens/Closes the Unit Goto Window")
        oSB.AppendLine("P - Toggle Power for the selected building")
        oSB.AppendLine("U - Sends the selected unit(s) to orbit")
        oSB.AppendLine("G - Toggles the grid in space views")
        oSB.AppendLine("T - Turns on Tracking mode")
        'oSB.AppendLine("W - Toggles wireframe mode")
        oSB.AppendLine("Z - overrides mouse wheel zooming")
        oSB.AppendLine("")
        oSB.AppendLine("Alt + C - captures a screenshot")
        oSB.AppendLine("Alt - Displays distance information from the current selection or Radar rings for the current unit")
        oSB.AppendLine("")
        oSB.AppendLine("F1 - Opens/Closes the Quick Help Table of Contents Window")
        oSB.AppendLine("F2 - Opens/Closes the Email Window")
        oSB.AppendLine("F3 - Opens/Closes the Battlegroup Management Window")
        oSB.AppendLine("F4 - Opens/Closes the Trade Management Window")
        oSB.AppendLine("F5 - Opens/Closes the Diplomacy Window")
        oSB.AppendLine("F6 - Opens/Closes the Colony Stats window")
        oSB.AppendLine("F7 - Opens/Closes the Budget Window")
        oSB.AppendLine("F8 - Opens/Closes the Mining Window")
        oSB.AppendLine("F9 - Opens/Closes the Agent Window")
        oSB.AppendLine("F10 - Opens/Closes the Formations Window")
        oSB.AppendLine("F11 - Opens/Closes the Guild Window")
        oSB.AppendLine("F12 - Compose an email to support")
        oSB.AppendLine("")
        oSB.AppendLine("Ctrl+C - Copy selected text to the clipboard")
        oSB.AppendLine("Ctrl+E - Selects the next engineer")
        oSB.AppendLine("Ctrl+F - Selects the next facility")
        oSB.AppendLine("Ctrl+I - Selects the next idle engineer")
        oSB.AppendLine("Ctrl+M - toggles whether mineral caches are rendered")
        oSB.AppendLine("Ctrl+N - Selects the next unit or facility")
        oSB.AppendLine("Ctrl+P - Selects the next unpowered facility in the environment")
        oSB.AppendLine("Ctrl+Q - Immediately closes the game")
        oSB.AppendLine("Ctrl+R - Toggles Set Rally Point Mode")
        oSB.AppendLine("Ctrl+S - Selects all objects that are similar to the currently selected objects")
        oSB.AppendLine("Ctrl+T - Tether unit to the current location.  (Will return after engaging in combat)")
        oSB.AppendLine("Ctrl+U - Selects the next unit")
        oSB.AppendLine("Ctrl+V - Paste text from the clipboard")
        oSB.AppendLine("Ctrl+(0-9) - Assigns selected objects to the control group of that number")
        oSB.AppendLine("Ctrl+Left - Shifts the Environment Display down by one environment.")
        oSB.AppendLine("Ctrl+Right - Shifts the Environment Display up by one environment.")
        oSB.AppendLine("Ctrl+Enter - Forces the Environment Display to goto the selected environment.")
        oSB.AppendLine("")
        oSB.AppendLine("(0-9) - Recalls control group selections, clearing the current selection")
        oSB.AppendLine("Shift+(0-9) - Recalls control group selections, without clearing the current selection")
        oSB.AppendLine("Ctrl+Shift+(0-9) - Recalls control group selections without clearing current selections and assigns resulting selects to the control group of that number")
        oSB.AppendLine("Ctrl+Shift+T - Clear Tether location")
        oSB.AppendLine("Ctrl+Shift+S - Selects all objects that are similarly named to the currently selected objects")
        oSB.AppendLine("Ctrl+Shift+Left - Shifts the Environment Display down by one environment, and goes there.")
        oSB.AppendLine("Ctrl+Shift+Right - Shifts the Environment Display up by one environment, and goes there.")
        oSB.AppendLine("")
        oSB.AppendLine("~ (tilde) - Toggles UI rendering")
        Return oSB.ToString
    End Function

    Public Shared Function GetHelpCmdText() As String
        Dim oSB As New System.Text.StringBuilder
        oSB.AppendLine("Private message: /pm <playername> <msg>")
        oSB.AppendLine("Reply to Last PM: /r <msg>")
        'oSB.AppendLine("Create Channel: /create <channelname>")
        'oSB.AppendLine("Join Channel: /join <channelname>, <password (optional)>")
        'oSB.AppendLine("Give Admin Rights: /admin <channelname>, <playername>")
        'oSB.AppendLine("Set Password: /setpassword <channelname>, <password>")
        'oSB.AppendLine("Kick Player: /kick <playername>, <channelname>")
        'oSB.AppendLine("Invite player: /invite <playername>, <channelname>")
        'oSB.AppendLine("Leave Channel: /leave <channelname>")
        oSB.AppendLine("Alias Chat (sends to aliased players): /alias <msg>")
        oSB.AppendLine("Chat in Channel: /<channelname> <msg>")
        oSB.AppendLine("Chat in Environment: /local <msg>")
        oSB.AppendLine("Display Keyboard Command List: /keymap")
        'oSB.AppendLine("List of players in a channel: /who <channelname>")
		oSB.AppendLine("Displays Amount of Time Logged On: /played")
        'oSB.AppendLine("Manually add a player relationship: /addrelationship <playername>")
        oSB.AppendLine("Adds the player to the ignore list: /ignoreadd <playername>")
        oSB.AppendLine("Removes a player from the ignore list: /ignoreremove <playername>")
        oSB.AppendLine("Shows a list of players on the ignore list: /ignore")
        'oSB.AppendLine("Shows the warpoints for the currently selected unit: /warpoints")
        Return oSB.ToString
	End Function

	Private Function AddToIgnoreList(ByVal sItem As String) As Boolean
		sItem = sItem.ToUpper.Trim
		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To mlIgnoreListUB
			If msIgnoreList(X) = sItem Then
				MyBase.moUILib.AddNotification("That player is already in the ignore list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return False
			ElseIf msIgnoreList(X) = "" AndAlso lIdx = -1 Then
				lIdx = X
			End If
		Next X
		If lIdx = -1 Then
			mlIgnoreListUB += 1
			ReDim Preserve msIgnoreList(mlIgnoreListUB)
			lIdx = mlIgnoreListUB
		End If
		msIgnoreList(lIdx) = sItem
		Return True
	End Function
	Private Function RemoveFromIgnoreList(ByVal sItem As String) As Boolean
		sItem = sItem.Trim.ToUpper
		Dim bFound As Boolean = False
		For X As Int32 = 0 To mlIgnoreListUB
			If msIgnoreList(X) = sItem Then
				bFound = True
				msIgnoreList(X) = ""
			End If
		Next X
		Return bFound
	End Function
	Private Function IgnoreMessage(ByVal sText As String) As Boolean
		If sText Is Nothing Then Return False
        Dim lIdx As Int32 = sText.IndexOf("tells you,")
        Try
            If lIdx <> -1 Then
                Dim sWho As String = sText.Substring(0, lIdx).Trim.ToUpper

                For X As Int32 = 0 To mlIgnoreListUB
                    If msIgnoreList(X) = sWho Then
                        Return True
                    End If
                Next X
            Else
                sText = sText.ToUpper
                For X As Int32 = 0 To mlIgnoreListUB
                    If sText.ToUpper.StartsWith(msIgnoreList(X)) Then Return True
                Next X
            End If
        Catch
        End Try

		Return False
	End Function

    Private Sub btnToggle_Click(ByVal sName As String) Handles btnToggle.Click
        mbIgnoreResize = True
        mbMinimized = Not mbMinimized

        If mbMinimized = True Then
            muSettings.ChatWindowState = 0
            'I am currently maximized, i need to minimize...
            '  if my top > Screen.Height - me.height - 15, then, scoot me down, otherwise, just change my height
            mlNormalHeight = Me.Height
            Me.Height = btnToggle.Height + 2
            If Me.Top > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - mlNormalHeight - 15 Then
                'Me.Top += mlNormalHeight - Me.Height
                Me.Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height - 3
            End If

            With btnToggle
                .ControlImageRect = grc_UI(elInterfaceRectangle.eUpArrow_Button_Normal)
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eUpArrow_Button_Disabled)
                .ControlImageRect_Normal = .ControlImageRect
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eUpArrow_Button_Down)
            End With
        Else
            muSettings.ChatWindowState = 1
            If Me.Top + mlNormalHeight > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then
                Me.Top -= mlNormalHeight - Me.Height
                If Me.Top < 0 Then Me.Top += Math.Abs(Me.Top)
            End If
            Me.Height = mlNormalHeight

            With btnToggle
                .ControlImageRect = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled)
                .ControlImageRect_Normal = .ControlImageRect
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eDownArrow_Button_Down)
            End With
        End If
        Me.IsDirty = True
        mbIgnoreResize = False
    End Sub

    Private Sub txtNew_TextChanged() Handles txtNew.TextChanged
        If txtNew.Caption.ToLower = "/r " Then
            If moTabs(mlCurrentTabIndex).PreviousPMSenders Is Nothing OrElse moTabs(mlCurrentTabIndex).PreviousPMSenders(0) = "" Then Return
            txtNew.Caption = "/pm " & moTabs(mlCurrentTabIndex).PreviousPMSenders(0) & " "
            txtNew.SelStart = txtNew.Caption.Length
            txtNew.CursorPos = txtNew.SelStart
        End If
    End Sub

    Private moDYKRandom As Random = Nothing
    Private mlLastDYKIdx As Int32 = -1
    Private Shared mlJumpsSinceLastDYK As Int32 = 0
    Private Shared msDYK() As String
    Private Shared Sub InitializeDYK()
        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim sBase As String = AppDomain.CurrentDomain.FriendlyName
        Dim lIdx As Int32 = sBase.IndexOf("."c)
        sBase = sBase.Substring(0, lIdx)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.DYK.txt")
        Device.IsUsingEventHandlers = False
        'moEffect = Effect.FromStream(GFXEngine.moDevice, oStream, Nothing, ShaderFlags.None, Nothing)

        Dim lUB As Int32 = -1
        Dim oRead As New IO.StreamReader(oStream)
        While oRead.EndOfStream = False
            Dim sLine As String = oRead.ReadLine
            If sLine <> "" Then
                lUB += 1
                ReDim Preserve msDYK(lUB)
                msDYK(lUB) = sLine
            End If
        End While
        oRead.Close()
        oRead.Dispose()
        oStream.Close()
        oStream.Dispose()
        oAssembly = Nothing
    End Sub

    Public Sub AddADYK()
        If muSettings.lDYK_Frequency = -1 Then Return

        mlJumpsSinceLastDYK += 1
        If mlJumpsSinceLastDYK < muSettings.lDYK_Frequency Then Return
        mlJumpsSinceLastDYK = 0

        If moDYKRandom Is Nothing Then moDYKRandom = New Random()
        If msDYK Is Nothing Then InitializeDYK()

        Dim lIdx As Int32 = moDYKRandom.Next(0, msDYK.Length)
        While lIdx = mlLastDYKIdx
            lIdx = moDYKRandom.Next(0, msDYK.Length)
        End While
        mlLastDYKIdx = lIdx

        AddChatMessage(-1, msDYK(lIdx), ChatMessageType.eSysAdminMessage, Date.MinValue, False)
    End Sub
End Class

'Public Class frmChat
'    Inherits UIWindow

'    Private WithEvents txtNew As UITextBox
'    Private WithEvents scrText As UIScrollBar
'    Private lblText() As UILabel
'    Private lblTitle As UILabel
'    Private lnDiv As UILine

'    Private msLines() As String
'    Private mlSenderID() As Int32
'    Private mlLineUB As Int32 = 100
'    Private mlLineUsed As Int32 = 0

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmChat initial props
'        With Me
'            .ControlName = "frmChat"
'            .Left = 590 'Should be right next to frmContents
'            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - 156     'bottom align
'            .Width = 384
'            .Height = 156
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            '.FillColor = System.Drawing.Color.FromArgb(192, 0, 32, 64)
'            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = True  'always render the chat interface
'            .mbAcceptReprocessEvents = True
'            .BorderLineWidth = 1
'        End With

'        'txtNew initial props
'        txtNew = New UITextBox(oUILib)
'        With txtNew
'            .ControlName = "txtNew"
'            .Left = 1
'            .Top = 137
'            .Width = 382
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            '.ForeColor = System.Drawing.Color.FromArgb(-16777216)
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'            '.BackColorEnabled = System.Drawing.Color.FromArgb(-1)
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
'            .MaxLength = 0
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(txtNew, UIControl))

'        'lblText initial props
'        Dim X As Int32
'        ReDim lblText(5)
'        For X = 0 To lblText.Length - 1
'            lblText(X) = New UILabel(oUILib)
'            With lblText(X)
'                .ControlName = "lblText(" & X & ")"
'                .Left = 3
'                .Top = txtNew.Top - ((X + 1) * 18)
'                .Width = 365
'                .Height = 17
'                .Enabled = True
'                .Visible = True
'                .Caption = ""
'                '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'                .ForeColor = muSettings.InterfaceBorderColor
'                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'                .DrawBackImage = False
'                .FontFormat = DrawTextFormat.Bottom Or DrawTextFormat.Left
'            End With
'            Me.AddChild(CType(lblText(X), UIControl))
'        Next X

'        'lblTitle initial props
'        lblTitle = New UILabel(oUILib)
'        With lblTitle
'            .ControlName = "lblTitle"
'            .Left = 3
'            .Top = 3
'            .Width = 97
'            .Height = 17
'            .Enabled = True
'            .Visible = True
'            .Caption = "Chat Window"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblTitle, UIControl))

'        'lnDiv initial props
'        lnDiv = New UILine(oUILib)
'        With lnDiv
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            .ControlName = "lnDiv"
'            .Enabled = True
'            .Height = 1
'            .Width = Me.Width
'            .Left = 0
'            .Top = lblTitle.Top + lblTitle.Height
'            .Visible = True
'        End With
'        Me.AddChild(CType(lnDiv, UIControl))

'        'scrText initial props
'        scrText = New UIScrollBar(oUILib, True)
'        With scrText
'            .ControlName = "scrText"
'            .Left = 366
'            .Top = lnDiv.Top + 3
'            .Width = 18
'            .Height = txtNew.Top - lnDiv.Top - 3
'            .Enabled = True
'            .Visible = True
'            .Value = 0
'            .MaxValue = 100
'            .MinValue = 0
'            .SmallChange = 1
'            .LargeChange = 1
'            .ReverseDirection = False
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(scrText, UIControl))

'        ReDim msLines(mlLineUB)
'        ReDim mlSenderID(mlLineUB)
'        scrText.MaxValue = mlLineUB


'    End Sub

'    Private Sub txtNew_OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNew.OnKeyUp
'        If e.KeyCode = Keys.Enter Then
'            Dim sNewVal As String = Trim$(txtNew.Caption)
'            txtNew.Caption = ""

'            If sNewVal = "" Then Return

'            If sNewVal.ToLower.StartsWith("/observe") OrElse sNewVal.ToLower.StartsWith("/fps") Then
'                If sNewVal.ToLower.StartsWith("/fps") = True Then sNewVal = "/observe fps"

'                sNewVal = sNewVal.Substring(8).Trim
'                If sNewVal.ToLower.StartsWith("fps") = True Then sNewVal = "-254"

'                Dim lID As Int32 = CInt(Val(sNewVal))

'                If lID = -255 OrElse lID = -254 Then
'                    Dim ofrmObs As frmObserve = New frmObserve(goUILib, lID)
'                    ofrmObs.Visible = True
'                    ofrmObs = Nothing
'                    Return
'                End If

'                If lID < 1 Then
'                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
'                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
'                            lID = goCurrentEnvir.oEntity(X).ObjectID
'                            Exit For
'                        End If
'                    Next X
'                End If

'                If lID > 0 Then
'                    Dim ofrmObs As frmObserve = New frmObserve(goUILib, lID)
'                    ofrmObs.Visible = True
'                    ofrmObs = Nothing
'                End If

'                Return
'            End If

'            If sNewVal.ToLower.StartsWith("/clientver") = True Then
'                AddChatMessage(-1, "Client Version: " & gl_CLIENT_VERSION.ToString)
'                Return
'            End If

'            If sNewVal.ToLower.StartsWith("/h") = True OrElse (sNewVal.ToLower.StartsWith("/") AndAlso sNewVal.Length = 1) Then
'                'ok, display help
'                AddChatMessage(-1, "Chat Help Command List:")
'                AddChatMessage(-1, "Private message: /pm <playername>, <msg>")
'                AddChatMessage(-1, "Create Channel: /create <channelname>")
'                AddChatMessage(-1, "Join Channel: /join <channelname>, <password (optional)>")
'                AddChatMessage(-1, "Give Admin Rights: /admin <channelname>, <playername>")
'                AddChatMessage(-1, "Set Password: /setpassword <channelname>, <password>")
'                AddChatMessage(-1, "Kick Player: /kick <playername>, <channelname>")
'                AddChatMessage(-1, "Invite player: /invite <playername>, <channelnames>")
'                AddChatMessage(-1, "Leave Channel: /leave <channelname>")
'                AddChatMessage(-1, "Alias: /alias <channelname>, <alias>")
'                AddChatMessage(-1, "Chat in Channel: /<channelname> <msg>")
'                AddChatMessage(-1, "Chat in Environment: <msg>")
'                AddChatMessage(-1, "Display Keyboard Command List: /keymap")
'            ElseIf sNewVal.ToLower.StartsWith("/keymap") = True Then
'                AddChatMessage(-1, "Keyboard Command List:")
'                AddChatMessage(-1, "Escape - opens options window")
'                AddChatMessage(-1, "Backspace - Goes up one environment level")
'                AddChatMessage(-1, "Home - Returns to the current environment")
'                AddChatMessage(-1, "C - captures a screenshot")
'                AddChatMessage(-1, "G - toggles the grid in space views")
'                AddChatMessage(-1, "T - Turns on Tracking mode")
'                AddChatMessage(-1, "W - Toggles wireframe mode")
'                AddChatMessage(-1, "F1 - Opens the Quick Help Table of Contents Window")
'                AddChatMessage(-1, "F3 - Opens the Battlegroup Management Window")
'                AddChatMessage(-1, "F4 - Opens the Trade Management Window")
'                AddChatMessage(-1, "F5 - Opens the Diplomacy Window")
'                AddChatMessage(-1, "F6 - Opens the Colony Stats window")
'                AddChatMessage(-1, "F7 - Opens the Budget Window")
'                AddChatMessage(-1, "F12 - Opens the Bug Window")
'                AddChatMessage(-1, "Ctrl+A - Adds the selected item's owner to your Diplomacy list")
'                AddChatMessage(-1, "Ctrl+E - Selects the next engineer")
'                AddChatMessage(-1, "Ctrl+F - Selects the next facility")
'                AddChatMessage(-1, "Ctrl+M - toggles whether mineral caches are rendered")
'                AddChatMessage(-1, "Ctrl+N - Selects the next unit or facility")
'                AddChatMessage(-1, "Ctrl+Q - Immediately closes the game")
'                AddChatMessage(-1, "Ctrl+R - Toggles Set Rally Point Mode")
'                AddChatMessage(-1, "Ctrl+U - Selects the next unit")
'                AddChatMessage(-1, "Ctrl+(0-9) - Assigns selected objects to the control group of that number")
'                AddChatMessage(-1, "(0-9) - Recalls control group selections, clearing the current selection")
'                AddChatMessage(-1, "Shift+(0-9) - Recalls control group selections, without clearing the current selection")
'                AddChatMessage(-1, "Ctrl+Shift+(0-9) - Recalls control group selections without clearing current selections and assigns resulting selects to the control group of that number")
'                AddChatMessage(-1, "~ (tilde) - toggles UI rendering")
'            ElseIf sNewVal.ToLower.StartsWith("/findentity") = True Then
'                Dim sTmp As String = Replace$(sNewVal.ToLower, "/findentity", "")
'                Dim sVals() As String = Split(sTmp.Trim, ",")
'                If sVals.GetUpperBound(0) = 1 Then
'                    Dim lID As Int32 = CInt(Val(sVals(0).Trim))
'                    Dim iTypeID As Int16 = CShort(Val(sVals(1).Trim))

'                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
'                        If goCurrentEnvir.lEntityIdx(X) = lID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iTypeID Then
'                            goCamera.mlCameraAtX = CInt(goCurrentEnvir.oEntity(X).LocX)
'                            goCamera.mlCameraAtZ = CInt(goCurrentEnvir.oEntity(X).LocZ)
'                            goCamera.mlCameraAtY = 0
'                            goCamera.mlCameraX = goCamera.mlCameraAtX
'                            goCamera.mlCameraZ = goCamera.mlCameraAtZ - 1000
'                            goCamera.mlCameraAtY = CInt(goCurrentEnvir.oEntity(X).LocY) + 1000
'                            AddChatMessage(-1, "Entity Found")
'                            Return
'                        End If
'                    Next X

'                    AddChatMessage(-1, "Entity not found")
'                End If
'            Else
'                Dim yMsg() As Byte

'                If sNewVal.StartsWith("/") = False Then
'                    sNewVal = goCurrentPlayer.PlayerName & ": " & sNewVal
'                End If

'                'TODO: if we are to do any pre-send processing, now is the time
'                If sNewVal.ToLower.StartsWith("/gm") = True AndAlso sNewVal.Length > 5 Then
'                    AddChatMessage(-1, goCurrentPlayer.PlayerName & ": " & sNewVal.Substring(4))
'                End If

'                'Prepare our message
'                ReDim yMsg(sNewVal.Length + 1)
'                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(yMsg, 2)
'                System.BitConverter.GetBytes(EpicaMessageCode.eChatMessage).CopyTo(yMsg, 0)

'                'determine who to send it to
'                If sNewVal.StartsWith("/") = True Then
'                    'ok, send to the primary server
'                    MyBase.moUILib.SendMsgToPrimary(yMsg)
'                Else
'                    'ok, send to the region server
'                    MyBase.moUILib.SendMsgToRegion(yMsg)
'                End If
'            End If

'        End If
'    End Sub

'    Public Sub AddChatMessage(ByVal lPlayerID As Int32, ByVal sText As String)
'        ''first, go thru and shift our lines
'        Dim rcTest As Rectangle = lblText(0).GetTextDimensions(sText)
'        While rcTest.Width > lblText(0).Width
'            Dim lStrLen As Int32 = sText.Length
'            Dim lNewLen As Int32 = CInt(Math.Floor(((lblText(0).Width - 5) / rcTest.Width) * lStrLen)) - 1
'            Dim lSplit As Int32 = InStrRev(sText, " ", lNewLen)

'            Dim sLeftVal As String = Mid$(sText, 1, lSplit - 1)
'            AddChatMessage(lPlayerID, sLeftVal)
'            sText = Mid$(sText, lSplit + 1)

'            rcTest = lblText(0).GetTextDimensions(sText)
'        End While

'        Dim X As Int32
'        For X = mlLineUsed To 0 Step -1
'            If X < mlLineUB Then
'                msLines(X + 1) = msLines(X)
'                mlSenderID(X + 1) = mlSenderID(X)
'            End If
'        Next X
'        If mlLineUsed <> mlLineUB Then mlLineUsed += 1
'        msLines(0) = sText
'        mlSenderID(0) = lPlayerID

'        RefreshView()
'    End Sub

'    Private Sub RefreshView()
'        Dim X As Int32
'        Dim lIdx As Int32

'        'Now, fill our labels
'        For X = 0 To lblText.Length - 1
'            lIdx = scrText.Value + X
'            If lIdx > mlLineUB Then
'                lblText(X).Caption = ""
'            Else
'                lblText(X).Caption = msLines(lIdx)
'                Select Case mlSenderID(lIdx)
'                    Case 0
'                        'System text
'                        lblText(X).ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'                    Case -1
'                        'Notification
'                        lblText(X).ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
'                    Case Is = glPlayerID
'                        'my own text
'                        lblText(X).ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'                    Case Else
'                        lblText(X).ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'                End Select
'            End If

'        Next X
'    End Sub

'    Private Sub scrText_ValueChange() Handles scrText.ValueChange
'        RefreshView()
'    End Sub

'End Class
