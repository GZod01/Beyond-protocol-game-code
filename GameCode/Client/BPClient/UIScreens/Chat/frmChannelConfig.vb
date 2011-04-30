Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmChannelConfig
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private lblChannelName As UILabel
	Private txtChannelName As UITextBox
	Private lblPassword As UILabel
	Private txtPassword As UITextBox
	Private lnDiv2 As UILine
	Private chkPublic As UICheckBox
	Private lblMembers As UILabel
	Private lstMembers As UIListBox
	Private lnDiv3 As UILine
	Private txtInvite As UITextBox

	Private WithEvents lstContacts As UIListBox
	Private WithEvents btnGiveAdmin As UIButton
	Private WithEvents btnInvite As UIButton
	Private WithEvents btnUpdate As UIButton
	Private WithEvents btnKick As UIButton
	Private WithEvents btnClose As UIButton

	Private msChannelName As String
	Private msOriginalPassword As String = ""
    Private mbOriginalPublic As Boolean = False
    Private mlChannelID As Int32 = -1
    Private mlMemberIDs() As Int32 = Nothing

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenChatConfigWindow)

		'frmChannelConfig initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eChannelConfig
            .ControlName = "frmChannelConfig"
            .Left = 377
            .Top = 134
            .Width = 256
            .Height = 450
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 5
			.Width = 156
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Channel Configuration"
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
			.Left = 2
			.Top = 25
			.Width = 255
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'lblChannelName initial props
		lblChannelName = New UILabel(oUILib)
		With lblChannelName
			.ControlName = "lblChannelName"
			.Left = 5
			.Top = 35
			.Width = 100
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Channel Name:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblChannelName, UIControl))

		'txtChannelName initial props
		txtChannelName = New UITextBox(oUILib)
		With txtChannelName
			.ControlName = "txtChannelName"
			.Left = 105
			.Top = 35
			.Width = 145
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtChannelName, UIControl))

		'lblPassword initial props
		lblPassword = New UILabel(oUILib)
		With lblPassword
			.ControlName = "lblPassword"
			.Left = 5
			.Top = 60
			.Width = 69
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Password:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblPassword, UIControl))

		'txtPassword initial props
		txtPassword = New UITextBox(oUILib)
		With txtPassword
			.ControlName = "txtPassword"
			.Left = 105
			.Top = 60
			.Width = 145
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtPassword, UIControl))

		'lnDiv2 initial props
		lnDiv2 = New UILine(oUILib)
		With lnDiv2
			.ControlName = "lnDiv2"
			.Left = 1
			.Top = 135
			.Width = 255
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv2, UIControl))

		'chkPublic initial props
		chkPublic = New UICheckBox(oUILib)
		With chkPublic
			.ControlName = "chkPublic"
			.Left = 35
			.Top = 83
			.Width = 178
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Channel Viewable to Public"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		Me.AddChild(CType(chkPublic, UIControl))

		'lblMembers initial props
		lblMembers = New UILabel(oUILib)
		With lblMembers
			.ControlName = "lblMembers"
			.Left = 5
			.Top = 140
			.Width = 125
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Channel Members"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMembers, UIControl))

		'lstMembers initial props
		lstMembers = New UIListBox(oUILib)
		With lstMembers
			.ControlName = "lstMembers"
			.Left = 5
			.Top = 159
			.Width = 245
			.Height = 120
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		Me.AddChild(CType(lstMembers, UIControl))

		'btnKick initial props
		btnKick = New UIButton(oUILib)
		With btnKick
			.ControlName = "btnKick"
			.Left = 152
			.Top = 285
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Kick Member"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnKick, UIControl))

		'lnDiv3 initial props
		lnDiv3 = New UILine(oUILib)
		With lnDiv3
			.ControlName = "lnDiv3"
			.Left = 1
			.Top = 315
			.Width = 255
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv3, UIControl))

		'lstContacts initial props
		lstContacts = New UIListBox(oUILib)
		With lstContacts
			.ControlName = "lstContacts"
			.Left = 5
			.Top = 320
			.Width = 245
			.Height = 100
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		Me.AddChild(CType(lstContacts, UIControl))

		'txtInvite initial props
		txtInvite = New UITextBox(oUILib)
		With txtInvite
			.ControlName = "txtInvite"
			.Left = 5
			.Top = 427
			.Width = 145
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtInvite, UIControl))

		'btnGiveAdmin initial props
		btnGiveAdmin = New UIButton(oUILib)
		With btnGiveAdmin
			.ControlName = "btnGiveAdmin"
			.Left = 5
			.Top = 285
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Make Admin"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnGiveAdmin, UIControl))

		'btnInvite initial props
		btnInvite = New UIButton(oUILib)
		With btnInvite
			.ControlName = "btnInvite"
			.Left = 152
			.Top = 425
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Invite"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnInvite, UIControl))

		'btnUpdate initial props
		btnUpdate = New UIButton(oUILib)
		With btnUpdate
			.ControlName = "btnUpdate"
			.Left = 30
			.Top = 105
			.Width = 200
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Update Configuration"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnUpdate, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

    Private Sub SubmitConfigMessage(ByVal yType As eyChatRoomCommandType, ByVal yPublicFlag As Byte, ByVal lTargetID As Int32, ByVal sPassword As String, ByVal sPlayerToInvite As String)

        lTargetID = Math.Abs(lTargetID)

        Dim lLen As Int32 = 0
        If (yType And eyChatRoomCommandType.SetChannelPublic) <> 0 Then
            lLen += 1
        End If
        If (yType And (eyChatRoomCommandType.ToggleAdminRights Or eyChatRoomCommandType.KickPlayer)) <> 0 Then
            lLen += 4
        End If
        If (yType And (eyChatRoomCommandType.JoinChannel Or eyChatRoomCommandType.SetChannelPassword)) <> 0 Then
            lLen += 20
        End If
        If (yType And eyChatRoomCommandType.InvitePlayer) <> 0 Then
            lLen += 30
        End If

        Dim yMsg(26 + lLen) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yMsg, lPos) : lPos += 2
        lPos += 4       'leave room for playerid
        yMsg(lPos) = yType : lPos += 1
        System.Text.ASCIIEncoding.ASCII.GetBytes(msChannelName).CopyTo(yMsg, lPos) : lPos += 20

        If (yType And eyChatRoomCommandType.SetChannelPublic) <> 0 Then
            yMsg(lPos) = yPublicFlag : lPos += 1
        End If
        If (yType And (eyChatRoomCommandType.ToggleAdminRights Or eyChatRoomCommandType.KickPlayer)) <> 0 Then
            System.BitConverter.GetBytes(lTargetID).CopyTo(yMsg, lPos) : lPos += 4
        End If
        If (yType And (eyChatRoomCommandType.JoinChannel Or eyChatRoomCommandType.SetChannelPassword)) <> 0 Then
            System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword).CopyTo(yMsg, lPos) : lPos += 20
        End If
        If (yType And eyChatRoomCommandType.InvitePlayer) <> 0 Then
            System.Text.ASCIIEncoding.ASCII.GetBytes(sPlayerToInvite).CopyTo(yMsg, lPos) : lPos += 30
        End If

        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

	Private Sub btnGiveAdmin_Click(ByVal sName As String) Handles btnGiveAdmin.Click
		If lstMembers Is Nothing = False AndAlso lstMembers.ListIndex > -1 Then
            Dim sMember As String = lstMembers.List(lstMembers.ListIndex)
            Dim lMemberID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
            If sMember <> "" Then
                SubmitConfigMessage(eyChatRoomCommandType.ToggleAdminRights, 0, lMemberID, "", "")
                ''/admin <channelname>, <playername>
                'Dim sText As String = "/admin " & msChannelName & ", " & sMember
                'Dim yMsg(sText.Length + 1) As Byte
                'System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
                'System.Text.ASCIIEncoding.ASCII.GetBytes(sText).CopyTo(yMsg, 2)
                'MyBase.moUILib.SendMsgToPrimary(yMsg)

                MyBase.moUILib.AddNotification("Attempting to grant " & sMember & " admin rights for " & msChannelName & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
		End If
	End Sub

	Private Sub btnInvite_Click(ByVal sName As String) Handles btnInvite.Click
		If txtInvite.Caption Is Nothing = False AndAlso txtInvite.Caption.Trim <> "" Then
			'/invite <playername>, <channelname>")
			'Prepare our message
            'Dim sText As String = "/invite " & txtInvite.Caption & ", " & msChannelName
            'Dim yMsg(sText.Length + 1) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
            'System.Text.ASCIIEncoding.ASCII.GetBytes(sText).CopyTo(yMsg, 2)
            'MyBase.moUILib.SendMsgToPrimary(yMsg)
            SubmitConfigMessage(eyChatRoomCommandType.InvitePlayer, 0, 0, "", txtInvite.Caption)

			MyBase.moUILib.AddNotification("Invitation sent to " & txtInvite.Caption & " to join " & msChannelName & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
	End Sub

	Private Sub btnKick_Click(ByVal sName As String) Handles btnKick.Click
		If lstMembers Is Nothing = False AndAlso lstMembers.ListIndex > -1 Then
            Dim sMember As String = lstMembers.List(lstMembers.ListIndex)
            Dim lMemberID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
			If sMember <> "" Then
				'/admin <channelname>, <playername>
                'Dim sText As String = "/kick " & sMember & ", " & msChannelName
                'Dim yMsg(sText.Length + 1) As Byte
                'System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
                'System.Text.ASCIIEncoding.ASCII.GetBytes(sText).CopyTo(yMsg, 2)
                'MyBase.moUILib.SendMsgToPrimary(yMsg)
                SubmitConfigMessage(eyChatRoomCommandType.KickPlayer, 0, lMemberID, "", "")

				MyBase.moUILib.AddNotification("Attempting to kick " & sMember & " from " & msChannelName & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If
		End If
	End Sub

	Private Sub btnUpdate_Click(ByVal sName As String) Handles btnUpdate.Click
		If msChannelName Is Nothing OrElse msChannelName = "" Then
			'"Create Channel: /create <channelname>")
			msChannelName = txtChannelName.Caption
			If msChannelName Is Nothing OrElse msChannelName.Trim = "" Then Return

            'Dim sText As String = "/create " & msChannelName

            ''If chkPublic.Value = True Then
            ''	sText &= ", 1"
            ''Else : sText &= ", 0"
            ''End If
            ''mbOriginalPublic = chkPublic.Value

            ''If txtPassword.Caption Is Nothing = False AndAlso txtPassword.Caption <> "" Then
            ''	sText &= ", " & txtPassword.Caption
            ''	msOriginalPassword = txtPassword.Caption
            ''End If

            'Dim yMsg(sText.Length + 1) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
            'System.Text.ASCIIEncoding.ASCII.GetBytes(sText).CopyTo(yMsg, 2)
            'MyBase.moUILib.SendMsgToPrimary(yMsg)

            Dim yPublic As Byte = 0
            Dim sPassword As String = ""
            If chkPublic.Value = True Then yPublic = 1
            If txtPassword.Caption Is Nothing = False AndAlso txtPassword.Caption <> "" Then sPassword = txtPassword.Caption
            SubmitConfigMessage(eyChatRoomCommandType.AddNewChatRoom Or eyChatRoomCommandType.SetChannelPublic Or eyChatRoomCommandType.SetChannelPassword, yPublic, -1, sPassword, "")
        Else
            Dim yType As eyChatRoomCommandType = 0
            If txtPassword.Caption Is Nothing = False Then
                If txtPassword.Caption <> "" Then
                    If msOriginalPassword <> txtPassword.Caption Then
                        yType = eyChatRoomCommandType.SetChannelPassword
                    End If
                End If
            End If
            If mbOriginalPublic <> chkPublic.Value Then
                mbOriginalPublic = chkPublic.Value
                yType = yType Or eyChatRoomCommandType.SetChannelPublic
            End If
            If yType <> 0 Then
                Dim yPublic As Byte = 0
                If chkPublic.Value = True Then yPublic = 1
                SubmitConfigMessage(yType, yPublic, -1, txtPassword.Caption, "")
            End If
        End If
	End Sub

	Private Sub lstContacts_ItemClick(ByVal lIndex As Integer) Handles lstContacts.ItemClick
		If lIndex > -1 Then
			txtInvite.Caption = lstContacts.List(lIndex)
		End If
	End Sub

    Public Sub SetFromChannel(ByVal sChannel As String, ByVal bPublic As Boolean, ByVal lID As Int32)
        msChannelName = sChannel
        mbOriginalPublic = bPublic
        mlChannelID = lID

        If msChannelName Is Nothing Then msChannelName = ""
        If msChannelName = "" Then Return

        Dim yMsg(35) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelDetails).CopyTo(yMsg, lPos) : lPos += 2
        lPos += 4        'for the player id appended at primary
        'System.BitConverter.GetBytes(mlChannelID).CopyTo(yMsg, lPos) : lPos += 4
        System.Text.ASCIIEncoding.ASCII.GetBytes(msChannelName).CopyTo(yMsg, lPos) : lPos += 30
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Public Sub HandleChannelDetails(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sChannelName As String = GetStringFromBytes(yData, lPos, 30) : lPos += 30

        If msChannelName.ToUpper <> sChannelName.ToUpper Then Return

        msOriginalPassword = GetStringFromBytes(yData, lPos, 30) : lPos += 30
        Dim yChannelStatus As Byte = yData(lPos) : lPos += 1

        mbOriginalPublic = (yChannelStatus And 1) <> 0

        Dim lMemberCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMemberIDs(lMemberCnt - 1) As Int32
        For X As Int32 = 0 To lMemberCnt - 1
            lMemberIDs(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'negative ID's are admins
        Next X

        'Now, what to do with it all....
        mlMemberIDs = lMemberIDs

        txtChannelName.Caption = msChannelName
        txtPassword.Caption = msOriginalPassword
        chkPublic.Value = mbOriginalPublic
    End Sub
 
    Private Sub frmChannelConfig_OnNewFrame() Handles Me.OnNewFrame
        If mlMemberIDs Is Nothing = False Then
            If mlMemberIDs.Length <> lstMembers.ListCount Then
                FillMembersList()
            Else
                Dim bChanged As Boolean = False
                For X As Int32 = 0 To lstMembers.ListCount - 1
                    Dim lID As Int32 = lstMembers.ItemData(X)
                    Dim sTemp As String = GetCacheObjectValue(Math.Abs(lID), ObjectType.ePlayer)
                    If lstMembers.List(X) <> sTemp Then
                        lstMembers.List(X) = sTemp
                        bChanged = True
                    End If
                    If lID < 0 Then
                        If lstMembers.ItemBold(X) <> True Then lstMembers.ItemBold(X) = True
                    ElseIf lstMembers.ItemBold(X) <> False Then
                        lstMembers.ItemBold(X) = False
                    End If
                Next X

                If bChanged = True Then lstMembers.SortList(False, False)
            End If
        ElseIf lstMembers.ListCount <> 0 Then
            lstMembers.Clear()
        End If

        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.PlayerRelUB + 1 <> lstContacts.ListCount Then
            FillRelsList()
            lstContacts.SortList(False, False)
        Else
            Dim bChanged As Boolean = False
            For X As Int32 = 0 To lstContacts.ListCount - 1
                Dim sTemp As String = GetCacheObjectValue(lstContacts.ItemData(X), ObjectType.ePlayer)
                If lstContacts.List(X) <> sTemp Then
                    lstContacts.List(X) = sTemp
                    bChanged = True
                End If
            Next X
            If bChanged = True Then lstContacts.SortList(False, False)
        End If
    End Sub

    Private Sub FillMembersList()
        lstMembers.Clear()
        If mlMemberIDs Is Nothing = False Then
            For X As Int32 = 0 To mlMemberIDs.GetUpperBound(0)
                Dim lID As Int32 = mlMemberIDs(X)
                Dim sTemp As String = GetCacheObjectValue(Math.Abs(lID), ObjectType.ePlayer)
                lstMembers.AddItem(sTemp, lID < 0)
                lstMembers.ItemData(lstMembers.NewIndex) = lID
            Next X
        End If
    End Sub

    Private Sub FillRelsList()
        lstContacts.Clear()

        'Let's load the player's rels now...
        Dim oTmpRel As PlayerRel
        Dim sName As String

        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
                oTmpRel = goCurrentPlayer.GetPlayerRelByIndex(X)
                If oTmpRel Is Nothing = False Then
                    sName = GetCacheObjectValue(oTmpRel.lThisPlayer, ObjectType.ePlayer)
                    lstContacts.AddItem(sName)
                    lstContacts.ItemData(lstContacts.NewIndex) = oTmpRel.lThisPlayer
                End If
            Next X

        End If
    End Sub
End Class