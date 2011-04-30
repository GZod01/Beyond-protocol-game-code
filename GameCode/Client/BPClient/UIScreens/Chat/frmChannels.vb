Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmChannels
	Inherits UIWindow

    Private Enum eyChatRoomAttr As Byte
        PublicChannel = 1
        UserIsAdmin = 2
        AllAreAdmin = 4
        PasswordProtected = 8
        PlayerIsPermitted = 16
        PlayerInChannel = 32
    End Enum

	Private lblTitle As UILabel
	Private WithEvents btnClose As UIButton
	Private lnDiv1 As UILine
    Private WithEvents lstChannels As UIListBox
	Private WithEvents btnJoin As UIButton
	Private WithEvents btnCreate As UIButton
    Private WithEvents btnAdmin As UIButton

    Private Structure ChatRoomListing
        Public sRoomName As String
        Public lID As Int32
        Public yAttr As Byte
        Public lMemberCount As Int32
    End Structure
    Private muRooms() As ChatRoomListing = Nothing

    Private mlLastUpdate As Int32 = 0

    Private mfraPassword As fraChannelPassword = Nothing

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmChannels initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eChannels
            .ControlName = "frmChannels"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 256
            .Height = 256
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
			.Width = 135
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Available Channels"
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
			.Left = 1
			.Top = 25
			.Width = 255
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'lstChannels initial props
		lstChannels = New UIListBox(oUILib)
		With lstChannels
			.ControlName = "lstChannels"
			.Left = 5
			.Top = 30
			.Width = 245
			.Height = 155
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		Me.AddChild(CType(lstChannels, UIControl))

		'btnJoin initial props
		btnJoin = New UIButton(oUILib)
		With btnJoin
			.ControlName = "btnJoin"
			.Left = 150
			.Top = 195
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Join Channel"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnJoin, UIControl))

		'btnCreate initial props
		btnCreate = New UIButton(oUILib)
		With btnCreate
			.ControlName = "btnCreate"
			.Left = 5
			.Top = 195
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Create New"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCreate, UIControl))

		'btnAdmin initial props
		btnAdmin = New UIButton(oUILib)
		With btnAdmin
			.ControlName = "btnAdmin"
			.Left = 75
			.Top = 225
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Admin"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnAdmin, UIControl))

        RequestRoomList()

		MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Private Sub RequestRoomList()
        Dim yMsg(5) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelList).CopyTo(yMsg, lPos) : lPos += 2
        MyBase.moUILib.SendMsgToPrimary(yMsg)
        mlLastUpdate = glCurrentCycle
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Function GetSelectedRoom(ByRef bFound As Boolean) As ChatRoomListing
        If lstChannels.ListIndex > -1 AndAlso muRooms Is Nothing = False Then
            'Dim lID As Int32 = lstChannels.ItemData(lstChannels.ListIndex)
            Dim sRoomName As String = lstChannels.List(lstChannels.ListIndex).ToUpper
            For X As Int32 = 0 To muRooms.GetUpperBound(0)
                If muRooms(X).sRoomName.ToUpper = sRoomName Then
                    bFound = True
                    Return muRooms(X)
                End If
            Next X
        End If

        MyBase.moUILib.AddNotification("Select a room in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        bFound = False
        Return Nothing
    End Function

    Private Sub btnAdmin_Click(ByVal sName As String) Handles btnAdmin.Click
        'do I have rights to admin the selected channel?
        Dim bFound As Boolean = False
        Dim uRoom As ChatRoomListing = GetSelectedRoom(bFound)
        If bFound = True Then
            If (uRoom.yAttr And (eyChatRoomAttr.AllAreAdmin Or eyChatRoomAttr.UserIsAdmin)) <> 0 Then
                'ok, let's admin it
                Dim ofrm As frmChannelConfig = New frmChannelConfig(MyBase.moUILib)
                With uRoom
                    ofrm.SetFromChannel(.sRoomName, (.yAttr And eyChatRoomAttr.PublicChannel) <> 0, .lID)
                End With
                ofrm.Visible = True
            Else
                MyBase.moUILib.AddNotification("You lack admin rights for that room.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If

    End Sub

    Private Sub btnCreate_Click(ByVal sName As String) Handles btnCreate.Click
        Dim ofrm As New frmChannelConfig(MyBase.moUILib)
        ofrm.SetFromChannel("", True, -1)
        ofrm.Visible = True
    End Sub

    Private Sub btnJoin_Click(ByVal sName As String) Handles btnJoin.Click
        Dim bFound As Boolean = False
        Dim uRoom As ChatRoomListing = GetSelectedRoom(bFound)
        If bFound = True Then

            If btnJoin.Caption = "Join Channel" Then
                If (uRoom.yAttr And eyChatRoomAttr.PasswordProtected) <> 0 Then
                    'ok, need to get the password
                    If mfraPassword Is Nothing Then
                        mfraPassword = New fraChannelPassword(MyBase.moUILib)
                        Me.AddChild(CType(mfraPassword, UIControl))
                        AddHandler mfraPassword.PasswordEntered, AddressOf PasswordEntered
                        AddHandler mfraPassword.PasswordDialogClosed, AddressOf PasswordDialogClosed
                    End If
                    mfraPassword.Left = Me.Width \ 2 - mfraPassword.Width \ 2
                    mfraPassword.Top = Me.Height \ 2 - mfraPassword.Height \ 2
                    mfraPassword.Visible = True

                    lstChannels.Enabled = False
                    btnJoin.Enabled = False
                    btnCreate.Enabled = False
                    btnAdmin.Enabled = False
                    btnClose.Enabled = False
                Else
                    JoinRoom(uRoom, Nothing)
                End If
            Else
                LeaveRoom(uRoom)
            End If

        End If
    End Sub

    Private Sub JoinRoom(ByVal uRoom As ChatRoomListing, ByVal sPassword As String)
        'If sPassword Is Nothing Then sPassword = "" Else sPassword = ", " & sPassword
        If sPassword Is Nothing Then sPassword = ""
        'Dim sNewVal As String = "/join " & uRoom.sRoomName & sPassword
        'Dim yData(sNewVal.Length + 1) As Byte
        'System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(yData, 2)
        'System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yData, 0)
        'MyBase.moUILib.SendMsgToPrimary(yData)
        Dim yData(46) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
        lPos += 4 'leave room for playerid
        yData(lPos) = eyChatRoomCommandType.JoinChannel : lPos += 1
        System.Text.ASCIIEncoding.ASCII.GetBytes(uRoom.sRoomName).CopyTo(yData, lPos) : lPos += 20
        System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword).CopyTo(yData, lPos) : lPos += 20
        MyBase.moUILib.SendMsgToPrimary(yData)

        MyBase.moUILib.AddNotification("Join Request sent to join " & uRoom.sRoomName & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub LeaveRoom(ByVal uRoom As ChatRoomListing)

        Dim yData(46) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
        lPos += 4 'leave room for playerid
        yData(lPos) = eyChatRoomCommandType.LeaveChatRoom : lPos += 1
        System.Text.ASCIIEncoding.ASCII.GetBytes(uRoom.sRoomName).CopyTo(yData, lPos) : lPos += 20
        System.Text.ASCIIEncoding.ASCII.GetBytes("").CopyTo(yData, lPos) : lPos += 20
        MyBase.moUILib.SendMsgToPrimary(yData)

        MyBase.moUILib.AddNotification("Leave Request sent to join " & uRoom.sRoomName & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub PasswordEntered(ByVal sValue As String)
        Dim bFound As Boolean = False
        Dim uRoom As ChatRoomListing = GetSelectedRoom(bFound)
        JoinRoom(uRoom, sValue)
    End Sub
    Private Sub PasswordDialogClosed()
        lstChannels.Enabled = True
        btnJoin.Enabled = True
        btnCreate.Enabled = True
        btnAdmin.Enabled = True
        btnClose.Enabled = True
    End Sub

    Public Sub HandleRequestChannelList(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode

        'Roomname(30)
        'lID (4)
        'Attrs (1)
        'MemberCount (4)
        lPos += 4       'for playerid
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim uNewList(lCnt - 1) As ChatRoomListing
        For X As Int32 = 0 To lCnt - 1
            With uNewList(X)
                .sRoomName = GetStringFromBytes(yData, lPos, 30) : lPos += 30
                '.lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yAttr = yData(lPos) : lPos += 1
                .lMemberCount = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With
            
        Next X

        'Now, just set our room list
        muRooms = uNewList
        FillRoomList()
    End Sub

    Private Sub SortRoomList()
        Try

            Dim lSorted(muRooms.GetUpperBound(0)) As Int32
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To muRooms.GetUpperBound(0)
                Dim sThisRoom As String = muRooms(X).sRoomName.ToUpper
                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To lSortedUB
                    If muRooms(Y).sRoomName.ToUpper > sThisRoom Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y

                lSortedUB += 1
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X

            If lSortedUB = muRooms.GetUpperBound(0) Then
                Dim uTmp(lSortedUB) As ChatRoomListing
                For X As Int32 = 0 To lSortedUB
                    uTmp(X) = muRooms(lSorted(X))
                Next X
                muRooms = uTmp
            End If

        Catch
        End Try
    End Sub

    Private Sub FillRoomList()
        lstChannels.Clear()
        lstChannels.oIconTexture = MyBase.moUILib.oInterfaceTexture
        lstChannels.RenderIcons = True

        If muRooms Is Nothing = False Then
            Try
                SortRoomList()

                For X As Int32 = 0 To muRooms.GetUpperBound(0)
                    lstChannels.AddItem(muRooms(X).sRoomName, False)
                    lstChannels.ApplyIconOffset(lstChannels.NewIndex) = False
                    lstChannels.rcIconRectangle(lstChannels.NewIndex) = Rectangle.Empty

                    lstChannels.ApplyIconOffset(lstChannels.NewIndex) = True

                    If (muRooms(X).yAttr And eyChatRoomAttr.PlayerInChannel) <> 0 Then
                        lstChannels.rcIconRectangle(lstChannels.NewIndex) = New Rectangle(112, 159, 13, 13)
                        lstChannels.IconForeColor(lstChannels.NewIndex) = System.Drawing.Color.FromArgb(255, 64, 128, 255)
                    ElseIf (muRooms(X).yAttr And eyChatRoomAttr.PasswordProtected) <> 0 Then
                        lstChannels.rcIconRectangle(lstChannels.NewIndex) = New Rectangle(209, 65, 14, 14)
                        lstChannels.IconForeColor(lstChannels.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
                    End If

                Next X
            Catch
                lstChannels.Clear()
            End Try
        End If
    End Sub

    Private Class fraChannelPassword
        Inherits UIWindow

        Public Event PasswordDialogClosed()
        Public Event PasswordEntered(ByVal sValue As String)

        Private lblPassword As UILabel
        Private txtPassword As UITextBox
        Private WithEvents btnSubmit As UIButton
        Private WithEvents btnCancel As UIButton
        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'frmChannelPassword initial props
            With Me
                .ControlName = "fraChannelPassword"
                .Left = 351
                .Top = 222
                .Width = 170
                .Height = 115
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Moveable = False
            End With

            'lblPassword initial props
            lblPassword = New UILabel(oUILib)
            With lblPassword
                .ControlName = "lblPassword"
                .Left = 5
                .Top = 5
                .Width = 133
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Channel Password:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPassword, UIControl))

            'txtPassword initial props
            txtPassword = New UITextBox(oUILib)
            With txtPassword
                .ControlName = "txtPassword"
                .Left = 15
                .Top = 30
                .Width = 140
                .Height = 20
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
                .PasswordChar = "*"
            End With
            Me.AddChild(CType(txtPassword, UIControl))

            'btnSubmit initial props
            btnSubmit = New UIButton(oUILib)
            With btnSubmit
                .ControlName = "btnSubmit"
                .Left = 35
                .Top = 55
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Submit"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSubmit, UIControl))

            'btnCancel initial props
            btnCancel = New UIButton(oUILib)
            With btnCancel
                .ControlName = "btnCancel"
                .Left = 35
                .Top = 85
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
        End Sub

        Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
            RaiseEvent PasswordEntered(txtPassword.Caption)
            CloseMe()
        End Sub

        Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
            CloseMe()
        End Sub

        Private Sub CloseMe()
            Me.Visible = False
            RaiseEvent PasswordDialogClosed()
        End Sub
    End Class

    Private Sub frmChannels_OnNewFrame() Handles Me.OnNewFrame
        If lstChannels Is Nothing = False AndAlso muRooms Is Nothing = False Then
            Try
                If muRooms.Length <> lstChannels.ListCount Then
                    FillRoomList()
                End If
            Catch
            End Try
        End If

        If glCurrentCycle - mlLastUpdate > 90 Then
            RequestRoomList()
        End If
    End Sub

    Private Sub lstChannels_ItemClick(ByVal lIndex As Integer) Handles lstChannels.ItemClick
        If lIndex > -1 Then
            Dim bFound As Boolean = False
            Dim uRoom As ChatRoomListing = GetSelectedRoom(bFound)
            If bFound = False Then Return
            If (uRoom.yAttr And eyChatRoomAttr.PlayerInChannel) <> 0 Then
                If btnJoin.Caption <> "Leave Channel" Then btnJoin.Caption = "Leave Channel"
            ElseIf btnJoin.Caption <> "Join Channel" Then
                btnJoin.Caption = "Join Channel"
            End If
        End If
    End Sub
End Class
