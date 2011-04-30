Public Class frmChannels


    Private Enum eyChatRoomAttr As Byte
        PublicChannel = 1
        UserIsAdmin = 2
        AllAreAdmin = 4
        PasswordProtected = 8
        PlayerIsPermitted = 16
        PlayerInChannel = 32
    End Enum

    Private Class ChatRoomListing
        Public sRoomName As String
        Public lID As Int32
        Public yAttr As Byte
        Public lMemberCount As Int32
    End Class

    Public Sub RequestRoomList()
        Dim yMsg(5) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelList).CopyTo(yMsg, lPos) : lPos += 2
        If gfrmChat Is Nothing = False Then gfrmChat.SendMsgToPrimary(yMsg)
        'mlLastUpdate = glCurrentCycle
    End Sub

    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        If gfrmChannelConfig Is Nothing = False Then
            gfrmChannelConfig.Close()
            gfrmChannelConfig = Nothing
        End If
        gfrmChannelConfig = New frmChannelConfig
        gfrmChannelConfig.SetFromChannel("", True, -1)
        gfrmChannelConfig.Visible = True
    End Sub
    Private Sub btnAdmin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdmin.Click
        'do I have rights to admin the selected channel?
        Dim lIdx As Int32 = lstChannels.SelectedIndex
        If moRooms Is Nothing = False AndAlso lIdx > -1 AndAlso lIdx < moRooms.Length Then
            Dim uRoom As ChatRoomListing = moRooms(lIdx)  'CType(lstChannels.SelectedItem, ChatRoomListing)
            If uRoom Is Nothing = False Then
                If (uRoom.yAttr And (eyChatRoomAttr.AllAreAdmin Or eyChatRoomAttr.UserIsAdmin)) <> 0 Then
                    'ok, let's admin it
                    If gfrmChannelConfig Is Nothing = False Then
                        gfrmChannelConfig.Close()
                        gfrmChannelConfig = Nothing
                    End If
                    gfrmChannelConfig = New frmChannelConfig()
                    With uRoom
                        gfrmChannelConfig.SetFromChannel(.sRoomName, (.yAttr And eyChatRoomAttr.PublicChannel) <> 0, .lID)
                    End With
                    gfrmChannelConfig.Show(Me)
                Else
                    MsgBox("You lack admin rights for that room.", MsgBoxStyle.OkOnly, "Insufficient Rights")
                End If
            End If
        End If
        
    End Sub
    Private Sub btnJoinLeave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJoinLeave.Click
        Dim lIdx As Int32 = lstChannels.SelectedIndex
        If moRooms Is Nothing = False AndAlso lIdx > -1 AndAlso lIdx < moRooms.Length Then
            Dim uRoom As ChatRoomListing = moRooms(lIdx) 'CType(lstChannels.SelectedItem, ChatRoomListing)
            If uRoom Is Nothing = False Then
                If btnJoinLeave.Text = "Join Channel" Then
                    Dim sPW As String = Nothing
                    If (uRoom.yAttr And eyChatRoomAttr.PasswordProtected) <> 0 Then
                        sPW = InputBox("This room is password protected, what is the password?", "Password", Nothing)
                    End If
                    JoinRoom(uRoom, sPW)
                Else
                    LeaveRoom(uRoom)
                End If
            End If
        End If

    End Sub

    Private Sub JoinRoom(ByVal uRoom As ChatRoomListing, ByVal sPassword As String)
        'If sPassword Is Nothing Then sPassword = "" Else sPassword = ", " & sPassword
        If sPassword Is Nothing Then sPassword = ""
         
        Dim yData(46) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
        lPos += 4 'leave room for playerid
        yData(lPos) = eyChatRoomCommandType.JoinChannel : lPos += 1
        System.Text.ASCIIEncoding.ASCII.GetBytes(uRoom.sRoomName).CopyTo(yData, lPos) : lPos += 20
        System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword).CopyTo(yData, lPos) : lPos += 20
        If gfrmChat Is Nothing = False Then gfrmChat.SendMsgToPrimary(yData)
    End Sub

    Private Sub LeaveRoom(ByVal uRoom As ChatRoomListing)

        Dim yData(46) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eChatChannelCommand).CopyTo(yData, lPos) : lPos += 2
        lPos += 4 'leave room for playerid
        yData(lPos) = eyChatRoomCommandType.LeaveChatRoom : lPos += 1
        System.Text.ASCIIEncoding.ASCII.GetBytes(uRoom.sRoomName).CopyTo(yData, lPos) : lPos += 20
        System.Text.ASCIIEncoding.ASCII.GetBytes("").CopyTo(yData, lPos) : lPos += 20
        If gfrmChat Is Nothing = False Then gfrmChat.SendMsgToPrimary(yData)
    End Sub


    Private Delegate Sub doFillRoomList()
    Private moRooms() As ChatRoomListing
    Private Sub FillRoomList()
        lstChannels.Items.Clear()

        If moRooms Is Nothing Then Return

        Try
            For X As Int32 = 0 To moRooms.GetUpperBound(0)
                lstChannels.Items.Add(moRooms(X).sRoomName)
            Next X
        Catch
            lstChannels.Items.Clear()
        End Try
    End Sub



    Public Sub HandleRequestChannelList(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode

        'Roomname(30)
        'lID (4)
        'Attrs (1)
        'MemberCount (4)
        lPos += 4       'for playerid
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oNewList(lCnt - 1) As ChatRoomListing
        For X As Int32 = 0 To lCnt - 1
            oNewList(X) = New ChatRoomListing
            With oNewList(X)
                .sRoomName = GetStringFromBytes(yData, lPos, 30) : lPos += 30
                '.lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yAttr = yData(lPos) : lPos += 1
                .lMemberCount = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With

        Next X

        'Now, just set our room list
        moRooms = oNewList
        lstChannels.Invoke(New doFillRoomList(AddressOf FillRoomList))
        'FillRoomList(oNewList)
    End Sub

    Private Sub lstChannels_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstChannels.SelectedIndexChanged

        Dim lIdx As Int32 = lstChannels.SelectedIndex
        If moRooms Is Nothing = False AndAlso lIdx > -1 AndAlso lIdx < moRooms.Length Then
            Dim uRoom As ChatRoomListing = moRooms(lIdx) 'CType(lstChannels.SelectedItem, ChatRoomListing)

            If uRoom Is Nothing Then Return
            If (uRoom.yAttr And eyChatRoomAttr.PlayerInChannel) <> 0 Then
                If btnJoinLeave.Text <> "Leave Channel" Then btnJoinLeave.Text = "Leave Channel"
            ElseIf btnJoinLeave.Text <> "Join Channel" Then
                btnJoinLeave.Text = "Join Channel"
            End If
        End If
        
    End Sub
End Class