Public Class frmChannelConfig

    Private msChannelName As String
    Private msOriginalPassword As String = ""
    Private mbOriginalPublic As Boolean = False
    Private mlChannelID As Int32 = -1

    Private Class MemberItem
        Public lID As Int32
        Public bAdmin As Boolean
    End Class
    Private moMembers() As MemberItem

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If msChannelName Is Nothing OrElse msChannelName = "" Then
            '"Create Channel: /create <channelname>")
            msChannelName = txtChannelName.Text
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
            If chkPublic.Checked = True Then yPublic = 1
            If txtPassword.Text Is Nothing = False AndAlso txtPassword.Text <> "" Then sPassword = txtPassword.Text
            SubmitConfigMessage(eyChatRoomCommandType.AddNewChatRoom Or eyChatRoomCommandType.SetChannelPublic Or eyChatRoomCommandType.SetChannelPassword, yPublic, -1, sPassword, "")
        Else
            Dim yType As eyChatRoomCommandType = 0
            If txtPassword.Text Is Nothing = False Then
                If txtPassword.Text <> "" Then
                    If msOriginalPassword <> txtPassword.Text Then
                        yType = eyChatRoomCommandType.SetChannelPassword
                    End If
                End If
            End If
            If mbOriginalPublic <> chkPublic.Checked Then
                mbOriginalPublic = chkPublic.Checked
                yType = yType Or eyChatRoomCommandType.SetChannelPublic
            End If
            If yType <> 0 Then
                Dim yPublic As Byte = 0
                If chkPublic.Checked = True Then yPublic = 1
                SubmitConfigMessage(yType, yPublic, -1, txtPassword.Text, "")
            End If
        End If
    End Sub
    Private Sub btnKick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKick.Click
        If lstMembers Is Nothing = False AndAlso lstMembers.SelectedIndex > -1 Then
            Dim lIdx As Int32 = lstMembers.SelectedIndex
            If moMembers Is Nothing = False AndAlso lIdx > -1 AndAlso lIdx < moMembers.Length Then
                Dim oMember As MemberItem = moMembers(lIdx) 'CType(lstMembers.SelectedItem, MemberItem)
                If oMember Is Nothing = False Then SubmitConfigMessage(eyChatRoomCommandType.KickPlayer, 0, oMember.lID, "", "")
            End If
        End If
    End Sub
    Private Sub btnAdmin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdmin.Click
        If lstMembers Is Nothing = False AndAlso lstMembers.SelectedIndex > -1 Then
            Dim lIdx As Int32 = lstMembers.SelectedIndex
            If moMembers Is Nothing = False AndAlso lIdx > -1 AndAlso lIdx < moMembers.Length Then
                Dim oMember As MemberItem = moMembers(lIdx) 'CType(lstMembers.SelectedItem, MemberItem)
                If oMember Is Nothing = False Then SubmitConfigMessage(eyChatRoomCommandType.ToggleAdminRights, 0, oMember.lID, "", "")
            End If
        End If
    End Sub
    Private Sub btnInvite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInvite.Click
        If txtInvite.Text Is Nothing = False AndAlso txtInvite.Text.Trim <> "" Then
            '/invite <playername>, <channelname>")
            'Prepare our message
            'Dim sText As String = "/invite " & txtInvite.Caption & ", " & msChannelName
            'Dim yMsg(sText.Length + 1) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
            'System.Text.ASCIIEncoding.ASCII.GetBytes(sText).CopyTo(yMsg, 2)
            'MyBase.moUILib.SendMsgToPrimary(yMsg)
            SubmitConfigMessage(eyChatRoomCommandType.InvitePlayer, 0, 0, "", txtInvite.Text)
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
        If gfrmChat Is Nothing = False Then gfrmChat.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub SubmitConfigMessage(ByVal yType As eyChatRoomCommandType, ByVal yPublicFlag As Byte, ByVal lTargetID As Int32, ByVal sPassword As String, ByVal sPlayerToInvite As String)
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

        If gfrmChat Is Nothing = False Then gfrmChat.SendMsgToPrimary(yMsg)
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
        Dim oMembers(lMemberCnt - 1) As MemberItem
        For X As Int32 = 0 To lMemberCnt - 1
            oMembers(X) = New MemberItem
            With oMembers(X)
                .lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .bAdmin = .lID < 0
                .lID = Math.Abs(.lID)
            End With
        Next X

        'Now, what to do with it all....
        moMembers = oMembers

        Me.Invoke(New DoUpdate(AddressOf ActuallyDoUpdate))
    End Sub

    Private Delegate Sub DoUpdate()
    Private Sub ActuallyDoUpdate()
        txtChannelName.Text = msChannelName
        txtPassword.Text = msOriginalPassword
        chkPublic.Checked = mbOriginalPublic

        lstMembers.Items.Clear()
        If moMembers Is Nothing = False Then
            'lstMembers.DisplayMember = "DisplayVal"
            'lstMembers.Sorted = True
            For X As Int32 = 0 To moMembers.GetUpperBound(0)
                If moMembers(X) Is Nothing = False Then
                    lstMembers.Items.Add(GetCacheObjectValue(moMembers(X).lID, ObjectType.ePlayer))
                End If
            Next X
        End If
    End Sub
End Class