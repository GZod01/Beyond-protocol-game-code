Option Strict On

Public Enum eyChatRoomAttr As Byte
    PublicChannel = 1
    UserIsAdmin = 2         'valid on a user-to-user basis
    AllAreAdmin = 4
    PasswordProtected = 8
    PlayerIsPermitted = 16  'valid on a user-to-user basis
    PlayerInChannel = 32    'valid on a user-to-user basis
End Enum

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

Public Class ChatRoom
    Public sChannelName As String
    Public sCompareChannelName As String
    Public sChannelPassword As String

    'For local channels - ExtendedID is the ObjectID of the environment, ExtendedTypeID is the ObjTypeID of the environment
    Public lExtendedID As Int32 = -1
    Public iExtendedTypeID As Int16 = -1

    Public oMembers() As Player
    Private mbHasAdminRights() As Boolean
    Public lMemberUB As Int32 = -1

    Public lInvitedId() As Int32

    Public yAttrs As eyChatRoomAttr

    Public lPlayerCount As Int32 = 0

    Private mlPrimaryIndices() As Int32 = Nothing
    Public Function GetPrimaryIdxList() As Int32()
        If mlPrimaryIndices Is Nothing = True Then

            ReDim mlPrimaryIndices(-1)

            For X As Int32 = 0 To lMemberUB
                If oMembers(X) Is Nothing = False AndAlso oMembers(X).bIsOnline = True Then

                    Dim lPrimary As Int32 = oMembers(X).lPrimaryIdx
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To mlPrimaryIndices.GetUpperBound(0)
                        If mlPrimaryIndices(Y) = lPrimary Then
                            bFound = True
                            Exit For
                        End If
                    Next Y
                    If bFound = False Then
                        Dim lUB As Int32 = mlPrimaryIndices.GetUpperBound(0)
                        lUB += 1
                        ReDim Preserve mlPrimaryIndices(lUB)
                        mlPrimaryIndices(lUB) = lPrimary
                    End If
                End If
            Next X

            ReDim muPrimaryRecipients(mlPrimaryIndices.GetUpperBound(0))
            For X As Int32 = 0 To mlPrimaryIndices.GetUpperBound(0)
                muPrimaryRecipients(X).lPrimaryIdx = mlPrimaryIndices(X)
                muPrimaryRecipients(X).lRecipients = Nothing
            Next X
        End If

        Return mlPrimaryIndices
    End Function

    Private muPrimaryRecipients() As PrimaryRecipients
    Public Function GetRecipientList(ByVal lPrimaryIndex As Int32) As Int32()
        If muPrimaryRecipients Is Nothing = True Then
            GetPrimaryIdxList()
            If muPrimaryRecipients Is Nothing Then Return Nothing
        End If

        Dim lPRIdx As Int32 = -1
        For X As Int32 = 0 To muPrimaryRecipients.GetUpperBound(0)
            If muPrimaryRecipients(X).lPrimaryIdx = lPrimaryIndex Then
                lPRIdx = X
                Exit For
            End If
        Next X
        If lPRIdx = -1 Then Return Nothing

        With muPrimaryRecipients(lPRIdx)
            If .lRecipients Is Nothing = True Then
                ReDim .lRecipients(lMemberUB)
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To lMemberUB
                    If oMembers(X) Is Nothing = False AndAlso oMembers(X).bIsOnline = True AndAlso oMembers(X).lPrimaryIdx = lPrimaryIndex Then
                        lIdx += 1
                        .lRecipients(lIdx) = oMembers(X).PlayerID
                    End If
                Next X
                If lIdx <> lMemberUB Then
                    ReDim Preserve .lRecipients(lIdx)
                End If
            End If

            Return .lRecipients
        End With

    End Function

    Public Function PlayerInChatRoom(ByVal lPlayerID As Int32) As Boolean
        If lPlayerID = 1 Then Return True
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).PlayerID = lPlayerID Then Return True
        Next X
        Return False
    End Function
    Public Function PlayerInvited(ByVal lPlayerID As Int32) As Boolean
        If lInvitedId Is Nothing Then Return False
        For X As Int32 = 0 To lInvitedId.GetUpperBound(0)
            If lInvitedId(X) = lPlayerID Then Return True
        Next X
        Return False
    End Function
    Public Function AdjustAttrForPlayer(ByVal lPlayerID As Int32, ByVal yAttr As Byte) As Byte
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).PlayerID = lPlayerID Then
                Dim yNewAttr As Byte = yAttr Or eyChatRoomAttr.PlayerInChannel
                If (mbHasAdminRights Is Nothing = False AndAlso mbHasAdminRights(X) = True) OrElse lPlayerID = 1 Then
                    yNewAttr = yNewAttr Or eyChatRoomAttr.UserIsAdmin
                End If
                Return yNewAttr
            End If
        Next X
        Return yAttr
    End Function
    Public Function PlayerIsAdmin(ByVal lPlayerID As Int32) As Boolean
        If lPlayerID = 1 Then Return True
        If (Me.yAttrs And eyChatRoomAttr.AllAreAdmin) <> 0 Then Return True
        If mbHasAdminRights Is Nothing Then Return False
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).PlayerID = lPlayerID Then Return mbHasAdminRights(X)
        Next X
        Return False
    End Function
    Public Function ToggleAdmin(ByVal lPlayerID As Int32, ByRef bResult As Boolean) As String
        If mbHasAdminRights Is Nothing Then ReDim mbHasAdminRights(lMemberUB)
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).PlayerID = lPlayerID Then
                mbHasAdminRights(X) = Not mbHasAdminRights(X)
                bResult = mbHasAdminRights(X)
                Return oMembers(X).sPlayerName
            End If
        Next X
        Return ""
    End Function

    Public Sub AddMember(ByRef oMember As Player)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False Then
                If oMembers(X).PlayerID = oMember.PlayerID Then
                    Return
                End If
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            lIdx = lMemberUB + 1
            ReDim Preserve oMembers(lIdx)
            ReDim Preserve mbHasAdminRights(lIdx)
            lMemberUB += 1
        End If
        oMembers(lIdx) = oMember
        mbHasAdminRights(lIdx) = False

        mlPrimaryIndices = Nothing
        muPrimaryRecipients = Nothing

        lPlayerCount += 1
    End Sub
    Public Sub RemoveMember(ByVal lMemberID As Int32)
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).PlayerID = lMemberID Then
                oMembers(X) = Nothing
                Exit For
            End If
        Next X

        mlPrimaryIndices = Nothing
        muPrimaryRecipients = Nothing

        lPlayerCount -= 1
    End Sub
    Public Sub MemberChangePrimaryIndex()
        muPrimaryRecipients = Nothing
        mlPrimaryIndices = Nothing
    End Sub
    Public Sub MemberOnlineStatusChange()
        muPrimaryRecipients = Nothing
        mlPrimaryIndices = Nothing
    End Sub
    Public Function GetChatRoomDetailsMsg(ByVal lRequestorID As Int32) As Byte()
        lPlayerCount = 0
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False Then
                lPlayerCount += 1
            End If
        Next X

        Dim lPCnt As Int32 = lPlayerCount
        Dim yResp(70 + (lPCnt * 4)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelDetails).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(lRequestorID).CopyTo(yResp, lPos) : lPos += 4
        StringToBytes(sChannelName).CopyTo(yResp, lPos) : lPos += 30
        If sChannelPassword Is Nothing = False Then StringToBytes(sChannelPassword).CopyTo(yResp, lPos)
        lPos += 30
        yResp(lPos) = yAttrs : lPos += 1
        System.BitConverter.GetBytes(lPCnt).CopyTo(yResp, lPos) : lPos += 4
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False Then
                lPCnt -= 1
                If lPCnt < 0 Then Exit For

                Dim lPlayerID As Int32 = oMembers(X).PlayerID
                If mbHasAdminRights Is Nothing = False AndAlso mbHasAdminRights(X) = True Then lPlayerID = -lPlayerID
                System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
            End If
        Next X

        Return yResp
    End Function

    Public Sub AddInvited(ByVal lPlayerID As Int32)
        If lInvitedId Is Nothing = False Then
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To lInvitedId.GetUpperBound(0)
                If lInvitedId(X) = lPlayerID Then
                    Return
                ElseIf lInvitedId(X) = -1 AndAlso lIdx = -1 Then
                    lIdx = X
                End If
            Next X
            If lIdx = -1 Then
                ReDim Preserve lInvitedId(lInvitedId.GetUpperBound(0) + 1)
                lIdx = lInvitedId.GetUpperBound(0)
            End If
            lInvitedId(lIdx) = lPlayerID
        Else
            ReDim lInvitedId(0)
            lInvitedId(0) = lPlayerID
        End If
    End Sub
    Public Sub RemoveInvited(ByVal lPlayerID As Int32)
        If lInvitedId Is Nothing Then Return
        For X As Int32 = 0 To lInvitedId.GetUpperBound(0)
            If lInvitedId(X) = lPlayerID Then
                lInvitedId(X) = -1
            End If
        Next X
    End Sub

    Public Sub SendToRoomMembers(ByVal lFromID As Int32, ByVal sMsg As String, ByVal bLogEvent As Boolean)
        Dim yMsg() As Byte = ChatManager.CreateChatMsg(lFromID, sMsg, ChatManager.ChatMessageType.ChannelMessage, bLogEvent, sChannelName)
        Dim lPrims() As Int32 = GetPrimaryIdxList()
        If lPrims Is Nothing Then Return
        For X As Int32 = 0 To lPrims.GetUpperBound(0)
            Dim lRecs() As Int32 = GetRecipientList(lPrims(X))
            If lRecs Is Nothing = False Then
                Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecs)
                goMsgSys.SendToPrimary(lPrims(X), yFinal)
            End If
        Next X
    End Sub


#Region "  Global / Singletons  "
    Private Shared moChatRooms() As ChatRoom
    Private Shared mlChatRoomUB As Int32 = -1

    Public Shared Function GetChatRoomByName(ByVal sName As String) As ChatRoom
        If moChatRooms Is Nothing = True Then Return Nothing
        Dim lCurUB As Int32 = Math.Min(mlChatRoomUB, moChatRooms.GetUpperBound(0))

        Dim sSearchName As String = sName.ToUpper
        For X As Int32 = 0 To lCurUB
            If moChatRooms(X) Is Nothing = False Then
                If moChatRooms(X).sCompareChannelName = sSearchName Then
                    Return moChatRooms(X)
                End If
            End If
        Next X
        Return Nothing
    End Function
    Public Shared Function GetOrAddChatRoomByName(ByVal sRoomName As String) As ChatRoom
        Dim sSearch As String = sRoomName.ToUpper
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlChatRoomUB
            If moChatRooms(X) Is Nothing = False Then
                If moChatRooms(X).sCompareChannelName = sSearch Then Return moChatRooms(X)
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            lIdx = mlChatRoomUB + 1
            ReDim Preserve moChatRooms(lIdx)
            mlChatRoomUB += 1
        End If
        moChatRooms(lIdx) = New ChatRoom
        With moChatRooms(lIdx)
            .sChannelName = sRoomName
            .sCompareChannelName = sRoomName.ToUpper
            .lExtendedID = -1
            .iExtendedTypeID = -1
        End With
        Return moChatRooms(lIdx)
    End Function
    Public Shared Function GetOrAddChatRoomByEnvir(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As ChatRoom
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlChatRoomUB
            If moChatRooms(X) Is Nothing = False Then
                If moChatRooms(X).lExtendedID = lEnvirID AndAlso moChatRooms(X).iExtendedTypeID = iEnvirTypeID Then Return moChatRooms(X)
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            lIdx = mlChatRoomUB + 1
            ReDim Preserve moChatRooms(lIdx)
            mlChatRoomUB += 1
        End If
        moChatRooms(lIdx) = New ChatRoom
        With moChatRooms(lIdx)
            .sChannelName = "_CreatedEnvirRoom" & lEnvirID & "," & iEnvirTypeID
            .sCompareChannelName = .sChannelName.ToUpper
            .lExtendedID = lEnvirID
            .iExtendedTypeID = iEnvirTypeID
        End With
        Return moChatRooms(lIdx)
    End Function

    Public Shared Sub CreateStandardRooms()
        Dim oRoom As ChatRoom = GetOrAddChatRoomByName("General")
        If oRoom Is Nothing = False Then
            oRoom.yAttrs = eyChatRoomAttr.PublicChannel
            oRoom.sChannelPassword = ""
            oRoom.lExtendedID = -1
            oRoom.iExtendedTypeID = -1
        End If
        'oRoom = GetOrAddChatRoomByName("FAQ_CSR")
        'If oRoom Is Nothing = False Then
        '    oRoom.yAttrs = eyChatRoomAttr.PublicChannel
        '    oRoom.sChannelPassword = ""
        '    oRoom.lExtendedID = -1
        '    oRoom.iExtendedTypeID = -1
        'End If
        oRoom = GetOrAddChatRoomByName("FR")
        If oRoom Is Nothing = False Then
            oRoom.yAttrs = eyChatRoomAttr.PublicChannel
            oRoom.sChannelPassword = ""
            oRoom.lExtendedID = -1
            oRoom.iExtendedTypeID = -1
        End If
        oRoom = GetOrAddChatRoomByName("Emperor")
        If oRoom Is Nothing = False Then
            oRoom.yAttrs = 0
            oRoom.sChannelPassword = ""
            oRoom.lExtendedID = -1
            oRoom.iExtendedTypeID = -1
        End If
    End Sub

    Public Shared Function HandleRequestChannelList(ByVal yData() As Byte) As Byte()
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim yEmperor As Byte = yData(6)

        Dim oPlayer As Player = GetPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return Nothing

        Dim lCnt As Int32 = 0

        Dim lCurUB As Int32 = -1
        If moChatRooms Is Nothing = False Then lCurUB = Math.Min(mlChatRoomUB, moChatRooms.GetUpperBound(0))
        Dim bIncluded(lCurUB) As Boolean
        For X As Int32 = 0 To lCurUB
            If moChatRooms(X).lExtendedID = -1 AndAlso moChatRooms(X).iExtendedTypeID = -1 AndAlso ((moChatRooms(X).yAttrs And eyChatRoomAttr.PublicChannel) <> 0 OrElse moChatRooms(X).PlayerInChatRoom(lPlayerID) = True OrElse moChatRooms(X).PlayerInvited(lPlayerID)) Then
                lCnt += 1
                bIncluded(X) = True
            ElseIf yEmperor <> 0 AndAlso moChatRooms(X).sCompareChannelName = "EMPEROR" Then
                lCnt += 1
                bIncluded(X) = True
            Else : bIncluded(X) = False
            End If
        Next X

        Dim yResp(9 + (lCnt * 35)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelList).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lPos) : lPos += 4

        For X As Int32 = 0 To lCurUB

            If bIncluded(X) = True Then
                lCnt -= 1
                If lCnt < 0 Then Exit For
                With moChatRooms(X)
                    StringToBytes(.sChannelName).CopyTo(yResp, lPos) : lPos += 30
                    'System.BitConverter.GetBytes(.lID).CopyTo(yResp, lPos) : lPos += 4
                    Dim yTmpAttr As Byte = .yAttrs
                    If .sChannelPassword <> "" Then
                        yTmpAttr = yTmpAttr Or eyChatRoomAttr.PasswordProtected
                        If .yAttrs <> yTmpAttr Then .yAttrs = CType(yTmpAttr, eyChatRoomAttr)
                    End If

                    yTmpAttr = .AdjustAttrForPlayer(lPlayerID, yTmpAttr)
                    yResp(lPos) = yTmpAttr : lPos += 1

                    System.BitConverter.GetBytes(.lPlayerCount).CopyTo(yResp, lPos) : lPos += 4
                End With
            End If

        Next X

        Return yResp
    End Function
#End Region

End Class
