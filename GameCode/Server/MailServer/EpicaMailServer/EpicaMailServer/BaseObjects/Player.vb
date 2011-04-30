Public Class Player
    Public PlayerID As Int32

    Public sPlayerName As String
    Public sEmpireName As String
    Public sRaceName As String
    Public sEmailAddress As String
    Public yGender As Byte

    Public sCompareName As String

    Public lPrimaryIdx As Int32         'socket index of the primary server who 'owns' the player object or where the player can be reached...

    Public bIsOnline As Boolean = False 'indicates whether the player is online or not as registered by the primary server...

    Public oGuild As Guild = Nothing

    Public oAliasedAs As Player
    Public oLoggedOnAliases() As Player
    Public lLoggedOnAliasUB As Int32 = -1

    Public bSystemAdmin As Boolean = False

    Public oLocalChatChannel As ChatRoom = Nothing          'the player's local chat channel

    Public oChannels() As ChatRoom = Nothing                'chat rooms the player is currently in

    Public dtCoolDownEnds As Date = Date.MinValue
    Public lCoolDownBy As Int32 = -1

    Private mlPrimaryIndices() As Int32 = Nothing
    Public Function GetPrimaryIdxList() As Int32()
        If mlPrimaryIndices Is Nothing = True Then

            ReDim mlPrimaryIndices(-1)

            If bIsOnline = True Then
                ReDim mlPrimaryIndices(0)
                mlPrimaryIndices(0) = lPrimaryIdx
            End If

            For X As Int32 = 0 To lLoggedOnAliasUB
                If oLoggedOnAliases(X) Is Nothing = False AndAlso oLoggedOnAliases(X).bIsOnline = True Then

                    Dim lPrimary As Int32 = oLoggedOnAliases(X).lPrimaryIdx
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
                ReDim .lRecipients(lLoggedOnAliasUB + 1)
                Dim lIdx As Int32 = -1

                If lPrimaryIndex = Me.lPrimaryIdx Then
                    lIdx += 1
                    .lRecipients(lIdx) = Me.PlayerID
                End If

                For X As Int32 = 0 To lLoggedOnAliasUB
                    If oLoggedOnAliases(X) Is Nothing = False AndAlso oLoggedOnAliases(X).bIsOnline = True AndAlso oLoggedOnAliases(X).lPrimaryIdx = lPrimaryIndex Then
                        lIdx += 1
                        .lRecipients(lIdx) = oLoggedOnAliases(X).PlayerID
                    End If
                Next X
                If lIdx <> .lRecipients.GetUpperBound(0) Then
                    ReDim Preserve .lRecipients(lIdx)
                End If
            End If

            Return .lRecipients
        End With

    End Function

    Public Sub AddAliasLoggedOn(ByRef oNewPlayer As Player)
        Dim lPlayerID As Int32 = oNewPlayer.PlayerID
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lLoggedOnAliasUB
            If oLoggedOnAliases(X) Is Nothing = False Then
                If oLoggedOnAliases(X).PlayerID = lPlayerID Then Return
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            lIdx = lLoggedOnAliasUB + 1
            ReDim Preserve oLoggedOnAliases(lIdx)
            lLoggedOnAliasUB = lIdx
        End If
        oLoggedOnAliases(lIdx) = oNewPlayer

        muPrimaryRecipients = Nothing
        mlPrimaryIndices = Nothing

        'Now, alert the player that the new alias is online
        If lPlayerID <> 1 AndAlso lPlayerID <> 2 AndAlso lPlayerID <> 221 AndAlso lPlayerID <> 131 AndAlso lPlayerID <> 2067 AndAlso lPlayerID <> 7 AndAlso lPlayerID <> 2076 AndAlso lPlayerID <> 1780 Then
            Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, oNewPlayer.sPlayerName & " has logged in as an alias of " & Me.sPlayerName & ".", ChatManager.ChatMessageType.AliasChatMessage, False, "")
            Dim lPrims() As Int32 = GetPrimaryIdxList()
            If lPrims Is Nothing = False Then
                For X As Int32 = 0 To lPrims.GetUpperBound(0)
                    Dim lRecipients() As Int32 = GetRecipientList(lPrims(X))
                    If lRecipients Is Nothing = False Then
                        Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                        goMsgSys.SendToPrimary(lPrims(X), yFinal)
                    End If
                Next X
            End If
        End If
    End Sub
    Public Sub RemoveAliasLoggedOn(ByVal lPlayerID As Int32)
        Dim yMsg() As Byte = Nothing
        For X As Int32 = 0 To lLoggedOnAliasUB
            If oLoggedOnAliases(X) Is Nothing = False Then
                If oLoggedOnAliases(X).PlayerID = lPlayerID Then
                    If lPlayerID <> 1 AndAlso lPlayerID <> 2 AndAlso lPlayerID <> 221 AndAlso lPlayerID <> 131 AndAlso lPlayerID <> 2067 AndAlso lPlayerID <> 7 AndAlso lPlayerID <> 2076 AndAlso lPlayerID <> 1780 Then
                        yMsg = ChatManager.CreateChatMsg(lPlayerID, oLoggedOnAliases(X).sPlayerName & " has logged off of the alias account for " & Me.sPlayerName & ".", ChatManager.ChatMessageType.AliasChatMessage, False, "")
                    End If
                    oLoggedOnAliases(X) = Nothing
                    Exit For
                End If
            End If
        Next X
        mlPrimaryIndices = Nothing
        muPrimaryRecipients = Nothing

        If yMsg Is Nothing = False Then
            Dim lPrims() As Int32 = GetPrimaryIdxList()
            If lPrims Is Nothing = False Then
                For X As Int32 = 0 To lPrims.GetUpperBound(0)
                    Dim lRecipients() As Int32 = GetRecipientList(lPrims(X))
                    If lRecipients Is Nothing = False Then
                        Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                        goMsgSys.SendToPrimary(lPrims(X), yFinal)
                    End If
                Next X
            End If
        End If
    End Sub
    Public Sub AliasedMemberChangePrimaryIndex()
        muPrimaryRecipients = Nothing
        mlPrimaryIndices = Nothing
    End Sub
    Public Sub PlayerChangedPrimaryIndex()
        'reset oGuild
        If oGuild Is Nothing = False Then oGuild.MemberChangePrimaryIndex()
        'if oAliasedAs is nothing = False, reset it
        If oAliasedAs Is Nothing = False Then oAliasedAs.AliasedMemberChangePrimaryIndex()
        'reset me
        mlPrimaryIndices = Nothing
        muPrimaryRecipients = Nothing
        If oLocalChatChannel Is Nothing = False Then oLocalChatChannel.MemberChangePrimaryIndex()
        If oChannels Is Nothing = False Then
            For X As Int32 = 0 To oChannels.GetUpperBound(0)
                If oChannels(X) Is Nothing = False Then oChannels(X).MemberChangePrimaryIndex()
            Next X
        End If
    End Sub
    Public Sub PlayerLoggedOff()
        If oGuild Is Nothing = False Then oGuild.MemberOnlineStatusChange()
        If oAliasedAs Is Nothing = False Then oAliasedAs.RemoveAliasLoggedOn(Me.PlayerID)
        mlPrimaryIndices = Nothing
        muPrimaryRecipients = Nothing
        If oLocalChatChannel Is Nothing = False Then oLocalChatChannel.RemoveMember(Me.PlayerID)
        If oChannels Is Nothing = False Then
            For X As Int32 = 0 To oChannels.GetUpperBound(0)
                If oChannels(X) Is Nothing = False Then oChannels(X).MemberOnlineStatusChange()
            Next X
        End If
    End Sub 
    Public Sub PlayerChangeEnvironments(ByVal lID As Int32, ByVal iTypeID As Int16)
        If oLocalChatChannel Is Nothing = False AndAlso oLocalChatChannel.lExtendedID > 0 AndAlso oLocalChatChannel.iExtendedTypeID > 0 Then oLocalChatChannel.RemoveMember(Me.PlayerID)
        oLocalChatChannel = Nothing

        If lID > -1 AndAlso iTypeID > -1 Then
            'Now, find the environment
            oLocalChatChannel = ChatRoom.GetOrAddChatRoomByEnvir(lID, iTypeID)
            If oLocalChatChannel Is Nothing = False Then
                'oLocalChatChannel.yAttrs = 0
                oLocalChatChannel.AddMember(Me)
            End If
        End If
    End Sub

    Public Sub AddChannel(ByRef oChannel As ChatRoom)
        If oChannels Is Nothing = False Then
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To oChannels.GetUpperBound(0)
                If oChannels(X) Is Nothing = False Then
                    If oChannels(X).sCompareChannelName = oChannel.sCompareChannelName Then Return
                ElseIf lIdx = -1 Then
                    lIdx = X
                End If
            Next X
            If lIdx = -1 Then
                lIdx = oChannels.GetUpperBound(0) + 1
                ReDim Preserve oChannels(lIdx)
            End If
            oChannels(lIdx) = oChannel
        Else
            ReDim oChannels(0)
            oChannels(0) = oChannel
        End If
    End Sub
    Public Sub RemoveChannel(ByRef oChannel As ChatRoom)
        If oChannels Is Nothing Then Return
        For X As Int32 = 0 To oChannels.GetUpperBound(0)
            If oChannels(X) Is Nothing = False AndAlso oChannels(X).sCompareChannelName = oChannel.sCompareChannelName Then
                oChannels(X) = Nothing
            End If
        Next X
    End Sub
End Class

