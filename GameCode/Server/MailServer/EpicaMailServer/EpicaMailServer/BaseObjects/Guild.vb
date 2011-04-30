Public Class Guild
    Public lGuildID As Int32
    Public oMembers() As Player
    Public lMemberUB As Int32 = -1

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

    Public Sub AddMember(ByRef oPlayer As Player)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False Then
                If oMembers(X).PlayerID = oPlayer.PlayerID Then
                    Return
                End If
            ElseIf lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            lIdx = lMemberUB + 1
            ReDim Preserve oMembers(lIdx)
            lMemberUB = lIdx
        End If
        oMembers(lIdx) = oPlayer
        mlPrimaryIndices = Nothing
        muPrimaryRecipients = Nothing
    End Sub
    Public Sub RemoveMember(ByVal lMemberID As Int32)
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False Then
                If oMembers(X).PlayerID = lMemberID Then
                    oMembers(X) = Nothing
                    Exit For
                End If
            End If
        Next X
        mlPrimaryIndices = Nothing
        muPrimaryRecipients = Nothing
    End Sub
    Public Sub MemberChangePrimaryIndex()
        muPrimaryRecipients = Nothing
        mlPrimaryIndices = Nothing
    End Sub
    Public Sub MemberOnlineStatusChange()
        muPrimaryRecipients = Nothing
        mlPrimaryIndices = Nothing
    End Sub

    Private Shared moGuilds() As Guild
    Private Shared mlGuildUB As Int32 = -1
    Private Shared mlGuildIdx() As Int32

    Public Shared Function GetOrAddGuild(ByVal lID As Int32) As Guild
        If lID < 1 Then Return Nothing
        If mlGuildIdx Is Nothing Then ReDim mlGuildIdx(-1)
        Dim lCurUB As Int32 = Math.Min(mlGuildUB, mlGuildIdx.GetUpperBound(0))
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCurUB
            If mlGuildIdx(X) = lID Then
                Return moGuilds(X)
            ElseIf mlGuildIdx(X) = -1 AndAlso lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            mlGuildUB += 1
            ReDim Preserve mlGuildIdx(mlGuildUB)
            ReDim Preserve moGuilds(mlGuildUB)
            lIdx = mlGuildUB
        End If
        moGuilds(lIdx) = New Guild()
        moGuilds(lIdx).lGuildID = lID
        mlGuildIdx(lIdx) = lID

        Return moGuilds(lIdx)
    End Function
    Public Shared Function GetGuild(ByVal lID As Int32) As Guild
        If lID < 1 Then Return Nothing
        If mlGuildIdx Is Nothing Then Return Nothing
        Dim lCurUB As Int32 = Math.Min(mlGuildUB, mlGuildIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mlGuildIdx(X) = lID Then
                Return moGuilds(X)
            End If
        Next X
        Return Nothing
    End Function

End Class

