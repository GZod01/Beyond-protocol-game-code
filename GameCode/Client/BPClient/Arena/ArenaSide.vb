Option Strict On

Public Class ArenaSide
    Public lSideID As Int32         'a value indicating the side id

    Public lSideScore As Int32      'score for the side... used in overall "which side won" comparisons. For CTF games, this is flag captures. For DM, this is kills. For CP this is matches.

    Public oSidePlayers() As ArenaSidePlayer

    Public Sub AddArenaSidePlayer(ByVal lPlayerID As Int32)
        Dim lIdx As Int32 = -1
        If oSidePlayers Is Nothing = False Then
            For X As Int32 = 0 To oSidePlayers.GetUpperBound(0)
                If oSidePlayers(X) Is Nothing Then
                    If lIdx = -1 Then lIdx = X
                ElseIf oSidePlayers(X).lPlayerID = lPlayerID Then
                    Return
                End If
            Next X
        End If

        If lIdx = -1 Then
            If oSidePlayers Is Nothing Then
                ReDim oSidePlayers(0)
                lIdx = 0
            Else
                lIdx = oSidePlayers.GetUpperBound(0)
                lIdx += 1
                ReDim Preserve oSidePlayers(lIdx)
            End If
        End If

        oSidePlayers(lIdx) = New ArenaSidePlayer
        With oSidePlayers(lIdx)
            .lPlayerID = lPlayerID
            .lContestantUB = -1
            .oSide = Me
            .yFlags = 0
        End With
    End Sub
End Class
