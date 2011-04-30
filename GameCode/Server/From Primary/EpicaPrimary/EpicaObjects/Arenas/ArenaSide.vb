Option Strict On

Public Class ArenaSide
    Public lSideID As Int32         'a value indicating the side id

    Public lSideScore As Int32      'score for the side... used in overall "which side won" comparisons. For CTF games, this is flag captures. For DM, this is kills. For CP this is matches.

    Public oSidePlayers() As ArenaSidePlayer
End Class
