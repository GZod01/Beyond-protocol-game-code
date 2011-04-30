Option Strict On

Public Enum eyArenaSidePlayerFlag As Byte
    PlayerReady = 1
End Enum
Public Enum eyArenaPlayerActionCode As Byte
    PlayerLeaveArena = 0
    PlayerRequestJoin = 1
    PlayerJoinsArena = 2
    PlayerChangesSides = 3
    SetPlayerNotReady = 254
    SetPlayerReady = 255
End Enum
Public Class ArenaSidePlayer
    Public oSide As ArenaSide

    Public lPlayerID As Int32 

    Public lCapturePoints As Int32  'number of points from capturing whether in a CTF game or CP game

    Public yFlags As Byte           'uses eyArenaSidePlayerFlag

    Public oContestants() As ArenaContestant
    Public lContestantUB As Int32 = -1 

    Public Function AddContestant(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16) As ArenaContestant
        lContestantUB += 1
        Dim lIdx As Int32 = lContestantUB
        ReDim Preserve oContestants(lContestantUB)
        oContestants(lIdx) = New ArenaContestant(lEntityID, iEntityTypeID)
        Return oContestants(lIdx)
    End Function
End Class
