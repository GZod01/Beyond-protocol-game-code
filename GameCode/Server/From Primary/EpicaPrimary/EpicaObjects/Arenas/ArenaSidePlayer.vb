Option Strict On

Public Enum eyArenaSidePlayerFlag As Byte
    PlayerReady = 1
End Enum
Public Class ArenaSidePlayer
    Public oSide As ArenaSide

    Public lPlayerID As Int32
    Public oPlayer As Player        'pointer

    Public lCapturePoints As Int32  'number of points from capturing whether in a CTF game or CP game

    Public yFlags As Byte           'uses eyArenaSidePlayerFlag

    Public oContestants() As ArenaContestant
    Public lContestantUB As Int32 = -1

    Public oStats As ArenaStats

    Public Function AddContestant(ByRef oEntity As Epica_Entity) As ArenaContestant
        lContestantUB += 1
        Dim lIdx As Int32 = lContestantUB
        ReDim Preserve oContestants(lContestantUB)
        oContestants(lIdx) = New ArenaContestant(oEntity)
        Return oContestants(lIdx)
    End Function
End Class
