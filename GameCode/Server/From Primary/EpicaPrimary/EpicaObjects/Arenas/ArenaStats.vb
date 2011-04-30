Option Strict On

Public Class ArenaStats

    'All of these stats are cumulative. Adjust them using +=. That way, you can do averages, for example... Average Matches Won = lMatchesWon(lArenaType) / lMatchesPlayed(lArenaType)
    '  Or, average units involved in matches = lUnitsInvolvedInMatches(lArenaType) / lMatchesPlayer(lArenaType)
    Public lMatchesPlayed() As Int32                'by ArenaType
    Public lMatchesWon() As Int32                   'by ArenaType
    Public lUnitsInvolvedInMatches() As Int32       'by ArenaType
    Public lSidesInvolvedInMatches() As Int32       'by ArenaType
    Public lDifficultyRating() As Int32             'by ArenaType
    Public lUnitsKilled() As Int32                  'by ArenaType
    Public lUnitsLost() As Int32                    'by ArenaType

End Class
