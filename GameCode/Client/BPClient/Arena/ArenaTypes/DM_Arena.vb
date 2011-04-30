Option Strict On

''' <summary>
''' Deathmatch Arena - kill enemy units to gain points
''' </summary>
''' <remarks></remarks>
Public Class DM_Arena
    Inherits Arena

#Region "  Configuration  "
    Public lRounds As Int32             'number of rounds played

    Public lKillGoal As Int32           'number of kills needed to win the round. 0 indicates play until times up
    Public lRoundTimeLimit As Int32     'number of minutes in a round
#End Region  
End Class
