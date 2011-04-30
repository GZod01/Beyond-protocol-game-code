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

    Public Overrides Sub AppendGameModeSpecificData(ByRef yData() As Byte, ByRef lPos As Integer)
        System.BitConverter.GetBytes(lRounds).CopyTo(yData, lPos) : lPos += 4
        System.BitConverter.GetBytes(lKillGoal).CopyTo(yData, lPos) : lPos += 4
        System.BitConverter.GetBytes(lRoundTimeLimit).CopyTo(yData, lPos) : lPos += 4
    End Sub

    Public Overrides Function GetGameModeSpecificDataLen() As Integer
        Return 12
    End Function
End Class
