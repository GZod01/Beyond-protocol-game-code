Option Strict On

''' <summary>
''' Conquest Arena - control points on the map to get to the enemy's command center and then blow it up to win the round
''' </summary>
''' <remarks></remarks>
Public Enum eyCFlags As Byte
    ControlPointAbilities = 0       'indicates that the control points on the map give abilities/powers or if they are disabled
    RandomizeAbilities = 1          'indicates whether to use the map's default abilities or to randomize them
End Enum
Public Class C_Arena
    Inherits Arena

#Region "  Configuration  "
    Public lRounds As Int32         'number of rounds in this match
    Public yCFlags As Byte         'uses eyCFlags enum
#End Region

    Public Overrides Sub AppendGameModeSpecificData(ByRef yData() As Byte, ByRef lPos As Integer)
        '
    End Sub

    Public Overrides Function GetGameModeSpecificDataLen() As Integer
        Return 0
    End Function
End Class