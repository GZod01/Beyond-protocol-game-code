Option Strict On

''' <summary>
''' Control Point Arena - Control points on the map for an amount of time to win the round
''' </summary>
''' <remarks></remarks>
Public Enum eyCPFlags As Byte
    ControlPointAbilities = 0       'indicates that the control points on the map give abilities/powers or if they are disabled
    RandomizeAbilities = 1          'indicates whether to use the map's default abilities or to randomize them
End Enum
Public Class CP_Arena
    Inherits Arena

#Region "  Configuration  "
    Public lCPGoal As Int32         'number of control points to get access to
    Public lHoldTime As Int32       'number of seconds to hold the control points before the side wins the round
    Public lRounds As Int32         'number of rounds in this match

    Public yCPFlags As Byte         'uses eyCPFlags enum
#End Region

    Public Overrides Sub AppendGameModeSpecificData(ByRef yData() As Byte, ByRef lPos As Integer)

    End Sub

    Public Overrides Function GetGameModeSpecificDataLen() As Integer
        Return 0
    End Function
End Class
