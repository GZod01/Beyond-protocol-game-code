Option Strict On

Public Enum eyCTFFlags As Byte
    ResetOnCapture = 1                      'indicates that the game resets and units are placed back at "base" when a flag is captured
    OpenSeason = 2                          'if set, an enemy flag can be captured even if the player's flag is not at their base
    AutoReturnFlag = 4                      'if set, when the flag is picked up by the owner, it is automatically returned to base, otherwise it must be manually brought back
    GroundUnitsCanGetFlag = 8               'indicates whether ground units can get flag
    AirUnitsCanGetFlag = 16                 'indicates whether flying units can get flag
    Engineers = 32                          'indicates whether engineers can get flag
End Enum
''' <summary>
''' Capture the Flag Arena - capture enemy flags to score points
''' </summary>
''' <remarks></remarks>
Public Class CTF_Arena
    Inherits Arena

#Region "  Configuration  "
    Public lGoal As Int32                   'how many scores are required by a side in order to win
    Public yCTFFlags As Byte                'enumeration of eyCTFFlags
    Public lResetDelayTime As Int32         'in seconds, how long before the game resets and restarts when a flag is captured - only used if the ResetOnCapture flag is set

    Public yFlagCapturerMaxSpeed As Byte    'indicates max speed of the flag capturer
#End Region

    Public Overrides Sub AppendGameModeSpecificData(ByRef yData() As Byte, ByRef lPos As Integer)
        '
    End Sub

    Public Overrides Function GetGameModeSpecificDataLen() As Integer
        Return 0
    End Function
End Class
