Option Strict On

''' <summary>
''' Contested Scenario Arena - missions to be ran - quests - but places teams against each other
''' </summary>
''' <remarks></remarks>
Public Class CS_Arena
    Inherits Arena

    Public Overrides Sub AppendGameModeSpecificData(ByRef yData() As Byte, ByRef lPos As Integer)
        '
    End Sub

    Public Overrides Function GetGameModeSpecificDataLen() As Integer
        Return 0
    End Function
End Class
