Option Strict On

''' <summary>
''' Scenario Arena - missions to be ran - quests
''' </summary>
''' <remarks></remarks>
Public Class S_Arena
    Inherits Arena

    Public Overrides Sub AppendGameModeSpecificData(ByRef yData() As Byte, ByRef lPos As Integer)
        '
    End Sub

    Public Overrides Function GetGameModeSpecificDataLen() As Integer
        Return 0
    End Function
End Class
