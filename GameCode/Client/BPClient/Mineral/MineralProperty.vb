Public Class MineralProperty
    Inherits Base_GUID

    Public MineralPropertyName As String

    Public yKnowledgeLevel As Byte = 0

    Public Shared Function MineralPropertyKnowledgeLevel(ByVal lPropertyID As Int32) As Byte
        For X As Int32 = 0 To glMineralPropertyUB
            If glMineralPropertyIdx(X) = lPropertyID Then Return goMineralProperty(X).yKnowledgeLevel
        Next
        Return 0
    End Function
End Class

Public Enum eMinPropKnowledgeLevel As Byte
    eExistence = 0              'Min Name and ACTUAL
    eAdvancedKnowledge = 1      'min property KNOWN value
    eExpertKnowledge = 2        'alloy builder property
End Enum