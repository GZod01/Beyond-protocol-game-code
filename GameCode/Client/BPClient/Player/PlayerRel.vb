Public Class PlayerRel
    Public lPlayerRegards As Int32
    Public lThisPlayer As Int32
    Public WithThisScore As Byte

    Public TargetScore As Byte
    Public lNextUpdateCycle As Int32 = -1

    'Only one player intel because this assumes one player in this relationship is the current player
    Private mlIntelIdx As Int32 = -1
    Public ReadOnly Property oPlayerIntel() As PlayerIntel
        Get
            If mlIntelIdx = -1 Then
                If lPlayerRegards = glPlayerID Then
                    For X As Int32 = 0 To glPlayerIntelUB
                        If glPlayerIntelIdx(X) = lThisPlayer Then
                            mlIntelIdx = X
                            Exit For
                        End If
                    Next X
                Else
                    For X As Int32 = 0 To glPlayerIntelUB
                        If glPlayerIntelIdx(X) = lPlayerRegards Then
                            mlIntelIdx = X
                            Exit For
                        End If
                    Next X
                End If

                If mlIntelIdx <> -1 Then Return goPlayerIntel(mlIntelIdx)
            Else : Return goPlayerIntel(mlIntelIdx)
            End If
            Return Nothing
        End Get
    End Property
End Class
