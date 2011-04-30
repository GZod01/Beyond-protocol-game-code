Public Class MailQueue

    Private mlItemIdx() As Int32
    Public mlItemUB As Int32 = -1

    Private mlLastIndex As Int32 = 0

    Private mlAddCnt As Int32 = 0           'number of adds in progress

    Public Sub AddNew(ByVal lMailIdx As Int32)
        mlAddCnt += 1

        Dim lIdx As Int32 = -1

        For X As Int32 = 0 To mlItemUB
            If mlItemIdx(X) = -1 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            ReDim Preserve mlItemIdx(mlItemUB + 1)
            mlItemUB += 1
            lIdx = mlItemUB
        End If
        mlItemIdx(lIdx) = lMailIdx

        mlAddCnt -= 1
    End Sub

    Public Function GetNextItem() As Int32
        If mlAddCnt <> 0 Then
            LogEvent("Waited a cycle, lCnt: " & mlAddCnt)
            Return -1
        End If

        If mlLastIndex > mlItemUB Then mlLastIndex = 0

        For X As Int32 = mlLastIndex To mlItemUB
            If mlItemIdx(X) <> -1 Then
                mlLastIndex = X

                Dim lTempRet As Int32 = mlItemIdx(X)
                mlItemIdx(X) = -1
                Return lTempRet
            End If
        Next X
        mlLastIndex = 0
        Return -1
    End Function

End Class