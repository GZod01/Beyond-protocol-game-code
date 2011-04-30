Option Strict On

Public Class PlayerCommFolder
    Public Enum ePCF_ID_HARDCODES As Integer
        eInbox_PCF = -2
        eOutbox_PCF = -3
        eDrafts_PCF = -4
        eDeleted_PCF = -5
    End Enum

    Public PCF_ID As Int32
    Public PlayerID As Int32
    Public FolderName As String

    Public PlayerMsgs() As PlayerComm
    Public PlayerMsgsIdx() As Int32
    Public PlayerMsgUB As Int32 = -1

    Public bHasBeenSorted As Boolean = False

    Public Sub SortEmails()
        If bHasBeenSorted = True Then Return

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1

        If PlayerMsgs Is Nothing OrElse PlayerMsgsIdx Is Nothing Then Return

        Dim lItemUB As Int32 = Math.Min(PlayerMsgs.GetUpperBound(0), PlayerMsgsIdx.GetUpperBound(0))
        bHasBeenSorted = True

        For X As Int32 = 0 To lItemUB
            Dim lIdx As Int32 = -1

            If PlayerMsgsIdx(X) <> -1 AndAlso PlayerMsgs(X) Is Nothing = False Then

                For Y As Int32 = 0 To lSortedUB
                    If PlayerMsgsIdx(Y) < PlayerMsgsIdx(X) Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    bHasBeenSorted = False
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            End If
        Next X

        Dim oMsgs(lSortedUB) As PlayerComm
        Dim lIndices(lSortedUB) As Int32
        Try
            For X As Int32 = 0 To lSortedUB
                oMsgs(X) = PlayerMsgs(lSorted(X))
                lIndices(X) = PlayerMsgsIdx(lSorted(X))
            Next X
            PlayerMsgUB = lSortedUB
            PlayerMsgs = oMsgs
            PlayerMsgsIdx = lIndices
        Catch
        End Try

        'bHasBeenSorted = True
    End Sub
End Class
