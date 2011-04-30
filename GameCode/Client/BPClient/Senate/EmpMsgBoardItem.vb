Option Strict On

Public Class EmpMsgBoardItem
    Public lPosterID As Int32
    Public lPostedOn As Int32

    Public sMsg As String = ""

    Private mbDetailsRequested As Boolean = False
    Public Sub RequestDetails()
        If mbDetailsRequested = True Then Return
        mbDetailsRequested = True
        Dim yMsg(14) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
        lPos += 4       'leave room for PlayerId
        yMsg(lPos) = eySenateRequestDetailsType.EmpChmbrMsg : lPos += 1
        System.BitConverter.GetBytes(lPosterID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPostedOn).CopyTo(yMsg, lPos) : lPos += 4
        goUILib.SendMsgToPrimary(yMsg)
    End Sub
End Class
