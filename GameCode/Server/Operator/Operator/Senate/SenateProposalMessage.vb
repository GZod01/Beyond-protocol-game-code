Option Strict On

Public Structure SenateProposalMessage
    Public lPosterID As Int32
    Public yMsgData() As Byte
    Public lPostedOn As Int32           'date in GMT

    Private moPostedBy As Player
    Public ReadOnly Property PostedBy() As Player
        Get
            If moPostedBy Is Nothing OrElse moPostedBy.ObjectID <> lPosterID Then
                moPostedBy = GetEpicaPlayer(lPosterID)
            End If
            Return moPostedBy
        End Get
    End Property

    Public Function GetMsgDetail(ByVal lProposalID As Int32, ByVal lForPlayerID As Int32) As Byte()
        Dim lMsgLen As Int32 = 0
        If yMsgData Is Nothing = False Then lMsgLen = yMsgData.Length

        Dim yMsg(42 + lMsgLen) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2      'msgcode?

        System.BitConverter.GetBytes(lForPlayerID).CopyTo(yMsg, lPos) : lPos += 4

        yMsg(lPos) = eySenateRequestDetailsType.ProposalMsg : lPos += 1
        System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPosterID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPostedOn).CopyTo(yMsg, lPos) : lPos += 4

        If PostedBy Is Nothing = False Then
            PostedBy.PlayerName.CopyTo(yMsg, lPos)
        End If
        lPos += 20

        System.BitConverter.GetBytes(lMsgLen).CopyTo(yMsg, lPos) : lPos += 4
        If yMsgData Is Nothing = False Then yMsgData.CopyTo(yMsg, lPos)
        lPos += lMsgLen

        Return yMsg
    End Function

    Public Function SaveObject(ByVal lProposalID As Int32) As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try

            sSQL = "INSERT INTO tblSenateProposalMsg (ProposalID, PosterID, PostedOn, MsgData) VALUES (" & lProposalID & _
             ", " & lPosterID & ", " & lPostedOn & ", '" & MakeDBStr(BytesToString(yMsgData)) & "')"

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If

            bResult = True
        Catch ex As Exception
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object SenateProposalMessage. Reason: " & ex.Message)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function GetSaveObjectText(ByVal lProposalID As Int32) As String
        Try
            Return "INSERT INTO tblSenateProposalMsg (ProposalID, PosterID, PostedOn, MsgData) VALUES (" & lProposalID & _
               ", " & lPosterID & ", " & lPostedOn & ", '" & MakeDBStr(BytesToString(yMsgData)) & "')"
        Catch ex As Exception
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object SenateProposalMessage. Reason: " & ex.Message)
        End Try
        Return ""
    End Function
End Structure

