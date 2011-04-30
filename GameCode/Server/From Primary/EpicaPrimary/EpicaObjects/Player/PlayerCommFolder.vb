Option Strict On

Public Class PlayerCommFolder
    Public Enum ePCF_ID_HARDCODES As Integer
        eInbox_PCF = -2
        eOutbox_PCF = -3
        eDrafts_PCF = -4
        eDeleted_PCF = -5
    End Enum

    Public PCF_ID As Int32 = 0             'the unique identifier for this player comm folder, 0 indicates unsaved
    Public PlayerID As Int32 = -1
    Public FolderName(19) As Byte

    'These are the Messages in this folder
    Public PlayerMsgs() As PlayerComm
    Public PlayerMsgsIdx() As Int32
    Public PlayerMsgsUB As Int32 = -1

    Private moPlayer As Player = Nothing
    Public ReadOnly Property oPlayer() As Player
        Get
            If moPlayer Is Nothing OrElse moPlayer.ObjectID <> PlayerID Then moPlayer = GetEpicaPlayer(PlayerID)
            Return moPlayer
        End Get
    End Property

    Public Function AddPlayerComm(ByVal lPC_ID As Int32, ByVal iEncryption As Short, ByVal sBody As String, ByVal sTitle As String, ByVal lSentByID As Int32, ByVal lSentOn As Int32, ByVal bRead As Boolean, ByVal sRecipients As String) As PlayerComm
        Dim X As Int32
        Dim lIdx As Int32

        lIdx = -1
        For X = 0 To PlayerMsgsUB
            'check if any messages were deleted
            If lPC_ID <> -1 AndAlso PlayerMsgsIdx(X) = lPC_ID Then
                'msg is already here, do not add it again
                Return PlayerMsgs(X)
            ElseIf lIdx = -1 AndAlso PlayerMsgsIdx(X) = -1 Then
                lIdx = X
                If lPC_ID = -1 Then Exit For
            End If
        Next X

        If lIdx = -1 Then
            'we need to add a new memory location
            PlayerMsgsUB += 1
            ReDim Preserve PlayerMsgs(PlayerMsgsUB)
            ReDim Preserve PlayerMsgsIdx(PlayerMsgsUB)
            PlayerMsgs(PlayerMsgsUB) = New PlayerComm()
            lIdx = PlayerMsgsUB
        Else : PlayerMsgs(lIdx) = New PlayerComm()
        End If

        With PlayerMsgs(lIdx)
            .ObjectID = lPC_ID
            .ObjTypeID = ObjectType.ePlayerComm
            .EncryptLevel = iEncryption
            .MsgBody = StringToBytes(sBody)
            .MsgTitle = StringToBytes(sTitle)
            .SentByID = lSentByID
            .SentOn = lSentOn
            .SentToList = StringToBytes(sRecipients)
            If bRead = True Then .MsgRead = 1 Else .MsgRead = 0
            .PCF_ID = Me.PCF_ID
            .PlayerID = Me.PlayerID

            If lPC_ID < 1 Then
                If .SaveObject(Me.PlayerID) = False Then Return Nothing
            End If
            PlayerMsgsIdx(lIdx) = .ObjectID
        End With

        Return PlayerMsgs(lIdx)
    End Function

    Public Sub RemovePlayerComm(ByVal lPC_ID As Int32)
        Dim X As Int32
        For X = 0 To PlayerMsgsUB
            If PlayerMsgsIdx(X) = lPC_ID Then
                PlayerMsgs(X) = Nothing
                PlayerMsgsIdx(X) = -1
                Exit For
            End If
        Next X
    End Sub

    Public Function SaveFolder() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            If PCF_ID = 0 OrElse PCF_ID = -1 Then
                'We only INSERT PlayerCommFolders and only if their PCF_ID is 0 or -1 indicate an unsaved, PLAYER-made folder
                sSQL = "INSERT INTO tblPlayerCommFolder (PlayerID, FolderName) VALUES (" & PlayerID & ", '" & _
                    MakeDBStr(BytesToString(FolderName)) & "')"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PCF_ID) FROM tblPlayerCommFolder WHERE FolderName = '" & MakeDBStr(BytesToString(FolderName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    PCF_ID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End If

            'Now that that is done, we'll save all of the emails in this folder
            For X As Int32 = 0 To PlayerMsgsUB
				If PlayerMsgsIdx(X) > 0 Then
					PlayerMsgs(X).PCF_ID = Me.PCF_ID
					PlayerMsgs(X).PlayerID = Me.PlayerID
                    PlayerMsgs(X).SaveObject(Me.PlayerID)
				End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save PlayerCommFolder object. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function GetAddFolderMsg() As Byte()
        Dim yMsg(29) As Byte

		System.BitConverter.GetBytes(GlobalMessageCode.eAddPlayerCommFolder).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(PCF_ID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(PlayerID).CopyTo(yMsg, 6)
        FolderName.CopyTo(yMsg, 10)

        Return yMsg
    End Function

    Public Sub MoveMessageToFolder(ByVal lMsgID As Int32, ByRef oFolder As PlayerCommFolder)
        If oFolder Is Nothing Then Return

        Dim lToIdx As Int32 = -1
        For X As Int32 = 0 To oFolder.PlayerMsgsUB
            If oFolder.PlayerMsgsIdx(X) = lMsgID Then
                Return
            ElseIf lToIdx = -1 AndAlso oFolder.PlayerMsgsIdx(X) = -1 Then
                lToIdx = X
            End If
        Next X
        If lToIdx = -1 Then
            oFolder.PlayerMsgsUB += 1
            ReDim Preserve oFolder.PlayerMsgsIdx(oFolder.PlayerMsgsUB)
            ReDim Preserve oFolder.PlayerMsgs(oFolder.PlayerMsgsUB)
            lToIdx = oFolder.PlayerMsgsUB
        End If

        For X As Int32 = 0 To PlayerMsgsUB
            If PlayerMsgsIdx(X) = lMsgID Then
                'Ok, add it
                PlayerMsgs(X).PCF_ID = oFolder.PCF_ID
                oFolder.PlayerMsgs(lToIdx) = PlayerMsgs(X)
                oFolder.PlayerMsgsIdx(lToIdx) = lMsgID

                PlayerMsgs(X).SaveObject(Me.PlayerID)

                PlayerMsgs(X) = Nothing
                PlayerMsgsIdx(X) = -1
                Exit For
            End If
        Next X

    End Sub

    Public Sub DeleteFolder()
        Dim sSQL As String = "DELETE FROM tblPlayerCommFolder WHERE PCF_ID = " & Me.PCF_ID
        Dim oComm As OleDb.OleDbCommand

        Try
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "DeleteFolder: " & ex.Message)
        End Try
        oComm = Nothing

        Try
            sSQL = "DELETE FROM tblPlayerComm WHERE PCF_ID = " & Me.PCF_ID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "DeleteFolderContents: " & ex.Message)
        End Try
        oComm = Nothing

    End Sub
End Class
