
Public Structure PrimaryRecipients
    Public lPrimaryIdx As Int32
    Public lRecipients() As Int32
End Structure

Public Class ChatManager

    Public Enum ChatMessageType As Byte
        LocalMessage = 0
        SysAdminMessage = 1
        ChannelMessage = 2
        GuildChat = 3
        SenateMessage = 4
        PrivateMessage = 5
        NotificationMessage = 6
        AliasChatMessage = 7
    End Enum

    Public Shared Function GetPlayerNameWithTag(ByVal sName As String, ByVal lPlayerID As Int32) As String
        Dim sTag As String = ""
        If lPlayerID = 7 Then
            sTag = " <DSE>"
        ElseIf lPlayerID = 1 OrElse lPlayerID = 2 OrElse lPlayerID = 6 OrElse lPlayerID = 21296 Then
            sTag = " <DEV>"
        End If
        Return sName & sTag
    End Function

    Public Shared Function CreateChatMsg(ByVal lFromPlayer As Int32, ByVal sMsg As String, ByVal yType As ChatMessageType, ByVal bLogEvent As Boolean, ByVal sChannelName As String) As Byte()
        Dim yMsg() As Byte

        If bLogEvent = True Then LogChatMsg(lFromPlayer, sMsg, yType, sChannelName)

        ReDim yMsg(10 + sMsg.Length)
        System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lFromPlayer).CopyTo(yMsg, 2)
        yMsg(6) = yType
        System.BitConverter.GetBytes(sMsg.Length).CopyTo(yMsg, 7)
        StringToBytes(sMsg).CopyTo(yMsg, 11)
        Return yMsg
    End Function

    Public Shared Function AppendRecipientList(ByVal yChatMsg() As Byte, ByVal lRecipients() As Int32) As Byte()
        Dim lLen As Int32 = yChatMsg.GetUpperBound(0) + 4 + (lRecipients.Length * 4)
        Dim yFinal(lLen) As Byte
        yChatMsg.CopyTo(yFinal, 0)
        Dim lPos As Int32 = yChatMsg.Length
        System.BitConverter.GetBytes(lRecipients.Length).CopyTo(yFinal, lPos) : lPos += 4
        For X As Int32 = 0 To lRecipients.GetUpperBound(0)
            System.BitConverter.GetBytes(lRecipients(X)).CopyTo(yFinal, lPos) : lPos += 4
        Next X
        Return yFinal
    End Function

    Private Shared oFS As IO.FileStream
    Private Shared oWriter As IO.StreamWriter
    Private Shared lVals(-1) As Int32       'just a placeholder for something to synclock on
    Private Shared msPreviousFileName As String
    Private Shared mlPreviousFileNum As Int32 = 0
    Private Shared mbForceCloseLog As Boolean = False

    Private Shared Sub LogChatMsg(ByVal lFromPlayer As Int32, ByVal sMsg As String, ByVal yType As ChatMessageType, ByVal sChannel As String)
        'need to log FromPlayerID and the Msg
        Const ml_MAX_LOG_FILE_SIZE As Int32 = 5000000

        SyncLock lVals
            If oFS Is Nothing Then
                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                If sPath.EndsWith("\") = False Then sPath &= "\"
                mlPreviousFileNum = 0
                Dim sNewFileName As String = "Chat_" & Now.ToString("MMddyyyyHHmmss")
                oFS = New IO.FileStream(sPath & sNewFileName & "_" & mlPreviousFileNum & ".log", IO.FileMode.Create)
                msPreviousFileName = sNewFileName
                oWriter = New IO.StreamWriter(oFS)
                oWriter.AutoFlush = True
            End If

            oWriter.WriteLine(Now.ToString & "|" & lFromPlayer & "|" & CByte(yType) & "|" & sChannel & "|" & sMsg)

            If oFS.Length > ml_MAX_LOG_FILE_SIZE OrElse mbForceCloseLog = True Then
                mbForceCloseLog = False
                oWriter.Close()
                oFS.Close()
                oWriter.Dispose()
                oFS.Dispose()
                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                If sPath.EndsWith("\") = False Then sPath &= "\"
                Dim sNewFileName As String = "Chat_" & Now.ToString("MMddyyyyHHmmss")
                If msPreviousFileName = sNewFileName Then
                    mlPreviousFileNum += 1
                Else
                    mlPreviousFileNum = 0
                End If
                oFS = New IO.FileStream(sPath & sNewFileName & "_" & mlPreviousFileNum & ".log", IO.FileMode.Create)
                msPreviousFileName = sNewFileName
                oWriter = New IO.StreamWriter(oFS)
            End If
        End SyncLock
    End Sub

    Public Shared Sub ResetLog()
        mbForceCloseLog = True
        LogChatMsg(-1, "Reset Log Fired", ChatMessageType.NotificationMessage, "")
    End Sub

    ''' <summary>
    ''' Parses a chat message from a primary server
    ''' </summary>
    ''' <param name="lIndex"> Index of the Primary Server sending the message </param>
    ''' <param name="yData"> Contents of the message </param>
    ''' <remarks></remarks>
    Public Shared Sub HandleChatMsg(ByVal lIndex As Int32, ByVal yData() As Byte)
        'Ok, the message coming to us from the Primary Server is:
        'MsgCode (2)
        'SenderPlayerID (4)
        'MessageLen (4)
        'Message()

        Dim lPos As Int32 = 2       'for our msgcode
        Dim lSenderID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMsgLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sMsg As String = GetStringFromBytes(yData, lPos, lMsgLen) : lPos += lMsgLen
        sMsg = sMsg.Replace(vbCr, "").Replace(vbLf, "")
        Dim sMsgUpper As String = sMsg.ToUpper

        Dim oSender As Player = GetPlayer(lSenderID)
        If oSender Is Nothing Then
            LogEvent("ERROR: HandleChatMsg. No Player returned for SenderID of " & lSenderID)
            Return
        End If
        Dim sPlayerName As String = oSender.sPlayerName

        If oSender.dtCoolDownEnds <> Date.MinValue AndAlso oSender.dtCoolDownEnds > Now Then
            Dim lTmpRecpt() As Int32 = {lSenderID}
            Dim lSecondsRemain As Int32 = CInt(oSender.dtCoolDownEnds.Subtract(Now).TotalSeconds)
            Dim lHours As Int32 = lSecondsRemain \ 3600
            lSecondsRemain -= (lHours * 3600)
            Dim lMinutes As Int32 = lSecondsRemain \ 60
            lSecondsRemain -= (lMinutes * 60)
            Dim sRetStr As String = "You are in Cooldown and are unable to chat for "
            If lHours <> 0 Then sRetStr &= lHours & " hours "
            If lMinutes <> 0 Then sRetStr &= lMinutes & " minutes "
            If lSecondsRemain <> 0 Then sRetStr &= lSecondsRemain & " seconds."
            Dim yFinal() As Byte = AppendRecipientList(CreateChatMsg(lSenderID, sRetStr, ChatMessageType.SysAdminMessage, False, ""), lTmpRecpt)
            goMsgSys.SendToPrimary(lIndex, yFinal)
            Return
        End If

        'Ok, we'll take a page out of other MMO's
        'Dim sCommand As String = UCase$(Mid$(sMsg, 1, 3))
        Dim sParm1 As String = ""
        Dim sParm2 As String = ""
        Dim lTemp1 As Int32
        Dim lTemp2 As Int32

        Dim bOverrideSysAdminRights As Boolean = False

        'Ok, determine what our command is
        Dim lCmdType As ChatManager.ChatMessageType
        If sMsgUpper.StartsWith("/GU ") = True OrElse sMsgUpper.StartsWith("/GUILD ") = True Then
            lCmdType = ChatMessageType.GuildChat
        ElseIf sMsgUpper.StartsWith("/AL ") = True OrElse sMsgUpper.StartsWith("/ALIAS ") = True Then
            lCmdType = ChatMessageType.AliasChatMessage
        ElseIf sMsgUpper.StartsWith("/PM ") = True OrElse sMsgUpper.StartsWith("/TELL ") = True OrElse sMsgUpper.StartsWith("/T ") Then
            lCmdType = ChatMessageType.PrivateMessage
        ElseIf sMsgUpper.StartsWith("/LOCAL") = True OrElse sMsgUpper.StartsWith("/LO ") = True OrElse sMsgUpper.StartsWith("/") = False Then
            lCmdType = ChatMessageType.LocalMessage
        ElseIf sMsgUpper.StartsWith("/SYSTEM ") = True Then
            lCmdType = ChatMessageType.SysAdminMessage
        ElseIf sMsgUpper.StartsWith("/COOLDOWN") = True Then
            'ok, a csr is telling us to put a player on cooldown... syntax is /cooldown <playername>
            If lSenderID <> 221 AndAlso lSenderID <> 131 AndAlso lSenderID <> 2 AndAlso lSenderID <> 6 AndAlso lSenderID <> 1 AndAlso lSenderID <> 2067 AndAlso lSenderID <> 3253 AndAlso lSenderID <> 7 AndAlso lSenderID <> 80 AndAlso lSenderID <> 653 AndAlso lSenderID <> 2076 AndAlso lSenderID <> 1780 Then
                Dim lTmpRecpt() As Int32 = {lSenderID}
                Dim yFinal() As Byte = AppendRecipientList(CreateChatMsg(lSenderID, "Invalid Command, command unrecognized.", ChatMessageType.SysAdminMessage, False, ""), lTmpRecpt)
                goMsgSys.SendToPrimary(lIndex, yFinal)
                Return
            Else
                Dim lFirstSpace As Int32 = sMsgUpper.IndexOf(" "c)
                If lFirstSpace < 0 Then Return
                sParm1 = sMsg.Substring(lFirstSpace).Trim.ToUpper

                'ok, get that player...
                Dim bFound As Boolean = False
                For X As Int32 = 0 To glPlayerUB
                    If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).sCompareName = sParm1 Then
                        'k, found them
                        bFound = True

                        Dim oTarget As Player = goPlayer(X)

                        If oTarget.dtCoolDownEnds <> Date.MinValue Then
                            If lSenderID = 1 OrElse lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 7 OrElse lSenderID = oTarget.lCoolDownBy OrElse oTarget.lCoolDownBy = -1 OrElse lSenderID = 21296 Then
                                sMsg = "/SYSTEM " & oSender.sPlayerName & " has removed the cool down on " & oTarget.sPlayerName & "."
                                oTarget.dtCoolDownEnds = Date.MinValue
                                oTarget.lCoolDownBy = -1
                            Else
                                Dim lTmpRecpt() As Int32 = {lSenderID}
                                Dim yFinal() As Byte = AppendRecipientList(CreateChatMsg(lSenderID, "Unable to remove cooldown because another GM started it.", ChatMessageType.SysAdminMessage, False, ""), lTmpRecpt)
                                goMsgSys.SendToPrimary(lIndex, yFinal)
                                Return
                            End If
                        Else
                            oTarget.dtCoolDownEnds = Now.AddMinutes(60)
                            oTarget.lCoolDownBy = lSenderID
                            sMsg = "/SYSTEM " & oSender.sPlayerName & " has put " & oTarget.sPlayerName & " on cool down for 60 minutes. " & oTarget.sPlayerName & " will be unable to chat to anyone until the cool down is over."
                        End If

                        lCmdType = ChatMessageType.SysAdminMessage
                        bOverrideSysAdminRights = True
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    Dim lTmpRecpt() As Int32 = {lSenderID}
                    Dim yFinal() As Byte = AppendRecipientList(CreateChatMsg(lSenderID, "Unable to find a player by that name.", ChatMessageType.SysAdminMessage, False, ""), lTmpRecpt)
                    goMsgSys.SendToPrimary(lIndex, yFinal)
                    Return
                End If
            End If
        Else
            lCmdType = ChatMessageType.ChannelMessage
        End If


        lTemp1 = InStr(sMsg, " ", CompareMethod.Binary)
        If lTemp1 <> 0 AndAlso lCmdType = ChatMessageType.PrivateMessage Then lTemp2 = InStr(lTemp1 + 1, sMsg, " ", CompareMethod.Binary)
        If lTemp1 > 0 AndAlso lTemp2 > 0 Then
            sParm1 = Trim$(Mid$(sMsg, lTemp1 + 1, lTemp2 - lTemp1 - 1))
            sParm2 = Trim$(Mid$(sMsg, lTemp2 + 1))
        ElseIf lTemp1 > 0 Then
            sParm1 = Trim$(Mid$(sMsg, lTemp1 + 1))
            sParm2 = ""
            'Else
            'moClients(lIndex).SendData(CreateChatMsg(-1, "Unknown command or invalid syntax"))
        End If

        Dim yMsg() As Byte = Nothing
        Dim lRecipients() As Int32 = Nothing
        Select Case lCmdType
            Case ChatMessageType.GuildChat
                'ok, guild msg is formatted /cmd msg

                'First, is the player in a guild?
                If oSender.oGuild Is Nothing = False Then
                    'so, get our first space...
                    Dim lFirstSpace As Int32 = sMsg.IndexOf(" "c)
                    If lFirstSpace < 0 Then Return
                    sParm1 = sMsg.Substring(lFirstSpace).Trim

                    yMsg = CreateChatMsg(lSenderID, sPlayerName & " (GUILD): " & sParm1, ChatMessageType.GuildChat, True, oSender.oGuild.lGuildID.ToString)
                    Dim lPrimaryList() As Int32 = oSender.oGuild.GetPrimaryIdxList()
                    If lPrimaryList Is Nothing Then
                        LogEvent("Error: HandleChatMsg. GetPrimaryIdxList for Guild returned nothing.")
                        Return
                    End If
                    For X As Int32 = 0 To lPrimaryList.GetUpperBound(0)
                        lRecipients = oSender.oGuild.GetRecipientList(lPrimaryList(X))
                        If lRecipients Is Nothing = False Then
                            Dim yFinal() As Byte = AppendRecipientList(yMsg, lRecipients)
                            goMsgSys.SendToPrimary(lPrimaryList(X), yFinal)
                        End If
                    Next X

                    'Ensure yMsg and lRecipients are nothing, not further action is needed
                    yMsg = Nothing
                    lRecipients = Nothing
                Else
                    yMsg = CreateChatMsg(lSenderID, "You are not in a guild.", ChatMessageType.GuildChat, False, "")
                    ReDim lRecipients(0)
                    lRecipients(0) = lSenderID
                End If
            Case ChatMessageType.PrivateMessage
                '/pm NAME MSG
                If sParm1 <> "" AndAlso sParm2 <> "" Then
                    sParm1 = sParm1.ToUpper

                    Dim bFound As Boolean = False
                    For X As Int32 = 0 To glPlayerUB
                        If glPlayerIdx(X) <> -1 AndAlso goPlayer(X) Is Nothing = False AndAlso goPlayer(X).sCompareName = sParm1 Then
                            'k, found them
                            bFound = True
                            yMsg = CreateChatMsg(lSenderID, GetPlayerNameWithTag(sPlayerName, lSenderID) & " tells you, '" & sParm2 & "'", ChatMessageType.PrivateMessage, False, "")
                            ReDim lRecipients(0) : lRecipients(0) = goPlayer(X).PlayerID
                            Dim yFinal() As Byte = AppendRecipientList(yMsg, lRecipients)
                            goMsgSys.SendToPrimary(goPlayer(X).lPrimaryIdx, yFinal)

                            'Now, send our msg to the sender
                            yMsg = CreateChatMsg(lSenderID, "You tell " & goPlayer(X).sPlayerName & ", '" & sParm2 & "'", ChatMessageType.PrivateMessage, True, sPlayerName)
                            ReDim lRecipients(0) : lRecipients(0) = lSenderID

                            Exit For
                        End If
                    Next X

                    If bFound = False Then
                        yMsg = CreateChatMsg(lSenderID, "Unable to find a player by that name.", ChatMessageType.PrivateMessage, False, "")
                        ReDim lRecipients(0) : lRecipients(0) = lSenderID
                    End If
                Else
                    'moClients(lIndex).SendData(CreateChatMsg(-1, "Invalid Command, the Proper syntax is /pm PLAYERNAME, MSG", ChatMessageType.ePrivateMessage))
                    yMsg = CreateChatMsg(lSenderID, "Invalid Command, the Proper syntax is /pm PLAYERNAME, MSG", ChatMessageType.PrivateMessage, False, "")
                    ReDim lRecipients(0) : lRecipients(0) = lSenderID
                End If

            Case ChatMessageType.AliasChatMessage
                '/Alias MSG
                Dim lFirstSpace As Int32 = sMsg.IndexOf(" "c)
                If lFirstSpace < 0 Then Return
                sParm1 = sMsg.Substring(lFirstSpace).Trim
                yMsg = CreateChatMsg(lSenderID, sPlayerName & " (ALIAS): " & sParm1, ChatMessageType.AliasChatMessage, True, "")

                'am I aliasing someone?
                Dim oPlayer As Player = oSender
                If oSender.oAliasedAs Is Nothing = False Then
                    'yes, I am... set it to the sender's aliased as
                    oPlayer = oSender.oAliasedAs
                End If

                Dim lPrimaryList() As Int32 = oPlayer.GetPrimaryIdxList()
                If lPrimaryList Is Nothing = True Then
                    LogEvent("Error: HandleChatMsg. GetPrimaryIdxList for Alias returned nothing.")
                    Return
                End If

                For X As Int32 = 0 To lPrimaryList.GetUpperBound(0)
                    lRecipients = oPlayer.GetRecipientList(lPrimaryList(X))
                    If lRecipients Is Nothing = False Then
                        Dim yFinal() As Byte = AppendRecipientList(yMsg, lRecipients)
                        goMsgSys.SendToPrimary(lPrimaryList(X), yFinal)
                    End If
                Next X

                'Ensure yMsg and lRecipients are nothing, not further action is needed
                yMsg = Nothing
                lRecipients = Nothing
            Case ChatMessageType.LocalMessage
                If sMsg.StartsWith("/") = False Then sParm1 = sMsg
                If oSender.oLocalChatChannel Is Nothing = False Then
                    yMsg = CreateChatMsg(lSenderID, GetPlayerNameWithTag(sPlayerName, lSenderID) & ": " & sParm1, ChatMessageType.LocalMessage, True, oSender.oLocalChatChannel.sChannelName)
                    lRecipients = oSender.oLocalChatChannel.GetRecipientList(lIndex)
                End If
            Case ChatMessageType.SysAdminMessage
                If oSender.bSystemAdmin = False AndAlso bOverrideSysAdminRights = False Then
                    LogEvent("Possible Cheat: HandleChatMsg send SysAdminMessage. Player is not SysAdmin. PlayerID: " & lSenderID)
                    Return
                End If
                yMsg = CreateChatMsg(-1, sParm1, ChatMessageType.SysAdminMessage, True, "")
                ReDim lRecipients(0) : lRecipients(0) = -1
                Dim yFinal() As Byte = AppendRecipientList(yMsg, lRecipients)
                goMsgSys.SendToAllPrimaries(yFinal)

                lRecipients = Nothing
                yMsg = Nothing
            Case ChatMessageType.ChannelMessage

                lTemp1 = InStr(1, sMsg, " ", CompareMethod.Binary)
                If lTemp1 = 0 Then Return
                Dim sSearchName As String = UCase$(Trim$(Mid$(sMsg, 2, lTemp1 - 2)))
                'Recombine Parm1 and Parm2
                If sParm2 <> "" Then sParm1 &= ", " & sParm2

                Dim bFound As Boolean = False
                If oSender.oChannels Is Nothing = False Then
                    For X As Int32 = 0 To oSender.oChannels.GetUpperBound(0)
                        If oSender.oChannels(X) Is Nothing = True Then Continue For
                        If oSender.oChannels(X).sCompareChannelName = sSearchName Then
                            Dim oChannel As ChatRoom = oSender.oChannels(X)
                            'ok, player is member of the channel
                            bFound = True

                            Dim lPrims() As Int32 = oChannel.GetPrimaryIdxList()
                            If lPrims Is Nothing = True Then
                                LogEvent("Error: HandleChatMsg. GetPrimaryIdxList for Alias returned nothing.")
                                Return
                            End If
                            yMsg = CreateChatMsg(lSenderID, GetPlayerNameWithTag(oSender.sPlayerName, lSenderID) & " (" & oChannel.sChannelName & "): " & sParm1, ChatMessageType.ChannelMessage, True, oChannel.sChannelName)

                            If oSender.PlayerID = 1730 AndAlso oSender.oChannels(X).sCompareChannelName = "GENERAL" Then
                                goMsgSys.SendToPrimary(oSender.lPrimaryIdx, yMsg)
                            Else
                                For Y As Int32 = 0 To lPrims.GetUpperBound(0)
                                    lRecipients = oChannel.GetRecipientList(lPrims(Y))
                                    If lRecipients Is Nothing = False Then
                                        Dim yFinal() As Byte = AppendRecipientList(yMsg, lRecipients)
                                        goMsgSys.SendToPrimary(lPrims(Y), yFinal)
                                    End If
                                Next Y
                            End If

                            Exit For
                        End If
                    Next X
                End If
                If bFound = False Then
                    ReDim lRecipients(0) : lRecipients(0) = lSenderID
                    yMsg = CreateChatMsg(lSenderID, "Invalid Command, command unrecognized '" & sSearchName & "'.", ChatMessageType.SysAdminMessage, False, "")
                Else
                    lRecipients = Nothing : yMsg = Nothing
                End If

        End Select

        'If we have a msg here and a recipients list, send the msg to the calling Primary
        If yMsg Is Nothing = False AndAlso lRecipients Is Nothing = False Then
            Dim yFinal() As Byte = AppendRecipientList(yMsg, lRecipients)
            goMsgSys.SendToPrimary(lIndex, yFinal)
        End If

    End Sub

End Class
