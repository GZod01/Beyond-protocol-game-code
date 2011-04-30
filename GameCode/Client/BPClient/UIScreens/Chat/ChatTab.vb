Option Strict On

Public Class ChatTab

    Public Const mlLineUB As Int32 = 100

    Public Enum ChatFilter As Integer
        eLocalMessages = 1
        eSysAdminMessages = 2
        eChannelMessages = 4
        eAllianceMessages = 8
        eSenateMessages = 16
        ePMs = 32
        eNotificationMessages = 64
        eAliasChatMessage = 128
    End Enum
    Public Const ChatFilterCount As Int32 = 7

    Public lFilter As Int32
    Public sChannel As String = ""

    Public sTabName As String = "Chat"
    Public sMessagePrefix As String = ""

    Public lSequenceNumber As Int32 = 0

    Public msLines() As String
    Public myMessageType() As ChatMessageType
    Public mlSenderID() As Int32

    Public mlLineUsed As Int32 = 0

	'Public clrMessages() As System.Drawing.Color

    Public PreviousPMSenders(4) As String
    Public PreviousPMSenderIdx As Int32 = 0

	Public bHasNewMessage As Boolean = False

	Private Const ml_MAX_PREV_CHAT_CACHE As Int32 = 20
	Private msPreviousChats(ml_MAX_PREV_CHAT_CACHE - 1) As String
	Private mlPreviousChatUB As Int32 = -1
	Private mlPreviousChatIdx As Int32 = 0
	Public Sub AddPrevChatCache(ByVal sLine As String)
		mlPreviousChatUB += 1
		If mlPreviousChatUB > ml_MAX_PREV_CHAT_CACHE - 1 Then
			mlPreviousChatUB = ml_MAX_PREV_CHAT_CACHE - 1
		End If
		For X As Int32 = mlPreviousChatUB To 1 Step -1
			msPreviousChats(X) = msPreviousChats(X - 1)
		Next X
		msPreviousChats(0) = sLine
		mlPreviousChatIdx = -1
	End Sub
	Public Function GetPreviousChatCache(ByVal lDirChange As Int32) As String
		Dim sResult As String = ""
		mlPreviousChatIdx += lDirChange
		If mlPreviousChatIdx < 0 Then
			Return ""
		ElseIf mlPreviousChatIdx > mlPreviousChatUB Then
			mlPreviousChatIdx = mlPreviousChatUB
		End If
		If mlPreviousChatIdx < 0 Then Return ""
		sResult = msPreviousChats(mlPreviousChatIdx)
		Return sResult
	End Function


    Public Sub SaveTab()
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= goCurrentPlayer.PlayerName & ".cht"
        Dim oINI As InitFile = New InitFile(sFile)

        Dim sHdr As String = "ChatTab" & lSequenceNumber
        oINI.WriteString(sHdr, "TabName", sTabName)
        oINI.WriteString(sHdr, "Channel", sChannel)
        oINI.WriteString(sHdr, "MsgPrefix", sMessagePrefix)
        oINI.WriteString(sHdr, "Filter", lFilter.ToString)

		'For X As Int32 = 0 To clrMessages.GetUpperBound(0)
		'    With clrMessages(X)
		'        oINI.WriteString(sHdr, "Color" & X & "_R", .R.ToString)
		'        oINI.WriteString(sHdr, "Color" & X & "_G", .G.ToString)
		'        oINI.WriteString(sHdr, "Color" & X & "_B", .B.ToString)
		'    End With
		'Next X

        oINI = Nothing
    End Sub

    Public Sub AddLine(ByVal sText As String, ByVal lPlayerID As Int32, ByVal yType As ChatMessageType)
        Dim X As Int32
        If msLines Is Nothing Then
            ReDim msLines(mlLineUB)
            ReDim myMessageType(mlLineUB)
            ReDim mlSenderID(mlLineUB)
        End If
        For X = mlLineUsed To 0 Step -1
            If X < mlLineUB Then
                msLines(X + 1) = msLines(X)
                myMessageType(X + 1) = myMessageType(X)
                mlSenderID(X + 1) = mlSenderID(X)
            End If
        Next X
        If mlLineUsed <> mlLineUB Then mlLineUsed += 1
        msLines(0) = sText
        myMessageType(0) = yType
        mlSenderID(0) = lPlayerID

        If yType = ChatMessageType.ePrivateMessage Then
            Dim lTempLoc As Int32 = sText.IndexOf(","c)

            If lTempLoc <> -1 Then
                Dim sTemp As String = sText.Substring(0, lTempLoc)
                If sTemp.EndsWith("tells you") = True Then

                    Dim sNewSender As String = sTemp.Substring(0, sTemp.Length - 10)
                    If muSettings.ChatTimeStamps = True Then
                        sNewSender = sNewSender.Substring(6)
                    End If
                    Dim lIdx As Int32 = sNewSender.IndexOf("<")
                    If lIdx > -1 Then
                        sNewSender = sNewSender.Substring(0, lIdx - 1)
                    End If

                    Dim bFound As Boolean = False
                    For lTmpSender As Int32 = 0 To PreviousPMSenders.GetUpperBound(0)
                        If PreviousPMSenders(lTmpSender).ToUpper = sNewSender.ToUpper Then
                            bFound = True
                            For lTmpIdx As Int32 = lTmpSender To 1 Step -1
                                PreviousPMSenders(lTmpIdx) = PreviousPMSenders(lTmpIdx - 1)
                            Next lTmpIdx
                            Exit For
                        End If
                    Next lTmpSender

                    If bFound = False Then
                        For lTmpSender As Int32 = PreviousPMSenders.GetUpperBound(0) To 1 Step -1
                            PreviousPMSenders(lTmpSender) = PreviousPMSenders(lTmpSender - 1)
                        Next lTmpSender
                    End If
                    PreviousPMSenders(0) = sNewSender

                End If
            End If
        End If

        bHasNewMessage = True
    End Sub

    Public Sub ClearLines()
        ReDim msLines(mlLineUB)
        ReDim myMessageType(mlLineUB)
        ReDim mlSenderID(mlLineUB)
    End Sub

    Public Sub New()
        ReDim msLines(mlLineUB)
        ReDim myMessageType(mlLineUB)
        ReDim mlSenderID(mlLineUB) 

        For X As Int32 = 0 To PreviousPMSenders.GetUpperBound(0)
            PreviousPMSenders(X) = ""
        Next
		'ReDim clrMessages(ChatFilterCount - 1)

        'Now, the defaults...
		'clrMessages(ChatMessageType.eLocalMessage) = System.Drawing.Color.FromArgb(255, 0, 255, 255)
		'clrMessages(ChatMessageType.eSysAdminMessage) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
		'clrMessages(ChatMessageType.eChannelMessage) = System.Drawing.Color.FromArgb(255, 255, 128, 255)
		'clrMessages(ChatMessageType.eAllianceMessage) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
		'clrMessages(ChatMessageType.eSenateMessage) = System.Drawing.Color.FromArgb(255, 192, 192, 192)
		'clrMessages(ChatMessageType.ePrivateMessage) = System.Drawing.Color.FromArgb(255, 255, 128, 0)
		'clrMessages(ChatMessageType.eNotificationMessage) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
    End Sub

    Public Function AcceptsLine(ByVal lLineFilter As Int32, ByVal sLineChannel As String) As Boolean
        If (lFilter And lLineFilter) <> 0 Then
            If lLineFilter = ChatFilter.eChannelMessages Then
                If sLineChannel.ToLower.Trim = sChannel.ToLower OrElse sChannel.Trim = "" Then
                    Return True
                End If
            Else : Return True
            End If
        End If
        Return False
    End Function
End Class