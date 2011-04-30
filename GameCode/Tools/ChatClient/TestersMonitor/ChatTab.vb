Option Strict On

Public Class ChatTab
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
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer
    Private Const SND_ASYNC As Integer = &H1
    Private Const SND_ALIAS As Integer = &H10000


    Public oTab As TabPage
    Public sTabName As String = "Chat"
    Public sChannel As String = ""
    Public lSequenceNumber As Int32 = 0
    Public sMessagePrefix As String = ""
    Public lFilter As Int32
    Public sOverall As String

    Private PreviousPMSenders(4) As String
    Private PreviousPMSenderIdx As Int32 = 0

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

    Public Sub AddLine(ByVal sText As String, ByVal lPlayerID As Int32, ByVal yType As ChatMessageType)

        If yType = ChatMessageType.ePrivateMessage Then
            Dim lTempLoc As Int32 = sText.IndexOf(","c)
            If lTempLoc <> -1 Then
                Dim sTemp As String = sText.Substring(0, lTempLoc)
                If sTemp.EndsWith("tells you") = True Then

                    Dim sNewSender As String = sTemp.Substring(0, sTemp.Length - 10)
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

                    lTempLoc = sNewSender.LastIndexOf("}")
                    If lTempLoc <> -1 Then
                        sNewSender = sNewSender.Substring(lTempLoc + 1).Trim
                    End If
                    lTempLoc = sNewSender.LastIndexOf(" ")
                    If lTempLoc <> -1 Then
                        sNewSender = sNewSender.Substring(lTempLoc + 1).Trim
                    End If
                    PreviousPMSenders(0) = sNewSender


                End If
            End If
        End If


        Dim oRTB As RichTextBox = CType(oTab.Controls(0), RichTextBox)
        If oRTB Is Nothing = False Then
            oRTB.Invoke(New doUpdate(AddressOf delegateUpdateEvents), sText)
        End If
    End Sub

    Private Delegate Sub doUpdate(ByVal sEvent As String)
    Private Sub delegateUpdateEvents(ByVal sEvent As String)
        If sEvent.ToString.Contains("{\i " & glPlayerID.ToString & "} ") Then
            sEvent = sEvent.Replace("\cf4 ", "\cf1 ")
            sEvent = sEvent.Replace("\cf1 ", "\cf3 ")
        Else
            If sEvent.ToString.StartsWith("\cf2 ") Then
                Call PlaySound("SystemNotification", 0, SND_ASYNC Or SND_ALIAS)
            ElseIf sEvent.ToString.StartsWith("\cf4 ") Then
                sEvent = sEvent.Replace("\cf4 ", "\cf1 ")
            End If
            If sEvent.ToString.Contains("<DEV> (General): ") Or sEvent.ToString.Contains("<DSE> (General): ") Then
                sEvent = sEvent.Replace("\cf1 ", "\cf2 ")
            End If
        End If
        If sEvent.ToString.Contains(" (Emperor): ") Then
            'x = ChatMessageType.eSenateMessage
            sEvent = sEvent.Replace("\cf3 ", "\cf5 ")
        End If
        sEvent = "\cf7 " & Now.ToString("hh:mm") & " " & sEvent
        'If moCurrentTab Is Nothing = False Then
        sOverall = sEvent & "\par " & sOverall

        Dim oRTB As System.Windows.Forms.RichTextBox = CType(oTab.Controls(0), RichTextBox)
        If oRTB Is Nothing = False Then
            Dim sFinal As String = "{\rtf1\ansi{\fonttbl\f0\fmodern Arial;}{\colortbl;"

            For X As Int32 = 0 To ChatMessageType.eLast - 1
                sFinal &= GetColorRTF(CType(X, ChatMessageType))
            Next X
            sFinal &= "}\f0\fs20\pard " & sOverall
            oRTB.Rtf = sFinal
        End If
        'End If

    End Sub


    Private Function GetColorRTF(ByVal yType As ChatMessageType) As String
        Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 0, 0)
        Select Case yType
            Case ChatMessageType.eAliasChatMessage
                clrVal = AliasChatColor
            Case ChatMessageType.eAllianceMessage
                clrVal = GuildChatColor
            Case ChatMessageType.eChannelMessage
                clrVal = ChannelChatColor
            Case ChatMessageType.eLocalMessage
                clrVal = LocalChatColor
            Case ChatMessageType.eNotificationMessage
                clrVal = AlertChatColor
            Case ChatMessageType.ePrivateMessage
                clrVal = PMChatColor
            Case ChatMessageType.eSenateMessage
                clrVal = SenateChatColor
            Case ChatMessageType.eSysAdminMessage
                clrVal = AlertChatColor
            Case ChatMessageType.eYourChat
                clrVal = DefaultChatColor
        End Select
        Return "\red" & clrVal.R.ToString & "\green" & clrVal.G.ToString & "\blue" & clrVal.B.ToString & ";"
    End Function

    Public Function ProcessSlashR(ByRef bHasValues As Boolean, ByVal sNewVal As String) As String
        If PreviousPMSenders Is Nothing OrElse PreviousPMSenders(0) = "" Then
            bHasValues = False
            Return ""
        Else
            bHasValues = True
            Return "/pm " & PreviousPMSenders(0) & " " & sNewVal.Substring(2)
        End If
    End Function

    Public Sub New()
        ReDim PreviousPMSenders(4)
        For X As Int32 = 0 To 4
            PreviousPMSenders(X) = ""
        Next
    End Sub

    Public Sub SaveTab()
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= gsUserName & ".cht"
        Dim oINI As InitFile = New InitFile(sFile)

        Dim sHdr As String = "ChatTab" & lSequenceNumber
        oINI.WriteString(sHdr, "TabName", sTabName)
        oINI.WriteString(sHdr, "Channel", sChannel)
        oINI.WriteString(sHdr, "MsgPrefix", sMessagePrefix)
        oINI.WriteString(sHdr, "Filter", lFilter.ToString)
    End Sub
End Class