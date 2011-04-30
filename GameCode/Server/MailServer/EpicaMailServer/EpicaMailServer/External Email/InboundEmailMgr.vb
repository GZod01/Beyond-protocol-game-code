Option Strict On

'Does all of the inbound-email checking (should be instantiated once)
Public Class InboundEmailMgr
    'Private Const SEE_LOG_FILE As Int32 = 20
    'Private Const SEE_SET_RAWFILE_PREFIX As Int32 = 70
    'Private Const SEE_KEY_CODE As Int32 = 0&

    'Private Declare Function seeAttach Lib "SEE32.DLL" (ByVal NbrChans As Int32, ByVal KeyCode As Int32) As Int32
    'Private Declare Function seeClose Lib "SEE32.DLL" (ByVal Chan As Int32) As Int32
    'Private Declare Function seeDeleteEmail Lib "SEE32.DLL" (ByVal Chan As Int32, ByVal MsgNbr As Int32) As Int32
    'Private Declare Function seeErrorText Lib "SEE32.DLL" (ByVal Chan As Int32, ByVal Code As Int32, ByVal Buffer As String, ByVal BufLen As Int32) As Int32
    'Private Declare Function seeGetEmailCount Lib "SEE32.DLL" (ByVal Chan As Int32) As Int32
    'Private Declare Function seeGetEmailFile Lib "SEE32.DLL" (ByVal Chan As Int32, ByVal MsgNbr As Int32, ByVal FileName As String, ByVal EmailDir As String, ByVal AttachDir As String) As Int32
    'Private Declare Function seeIntegerParam Lib "SEE32.DLL" (ByVal Chan As Int32, ByVal Index As Int32, ByVal Value As Int32) As Int32
    'Private Declare Function seePop3Connect Lib "SEE32.DLL" (ByVal Chan As Int32, ByVal Server As String, ByVal User As String, ByVal Password As String) As Int32
    'Private Declare Function seeRelease Lib "SEE32.DLL" () As Int32
    'Private Declare Function seeStringParam Lib "SEE32.DLL" (ByVal Chan As Int32, ByVal Index As Int32, ByVal Value As String) As Int32

    Public sEmailDir As String
    Public sAttachDir As String
    Private mblID As Int64 = 0

    Public Function RetreiveEmails() As Int32
        Dim bConnected As Boolean = False
        Dim lResult As Int32 = 0
        Try

            Dim oMail As New Chilkat.MailMan
            'set up our credentials...
			oMail.UnlockComponent("MATTHEWCAMMAILQ_XkMwRUTd4Tpy")
            oMail.MailHost = gsInHostName
            oMail.PopPassword = gsEmailPassword
            oMail.PopUsername = gsEmailUserName

            ' Use CopyMail to leave email on the POP3 server,
            ' Use TransferMail to copy and remove it from the server.
            Dim bundle As Chilkat.EmailBundle
            bundle = oMail.TransferMail()       'use transfermail to pull and remove

            Dim sDir As String = AppDomain.CurrentDomain.BaseDirectory
            If sDir.EndsWith("\") = False Then sDir &= "\"

			'Now, get our messages
			If bundle Is Nothing = False Then
				If bundle.MessageCount <> 0 Then
                    LogEvent("Retrieved " & bundle.MessageCount & " inbound messages.")
					For X As Int32 = 0 To bundle.MessageCount - 1
						Dim oEmail As Chilkat.Email = bundle.GetEmail(X)

						'Generate a valid email name
						Dim sFile As String = "Eml" & mblID.ToString & ".eml"
						If oEmail.SaveEml(sDir & sFile) = False Then
							LogEvent("Could not save email file " & sFile)
						End If
					Next X
				End If
				bundle.Dispose()
			End If

			bundle = Nothing
			oMail.Dispose()
			oMail = Nothing
		Catch ex As Exception
			LogEvent(ex.Message)
		End Try
        Return lResult
    End Function

    Public Sub ParseEmailFile()
        Dim sFile As String = Dir(sEmailDir & "*.eml")

        If sFile <> "" Then

            Try
                Dim oEmail As New Chilkat.Email()
                If oEmail.LoadEml(sFile) = False Then
                    LogEvent("Unable to load email file: " & sFile & ". Reason: " & oEmail.LastErrorText)
                Else
                    Dim sTo As String = ""
                    Dim sFrom As String = ""
                    If oEmail.NumTo > 0 Then sTo = oEmail.GetTo(0)
                    sFrom = oEmail.From

                    'Do our cleanup of the FROM address
                    sFrom = sFrom.Replace(">", "").Replace("<", "").Trim
                    Dim lIdx As Int32 = sFrom.LastIndexOf(" "c)
                    If lIdx <> -1 Then sFrom = sFrom.Substring(lIdx).Trim
                    sFrom = sFrom.ToUpper.Replace("FROM:", "").Trim

                    'Do our cleanup of the TO address
                    sTo = sTo.Replace(">", "").Replace("<", "").Trim
                    lIdx = sTo.LastIndexOf(" "c)
                    If lIdx <> -1 Then sTo = sTo.Substring(lIdx).Trim
                    sTo = sTo.ToUpper.Replace("TO:", "").Trim

                    LogEvent("Parsing Inbound message from " & sFrom & " to " & sTo)

                    Dim sBody As String = oEmail.Body

                    If sFrom.ToLower.Contains("support@darkskyentertainment.com") = True Then
                        'Ok, support is responding to an email most likely...
                        'The 2 address will be SupportDirect_Playername@BeyondProtocol.com
                        If sTo.ToUpper.StartsWith("SUPPORTDIRECT_") = True Then
                            Dim sPlayerName As String = sTo.ToUpper.Replace("SUPPORTDIRECT_", "")
                            sPlayerName = sPlayerName.Replace("@BEYONDPROTOCOL.COM", "")

                            'We should have a player name...
                            Dim lPlayerID As Int32 = -1
                            Dim lPrimIdx As Int32 = 0
                            For X As Int32 = 0 To glPlayerUB
                                If glPlayerIdx(X) <> -1 AndAlso goPlayer(X) Is Nothing = False AndAlso goPlayer(X).sCompareName = sPlayerName Then
                                    'k, found them
                                    lPlayerID = goPlayer(X).PlayerID
                                    lPrimIdx = goPlayer(X).lPrimaryIdx
                                    Exit For
                                End If
                            Next X

                            If lPlayerID = -1 Then
                                'Send an external email to support that the player was not found
                                AddNewOutboundMailMsg(gsOutEmailFrom, "support@darkskyentertainment.com", "Undeliverable: No Player by that name", "Mail could not be delivered. There is no player by the name " & sPlayerName & " parsed from " & sTo & "." & vbCrLf & vbCrLf & "Original Message:" & vbCrLf & sBody, -1, -1, -1)
                            Else
                                Dim sData As String = sBody
                                If sData.Length > 5000 Then sData = sData.Substring(0, 5000)
                                Dim yData(17 + sData.Length) As Byte
                                Dim lPos As Int32 = 0
                                System.BitConverter.GetBytes(GlobalMessageCode.eSendOutMailMsg).CopyTo(yData, lPos) : lPos += 2
                                System.BitConverter.GetBytes(-1I).CopyTo(yData, lPos) : lPos += 4
                                System.BitConverter.GetBytes(19326).CopyTo(yData, lPos) : lPos += 4
                                System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                                System.BitConverter.GetBytes(sData.Length).CopyTo(yData, lPos) : lPos += 4
                                StringToBytes(sData).CopyTo(yData, lPos) : lPos += sData.Length
                                goMsgSys.SendToPrimary(lPrimIdx, yData)
                            End If
                        End If
                    End If

                    For X As Int32 = 0 To glMailMsgUB
                        If gyMailMsgUsed(X) <> 0 AndAlso goMailMsgs(X).sTo = sFrom AndAlso goMailMsgs(X).sReplyTo.ToUpper = sTo Then
                            goMailMsgs(X).ProcessResponse(sBody)
                            Exit For
                        End If
                    Next X
                End If

            Catch ex As Exception
                LogEvent(ex.Message)
            Finally
                Try
                    Kill(sEmailDir & sFile)
                Catch ex As Exception
                    LogEvent("Trying to kill inbound email: " & ex.Message)
                End Try
            End Try

        End If

    End Sub
End Class
