Option Strict On

'Does all of the inbound-email checking (should be instantiated once)
Public Class InboundEmailMgr
    Public sEmailDir As String
    Public sAttachDir As String

    Private Structure T_EmailMessage
        Public sTo As String
        Public sFrom As String
        Public sSubject As String
        Public sBody As String
    End Structure
    Private moEmailMessages() As T_EmailMessage

    Public Function RetreiveEmails() As Boolean
        Try
            Dim iMsgCount As Int32 = 0
            Dim email As New Pop3.Pop3Client(gsEmailUserName, gsEmailPassword, gsInHostName)
            email.OpenInbox()
            While (email.NextEmail() And iMsgCount < 100) 'Only read up to 100 messages, * 5000 body, for 500k ram
                ReDim Preserve moEmailMessages(iMsgCount)
                With moEmailMessages(iMsgCount)
                    .sTo = email.To
                    .sFrom = email.From
                    .sSubject = email.Subject
                    If email.Body.Length > 5000 Then
                        .sBody = email.Body.Substring(0, 5000)
                    Else
                        .sBody = email.Body
                    End If

                    'Pop3 client auto parses for plain-text first attachment, no need to continue unless we want to handle Attachments.
                    'If (email.IsMultipart) Then
                    '    Dim enumerator As System.Collections.IEnumerator = email.MultipartEnumerator
                    '    While (enumerator.MoveNext())
                    '        Dim multipart As Pop3.Pop3Component = CType(enumerator.Current, Pop3.Pop3Component)
                    '        'TODO: Handle this
                    '        If (multipart.IsBody) Then
                    '            'Console.WriteLine("Multipart body:" + multipart.Data)
                    '            If multipart.Data.Length > 5000 Then
                    '                .sBody = multipart.Data.Substring(0, 5000)
                    '            Else
                    '                .sBody = multipart.Data
                    '            End If
                    '        Else
                    '            'Console.WriteLine("Attachment name=" + multipart.Name)
                    '        End If
                    '    End While
                    'End If
                End With
                Dim bDeleted As Boolean = email.DeleteEmail()
                iMsgCount += 1
            End While
            email.CloseConnection()
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub ParseEmailFile()

        Try
            Dim sTo As String = ""
            Dim sFrom As String = ""
            Dim oEmailMessage As T_EmailMessage
            For Each oEmailMessage In moEmailMessages
                sTo = oEmailMessage.sTo
                sFrom = oEmailMessage.sFrom
                sFrom = sFrom.Replace(">", "").Replace("<", "").Trim

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

                Dim sBody As String = oEmailMessage.sBody

                For X As Int32 = 0 To glMailMsgUB
                    If gyMailMsgUsed(X) <> 0 AndAlso goMailMsgs(X).sTo = sFrom AndAlso goMailMsgs(X).sReplyTo.ToUpper = sTo Then
                        goMailMsgs(X).ProcessResponse(sBody)
                        Exit For
                    End If
                Next X
            Next

        Catch ex As Exception
            LogEvent(ex.Message)
        Finally
            Try
                ReDim moEmailMessages(0)
            Catch ex As Exception
                LogEvent("Trying to reset message array: " & ex.Message)
            End Try
        End Try
    End Sub
End Class
