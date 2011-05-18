Imports System.Collections
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Diagnostics

Namespace Pop3
    ''' <summary>
    ''' DLM: Stores the From:, To:, Subject:, body and attachments
    ''' within an email. Binary attachments are Base64-decoded
    ''' </summary>

    Public Class Pop3Message
        Private m_client As Socket

        Private m_messageComponents As Pop3MessageComponents

        Private m_from As String
        Private m_to As String
        Private m_subject As String
        Private m_contentType As String
        Private m_body As String
        Private m_FullMessage As String

        Private m_isMultipart As Boolean = False

        Private m_multipartBoundary As String

        Private Const m_fromState As Integer = 0
        Private Const m_toState As Integer = 1
        Private Const m_subjectState As Integer = 2
        Private Const m_contentTypeState As Integer = 3
        Private Const m_notKnownState As Integer = -99
        Private Const m_endOfHeader As Integer = -98

        ' this array corresponds with above
        ' enumerator ...

        Private m_lineTypeString As String() = {"From", "To", "Subject", "Content-Type"}

        Private m_messageSize As Int32 = 0
        Private m_inboxPosition As Int32 = 0

        Private m_pop3State As Pop3StateObject = Nothing

        Private m_manualEvent As New ManualResetEvent(False)

        Public ReadOnly Property MultipartEnumerator() As IEnumerator
            Get
                Return m_messageComponents.ComponentEnumerator
            End Get
        End Property

        Public ReadOnly Property IsMultipart() As Boolean
            Get
                Return m_isMultipart
            End Get
        End Property

        Public ReadOnly Property From() As String
            Get
                Return m_from
            End Get
        End Property

        Public ReadOnly Property [To]() As String
            Get
                Return m_to
            End Get
        End Property

        Public ReadOnly Property Subject() As String
            Get
                Return m_subject
            End Get
        End Property

        Public ReadOnly Property Body() As String
            Get
                Return m_body
            End Get
        End Property

        Public ReadOnly Property FullMessage() As String
            Get
                Return m_FullMessage
            End Get
        End Property


        Public ReadOnly Property InboxPosition() As Int32
            Get
                Return m_inboxPosition
            End Get
        End Property

        'send the data to server
        Private Sub Send(ByVal data As [String])
            Try
                ' Convert the string data to byte data 
                ' using ASCII encoding.

                Dim byteData As Byte() = Encoding.ASCII.GetBytes(data & vbCr & vbLf)

                ' Begin sending the data to the remote device.
                m_client.Send(byteData)
            Catch e As Exception
                Throw New Pop3SendException(e.ToString())
            End Try
        End Sub

        Private Sub StartReceiveAgain(ByVal data As String)
            ' receive more data if we expect more.
            ' note: a literal "." (or more) followed by
            ' "\r\n" in an email is prefixed with "." ...

            If Not data.EndsWith(vbCr & vbLf & "." & vbCr & vbLf) Then
                m_client.BeginReceive(m_pop3State.buffer, 0, Pop3StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), m_pop3State)
            Else
                ' stop receiving data ...
                m_manualEvent.[Set]()
            End If
        End Sub

        Private Sub ReceiveCallback(ByVal ar As IAsyncResult)
            Try
                ' Retrieve the state object and the client socket 
                ' from the asynchronous state object.

                Dim stateObj As Pop3StateObject = DirectCast(ar.AsyncState, Pop3StateObject)

                Dim client As Socket = stateObj.workSocket

                ' Read data from the remote device.
                Dim bytesRead As Integer = client.EndReceive(ar)

                If bytesRead > 0 Then
                    ' There might be more data, 
                    ' so store the data received so far.

                    stateObj.sb.Append(Encoding.ASCII.GetString(stateObj.buffer, 0, bytesRead))

                    ' read more data from pop3 server ...
                    StartReceiveAgain(stateObj.sb.ToString())
                End If
            Catch e As Exception
                m_manualEvent.[Set]()

                Throw New Pop3ReceiveException("RecieveCallback error" & e.ToString())
            End Try
        End Sub

        Private Sub StartReceive()
            ' start receiving data ...
            m_client.BeginReceive(m_pop3State.buffer, 0, Pop3StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReceiveCallback), m_pop3State)

            ' wait until no more data to be read ...
            m_manualEvent.WaitOne()
        End Sub

        Private Function GetHeaderLineType(ByVal line As String) As Integer
            Dim lineType As Integer = m_notKnownState

            For i As Integer = 0 To m_lineTypeString.Length - 1
                Dim match As String = m_lineTypeString(i)

                If Regex.Match(line, "^" & match & ":" & ".*$").Success Then
                    lineType = i
                    Exit For
                ElseIf line.Length = 0 Then
                    lineType = m_endOfHeader
                    Exit For
                End If
            Next

            Return lineType
        End Function

        Private Function ParseHeader(ByVal lines As String()) As Int32
            Dim numberOfLines As Integer = lines.Length
            Dim bodyStart As Int32 = 0

            For i As Integer = 0 To numberOfLines - 1
                Dim currentLine As String = lines(i).Replace(vbLf, "")

                Dim lineType As Integer = GetHeaderLineType(currentLine)

                Select Case lineType
                    ' From:
                    Case m_fromState
                        m_from = Pop3Parse.From(currentLine)
                        Exit Select

                        ' Subject:
                    Case m_subjectState
                        m_subject = Pop3Parse.Subject(currentLine)
                        Exit Select

                        ' To:
                    Case m_toState
                        m_to = Pop3Parse.[To](currentLine)
                        Exit Select

                        ' Content-Type
                    Case m_contentTypeState

                        m_contentType = Pop3Parse.ContentType(currentLine)

                        m_isMultipart = Pop3Parse.IsMultipart(m_contentType)

                        If m_isMultipart Then
                            ' if boundary definition is on next
                            ' line ...

                            If m_contentType.Substring(m_contentType.Length - 1, 1).Equals(";") Then
                                i += 1

                                m_multipartBoundary = Pop3Parse.MultipartBoundary(lines(i).Replace(vbLf, ""))
                            Else
                                ' boundary definition is on same
                                ' line as "Content-Type" ...

                                m_multipartBoundary = Pop3Parse.MultipartBoundary(m_contentType)
                            End If
                        End If

                        Exit Select

                    Case m_endOfHeader
                        bodyStart = i + 1
                        Exit Select
                End Select

                If bodyStart > 0 Then
                    Exit For
                End If
            Next

            Return (bodyStart)
        End Function

        Private Sub ParseEmail(ByVal lines As String())
            Dim startOfBody As Int32 = ParseHeader(lines)
            Dim numberOfLines As Int32 = lines.Length

            m_messageComponents = New Pop3MessageComponents(lines, startOfBody, m_multipartBoundary, m_contentType)
        End Sub

        Private Sub LoadEmail()
            ' tell pop3 server we want to start reading
            ' email (m_inboxPosition) from inbox ...

            Send("retr " & m_inboxPosition)

            ' start receiving email ...
            StartReceive()


            m_FullMessage = m_pop3State.sb.ToString()

            ' parse email ...
            ParseEmail(m_pop3State.sb.ToString().Split(New Char() {ControlChars.Cr}))

            ' remove reading pop3State ...
            m_pop3State = Nothing
        End Sub

        Public Sub New(ByVal position As Int32, ByVal size As Int32, ByVal client As Socket)
            m_inboxPosition = position
            m_messageSize = size
            m_client = client

            m_pop3State = New Pop3StateObject()
            m_pop3State.workSocket = m_client
            m_pop3State.sb = New StringBuilder()

            ' load email ...
            LoadEmail()

            ' get body (if it exists) ...
            Dim multipartEnumerator__1 As IEnumerator = MultipartEnumerator

            While multipartEnumerator__1.MoveNext()
                Dim multipart As Pop3Component = DirectCast(multipartEnumerator__1.Current, Pop3Component)

                If multipart.IsBody Then
                    m_body = multipart.Data
                    Exit While
                End If
            End While
        End Sub

        Public Overrides Function ToString() As String
            Dim enumerator As IEnumerator = MultipartEnumerator

            Dim ret As String = "From    : " & m_from & vbCr & vbLf & "To      : " & m_to & vbCr & vbLf & "Subject : " & m_subject & vbCr & vbLf

            While enumerator.MoveNext()
                ret += DirectCast(enumerator.Current, Pop3Component).ToString() & vbCr & vbLf
            End While

            Return ret
        End Function
    End Class
End Namespace
