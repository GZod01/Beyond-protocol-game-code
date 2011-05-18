Imports System.Collections
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Diagnostics

Namespace Pop3
    Public Class Pop3Client
        Private m_credential As Pop3Credential

        Private Const m_pop3port As Integer = 110
        Private Const MAX_BUFFER_READ_SIZE As Integer = 256

        Private m_inboxPosition As Int32 = 0
        Private m_directPosition As Int32 = -1

        Private m_socket As Socket = Nothing

        Private m_pop3Message As Pop3Message = Nothing

        Public Property UserDetails() As Pop3Credential
            Get
                Return m_credential
            End Get
            Set(ByVal value As Pop3Credential)
                m_credential = value
            End Set
        End Property

        Public ReadOnly Property From() As String
            Get
                Return m_pop3Message.From
            End Get
        End Property

        Public ReadOnly Property [To]() As String
            Get
                Return m_pop3Message.[To]
            End Get
        End Property

        Public ReadOnly Property Subject() As String
            Get
                Return m_pop3Message.Subject
            End Get
        End Property

        Public ReadOnly Property Body() As String
            Get
                Return m_pop3Message.Body
            End Get
        End Property

        Public ReadOnly Property FullMessage() As String
            Get
                Return m_pop3Message.FullMessage
            End Get
        End Property

        Public ReadOnly Property MultipartEnumerator() As IEnumerator
            Get
                Return m_pop3Message.MultipartEnumerator
            End Get
        End Property

        Public ReadOnly Property IsMultipart() As Boolean
            Get
                Return m_pop3Message.IsMultipart
            End Get
        End Property


        Public Sub New(ByVal user As String, ByVal pass As String, ByVal server As String)
            m_credential = New Pop3Credential(user, pass, server)
        End Sub

        Private Function GetClientSocket() As Socket
            Dim s As Socket = Nothing

            Try
                Dim hostEntry As IPHostEntry = Nothing

                ' Get host related information.
                hostEntry = Dns.Resolve(m_credential.Server)

                ' Loop through the AddressList to obtain the supported 
                ' AddressFamily. This is to avoid an exception that 
                ' occurs when the host IP Address is not compatible 
                ' with the address family 
                ' (typical in the IPv6 case).

                For Each address As IPAddress In hostEntry.AddressList
                    Dim ipe As New IPEndPoint(address, m_pop3port)

                    Dim tempSocket As New Socket(ipe.AddressFamily, SocketType.Stream, ProtocolType.Tcp)

                    tempSocket.Connect(ipe)

                    If tempSocket.Connected Then
                        ' we have a connection.
                        ' return this socket ...
                        s = tempSocket
                        Exit For
                    Else
                        Continue For
                    End If
                Next
            Catch e As Exception
                Throw New Pop3ConnectException(e.ToString())
            End Try

            ' throw exception if can't connect ...
            If s Is Nothing Then
                Throw New Pop3ConnectException("Error : connecting to " & Convert.ToString(m_credential.Server))
            End If

            Return s
        End Function

        'send the data to server
        Private Sub Send(ByVal data As [String])
            If m_socket Is Nothing Then
                Throw New Pop3MessageException("Pop3 connection is closed")
            End If

            Try
                ' Convert the string data to byte data 
                ' using ASCII encoding.

                Dim byteData As Byte() = Encoding.ASCII.GetBytes(data & vbCr & vbLf)

                ' Begin sending the data to the remote device.
                m_socket.Send(byteData)
            Catch e As Exception
                Throw New Pop3SendException(e.ToString())
            End Try
        End Sub

        Private Function GetPop3String() As String
            If m_socket Is Nothing Then
                Throw New Pop3MessageException("Connection to POP3 server is closed")
            End If

            Dim buffer As Byte() = New Byte(MAX_BUFFER_READ_SIZE - 1) {}
            Dim line As String = Nothing

            Try
                Dim byteCount As Integer = m_socket.Receive(buffer, buffer.Length, 0)

                line = Encoding.ASCII.GetString(buffer, 0, byteCount)
            Catch e As Exception
                Throw New Pop3ReceiveException(e.ToString())
            End Try

            Return line
        End Function

        Private Sub LoginToInbox()
            Dim returned As String

            ' send username ...
            Send("user " & Convert.ToString(m_credential.User))

            ' get response ...
            returned = GetPop3String()

            If Not returned.Substring(0, 3).Equals("+OK") Then
                Throw New Pop3LoginException("login not excepted")
            End If

            ' send password ...
            Send("pass " & Convert.ToString(m_credential.Pass))

            ' get response ...
            returned = GetPop3String()

            If Not returned.Substring(0, 3).Equals("+OK") Then
                Throw New Pop3LoginException("login/password not accepted")
            End If
        End Sub

        Public ReadOnly Property MessageCount() As Int32
            Get
                Dim count As Int32 = 0

                If m_socket Is Nothing Then
                    Throw New Pop3MessageException("Pop3 server not connected")
                End If

                Send("stat")

                Dim returned As String = GetPop3String()

                ' if values returned ...
                If Regex.Match(returned, "^.*\+OK[ |" & vbTab & "]+([0-9]+)[ |" & vbTab & "]+.*$").Success Then
                    ' get number of emails ...
                    count = Int32.Parse(Regex.Replace(returned.Replace(vbCr & vbLf, ""), "^.*\+OK[ |" & vbTab & "]+([0-9]+)[ |" & vbTab & "]+.*$", "$1"))
                End If

                Return (count)
            End Get
        End Property


        Public Sub CloseConnection()
            Send("quit")

            m_socket = Nothing
            m_pop3Message = Nothing
        End Sub

        Public Function DeleteEmail() As Boolean
            Dim ret As Boolean = False

            Send("dele " & m_inboxPosition)

            Dim returned As String = GetPop3String()

            If Regex.Match(returned, "^.*\+OK.*$").Success Then
                ret = True
            End If

            Return ret
        End Function

        Public Function NextEmail(ByVal directPosition As Int32) As Boolean
            Dim ret As Boolean

            If directPosition >= 0 Then
                m_directPosition = directPosition
                ret = NextEmail()
            Else
                Throw New Pop3MessageException("Position less than zero")
            End If

            Return ret
        End Function

        Public Function NextEmail() As Boolean
            Dim returned As String

            Dim pos As Int32

            If m_directPosition = -1 Then
                If m_inboxPosition = 0 Then
                    pos = 1
                Else
                    pos = m_inboxPosition + 1
                End If
            Else
                pos = m_directPosition + 1
                m_directPosition = -1
            End If

            ' send username ...
            Send("list " & pos.ToString())

            ' get response ...
            returned = GetPop3String()

            ' if email does not exist at this position
            ' then return false ...

            If returned.Substring(0, 4).Equals("-ERR") Then
                Return False
            End If

            m_inboxPosition = pos

            ' strip out CRLF ...
            Dim noCr As String() = returned.Split(New Char() {ControlChars.Cr})

            ' get size ...
            Dim elements As String() = noCr(0).Split(New Char() {" "c})

            Dim size As Int32 = Int32.Parse(elements(2))

            ' ... else read email data
            m_pop3Message = New Pop3Message(m_inboxPosition, size, m_socket)

            Return True
        End Function

        Public Sub OpenInbox()
            ' get a socket ...
            m_socket = GetClientSocket()

            ' get initial header from POP3 server ...
            Dim header As String = GetPop3String()

            If Not header.Substring(0, 3).Equals("+OK") Then
                Throw New Exception("Invalid initial POP3 response")
            End If

            ' send login details ...
            LoginToInbox()
        End Sub
    End Class
End Namespace
