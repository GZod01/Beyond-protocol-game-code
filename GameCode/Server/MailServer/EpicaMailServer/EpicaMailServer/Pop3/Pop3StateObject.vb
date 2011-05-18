Imports System.Collections
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Namespace Pop3
    ''' <summary>
    ''' Holds the current state of the client
    ''' socket.
    ''' </summary>

    Public Class Pop3StateObject
        ' Client socket.
        Public workSocket As Socket = Nothing

        ' Size of receive buffer.
        Public Const BufferSize As Integer = 256

        ' Receive buffer.
        Public buffer As Byte() = New Byte(BufferSize - 1) {}

        ' Received data string.
        Public sb As New StringBuilder()
    End Class
End Namespace
