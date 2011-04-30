Option Strict On

Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Public Class StateObject
    Public workSocket As Socket = Nothing
    Public BufferSize As Int32 = 32767
    Public Buffer(32767) As Byte
    Public sb As New StringBuilder()
End Class

Public Class NetSock
    'This class is inherited by NetSockClient and NetSockServer
    Public Event onConnect(ByVal Index As Int32)
    Public Event onError(ByVal Index As Int32, ByVal Description As String)
    Public Event onDataArrival(ByVal Index As Int32, ByVal Data As Byte(), ByVal TotalBytes As Int32)
    Public Event onDisconnect(ByVal Index As Int32)
    Public Event onSendComplete(ByVal Index As Int32, ByVal DataSize As Int32)
    Public Event onConnectionRequest(ByVal Index As Int32, ByVal oClient As Socket)
    Public Event onDataContinued(ByVal lRemainingBytes As Int32)

    Public SocketIndex As Int32 = 0     'this is important for when doing arrays of these classes

    Public PortNumber As Int32
    Private moListenThread As Threading.Thread
    Private Const ml_LISTEN_CHECK_INTERVAL As Int32 = 2000      'milliseconds for sleeper
    Private mbListening As Boolean

    Private moSocket As Socket      'the actual socket containment

    Public Sub SetLocalBinding(ByVal iPort As Int16)

        '#If Debug = True Then
        '        Exit Sub
        '#Else
        If moSocket Is Nothing Then moSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

        Dim oLocalhost As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName)
        Dim oIP As IPAddress = oLocalhost.AddressList(0)

        moSocket.Bind(New System.Net.IPEndPoint(oIP, iPort))

        oIP = Nothing
        oLocalhost = Nothing
        '#End If

    End Sub

    Public Sub Connect(ByVal sHostIP As String, ByVal lPortNumber As Int32)
        'this will connect...
        Try
            Dim oEndPoint As IPEndPoint
            Dim oHost As IPHostEntry
            Dim oAddress As IPAddress

            If moSocket Is Nothing Then
                moSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            End If
            PortNumber = lPortNumber
            'oHost = Dns.GetHostByAddress(sHostIP)
            oHost = Dns.GetHostByName(sHostIP)
            'oHost = Dns.GetHostEntry(sHostIP)
            oAddress = oHost.AddressList(0)
            oEndPoint = New IPEndPoint(oAddress, PortNumber)

            moSocket.BeginConnect(oEndPoint, AddressOf ConnectedCallback, moSocket)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub SendData(ByVal Data() As Byte)
        Try
            'Dim iDataLen As Int16 = Data.Length + 1
            Dim yTemp(Data.Length + 3) As Byte
            System.BitConverter.GetBytes(CInt(Data.Length)).CopyTo(yTemp, 0)
            Data.CopyTo(yTemp, 4)
            'moSocket.BeginSend(Data, 0, Data.Length, 0, AddressOf SendDataCallback, moSocket)
            moSocket.BeginSend(yTemp, 0, yTemp.Length, 0, AddressOf SendDataCallback, moSocket)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub SendLenAppendedData(ByVal Data() As Byte)
        'Now, this sends the data... but we don't put the length in front of it... it is assumed that Data()
        '  already has the message length in front of it
        Try
            moSocket.BeginSend(Data, 0, Data.Length, 0, AddressOf SendDataCallback, moSocket)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub Disconnect()
        On Error Resume Next
        moSocket.Shutdown(SocketShutdown.Both)
        moSocket.Close()
    End Sub

    Private Sub ConnectedCallback(ByVal ar As IAsyncResult)
        Try
            If moSocket.Connected = False Then RaiseEvent onError(SocketIndex, "Connection Refused by Remote Host.") : Exit Sub

            'Dim sTemp As String
            'Dim iEP As IPEndPoint
            'iEP = moSocket.RemoteEndPoint
            'sTemp = "Connected Callback, Remote: " & iEP.Address.ToString() & ":" & iEP.Port & ", Local: "
            'iEP = moSocket.LocalEndPoint
            'sTemp &= iEP.Address.ToString() & ":" & iEP.Port & vbCrLf
            'Debug.Write(sTemp)

            Dim oState As New StateObject()
            moSocket.BeginReceive(oState.Buffer, 0, oState.BufferSize, 0, AddressOf DataArrivalCallback, oState)
            RaiseEvent onConnect(SocketIndex)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub DataArrivalCallback(ByVal ar As IAsyncResult)
        Dim oState As StateObject = CType(ar.AsyncState, StateObject)
        Dim yTemp() As Byte = Nothing
        Dim lBytesRead As Int32
        Dim lPos As Int32
        Dim lLen As Int32

        Static xyLeftover() As Byte
        Static xlLeftoverSize As Int32
        Static xlExpectedLen As Int32

        Try
            'Tell the socket to end our receive so we can work, thank you mr. socket for letting us know. have a nice day
            lBytesRead = moSocket.EndReceive(ar)

            'now, that we are here, get our data
            Dim yData() As Byte = oState.Buffer
            If lBytesRead = 0 Then
                moSocket.Shutdown(SocketShutdown.Both)
                moSocket.Close()
                RaiseEvent onDisconnect(SocketIndex)
                Exit Sub
            End If
            ReDim oState.Buffer(32767)

            lPos = 0
            While lPos < lBytesRead
                If xlExpectedLen <> 0 Then
                    'ok, gotta take care of carryover first
                    If xlExpectedLen > lBytesRead Then
                        ReDim yTemp(xyLeftover.Length + lBytesRead - 1)
                        xyLeftover.CopyTo(yTemp, 0)
                        Array.Copy(yData, 0, yTemp, xlLeftoverSize, lBytesRead)
                        lPos += lBytesRead
                        xlExpectedLen -= lBytesRead
                        xlLeftoverSize = xyLeftover.Length
                    Else
                        ReDim yTemp(xyLeftover.Length + xlExpectedLen - 1)
                        xyLeftover.CopyTo(yTemp, 0)
                        Array.Copy(yData, 0, yTemp, xlLeftoverSize, xlExpectedLen)
                        lPos += xlExpectedLen
                        xlExpectedLen = 0
                    End If
                Else
                    lLen = System.BitConverter.ToInt32(yData, lPos)
                    lPos += 4
                    If (lBytesRead - lPos) < lLen Then
                        'not enough bytes in this message, set up to wait for next message
                        xlLeftoverSize = lBytesRead - lPos
                        ReDim xyLeftover(xlLeftoverSize - 1)
                        Array.Copy(yData, lPos, xyLeftover, 0, xlLeftoverSize)
                        xlExpectedLen = lLen - xlLeftoverSize
                        lPos += xlLeftoverSize
                    Else
                        'nope, all is here, let's process like normal
                        ReDim yTemp(lLen)
                        Array.Copy(yData, lPos, yTemp, 0, lLen)
                        lPos += lLen
                    End If
                End If

                If xlExpectedLen = 0 Then RaiseEvent onDataArrival(SocketIndex, yTemp, lLen) Else RaiseEvent onDataContinued(xlExpectedLen)
            End While
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
        End Try

        'This new Try...Catch block may hurt performance...
        Try
            'Set up our socket to receive the next data arrival
            moSocket.BeginReceive(oState.Buffer, 0, oState.BufferSize, 0, AddressOf DataArrivalCallback, oState)
        Catch
            RaiseEvent onError(SocketIndex, "Socket Begin Receive Failed: " & Err.Description)
        End Try

    End Sub

    Private Sub SendDataCallback(ByVal ar As IAsyncResult)
        Try
            'Dim oTmpSock As Socket = CType(ar.AsyncState, Socket)
            Dim lBytesSent As Int32 = moSocket.EndSend(ar)
            RaiseEvent onSendComplete(SocketIndex, lBytesSent)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub Listen()
        'ok, here, we will start a new thread to handle listening for clients...
        mbListening = True
        moListenThread = New Threading.Thread(AddressOf BeginListening)
        moListenThread.Start()
    End Sub

    Public Sub StopListening()
        mbListening = False
    End Sub

    'NOTE: This sub should be called from Listen() as a new thread
    Private Sub BeginListening()
        'Create the listener socket
        'Dim oListener As New TcpListener(Dns.GetHostEntry("localhost").AddressList(0), PortNumber)
        Dim oListener As New TcpListener(PortNumber)

        'start listening
        oListener.Start()
        'Now, we will loop while we are listening
        While mbListening
            'Ok, we are listening... check our listener
            While oListener.Pending
                'Ok, we have pending connections...
                RaiseEvent onConnectionRequest(SocketIndex, oListener.AcceptSocket())
            End While

            'Now, that we have connected everyone, let's sleep a while
            Threading.Thread.Sleep(ml_LISTEN_CHECK_INTERVAL)
        End While
    End Sub

    Public Sub MakeReadyToReceive()
        Dim oState As New StateObject()
        moSocket.BeginReceive(oState.Buffer, 0, oState.BufferSize, 0, AddressOf DataArrivalCallback, oState)
    End Sub

    Public Sub New()
        'Ok, just your basic construct, no biggy
    End Sub

    Public Sub New(ByVal oSocket As Socket)
        'Ok, now we received a socket that the program already has...
        moSocket = oSocket
    End Sub

End Class

Public Class InitFile
    ' API functions
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
    Private strFilename As String

    ' Constructor, accepting a filename
    Public Sub New(Optional ByVal Filename As String = "")
        If Filename = "" Then
            'Ok, use the app.path
            strFilename = System.AppDomain.CurrentDomain.BaseDirectory()
            If Right$(strFilename, 1) <> "\" Then strFilename = strFilename & "\"
            strFilename = strFilename & Replace$(System.AppDomain.CurrentDomain.FriendlyName().ToLower, ".exe", ".ini")
        Else
            strFilename = Filename
        End If
    End Sub

    ' Read-only filename property
    ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property

    Public Function GetString(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As String) As String
        ' Returns a string from your INI file
        Dim intCharCount As Integer
        Dim objResult As New System.Text.StringBuilder(256)
        intCharCount = GetPrivateProfileString(Section, Key, _
           [Default], objResult, objResult.Capacity, strFilename)
        If intCharCount > 0 Then Return Left(objResult.ToString, intCharCount) Else Return ""
    End Function

    Public Function GetInteger(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Integer) As Integer
        ' Returns an integer from your INI file
        Return GetPrivateProfileInt(Section, Key, _
           [Default], strFilename)
    End Function

    Public Function GetBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Boolean) As Boolean
        ' Returns a boolean from your INI file
        Return (GetPrivateProfileInt(Section, Key, _
           CInt([Default]), strFilename) = 1)
    End Function

    Public Sub WriteString(ByVal Section As String, _
      ByVal Key As String, ByVal Value As String)
        ' Writes a string to your INI file
        WritePrivateProfileString(Section, Key, Value, strFilename)
        Flush()
    End Sub

    Public Sub WriteInteger(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Integer)
        ' Writes an integer to your INI file
        WriteString(Section, Key, CStr(Value))
        Flush()
    End Sub

    Public Sub WriteBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Boolean)
        ' Writes a boolean to your INI file
        WriteString(Section, Key, CStr(CInt(Value)))
        Flush()
    End Sub

    Private Sub Flush()
        ' Stores all the cached changes to your INI file
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub

End Class
Public Module GlobalVars
    Public Function Exists(ByVal sFilename As String) As Boolean
        If Len(Trim$(sFilename)) > 0 Then
            On Error Resume Next
            sFilename = Dir$(sFilename, FileAttribute.Archive Or FileAttribute.Directory Or FileAttribute.Hidden Or FileAttribute.Normal Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
            Return Err.Number = 0 And Len(sFilename) > 0
        Else
            Return False
        End If

    End Function
End Module