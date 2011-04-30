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
    Public Event ZeroReturnNeedRerequest()

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
        Dim sLoc As String = ""
        Dim lMethod As Int32 = 0
        Try
            Dim oEndPoint As IPEndPoint = Nothing
            Dim oHost As IPHostEntry
            Dim oAddress As IPAddress = Nothing

            If moSocket Is Nothing Then
                moSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            End If
            PortNumber = lPortNumber

            'oHost = Dns.GetHostEntry(sHostIP)
            'oAddress = oHost.AddressList(0)

            oHost = Nothing
            Dim bGood As Boolean = False

            '2010-02-08: RTP Added checker if we are forcably connecting to an IP address.  If so do NOT resolve it's Dns and then use the returned hostname e.g. www.beyondprotocol.com
            Dim pattern As String = "^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$"
            Dim check As New System.Text.RegularExpressions.Regex(pattern)
            If check.IsMatch(sHostIP, 0) Then
                '2010-02-10: MSC added code here to simply generate the endpoint outright and drop DNS from the equation altogether
                Dim oIP As System.Net.IPAddress = Nothing
                If System.Net.IPAddress.TryParse(sHostIP, oIP) = True Then
                    oEndPoint = New System.Net.IPEndPoint(oIP, lPortNumber)
                Else
                    'I couldn't find a good way to push sHostIp into oHost.  So cheating and using the wrong dns call.
                    oHost = Dns.GetHostByName(sHostIP)
                End If
                bGood = True
            Else
                Try
                    sLoc = "GetByEntry"
                    oHost = Dns.GetHostEntry(sHostIP)
                    If oHost Is Nothing = False AndAlso oHost.AddressList Is Nothing = False Then
                        For Each oTmpAddr As IPAddress In oHost.AddressList
                            If oTmpAddr.AddressFamily = AddressFamily.InterNetwork Then
                                bGood = True
                                oAddress = oTmpAddr
                                lMethod = 1
                                Exit For
                            End If
                        Next
                    End If
                Catch
                End Try
            End If

            If bGood = False Then
                Try
                    sLoc = "GetByName"
                    oHost = Dns.GetHostByName(sHostIP)
                    If oHost Is Nothing = False AndAlso oHost.AddressList Is Nothing = False Then
                        For Each oTmpAddr As IPAddress In oHost.AddressList
                            If oTmpAddr.AddressFamily = AddressFamily.InterNetwork Then
                                bGood = True
                                oAddress = oTmpAddr
                                lMethod = 2
                                Exit For
                            End If
                        Next
                    End If
                Catch
                End Try

                If bGood = False Then
                    sLoc = "GetByAddr"
                    Try
                        oHost = Dns.GetHostByAddress(sHostIP)
                    Catch
                    End Try

                    If oHost Is Nothing = False AndAlso oHost.AddressList Is Nothing = False Then
                        For Each oTmpAddr As IPAddress In oHost.AddressList
                            If oTmpAddr.AddressFamily = AddressFamily.InterNetwork Then
                                If IsNumeric(sHostIP.Replace(".", "")) = True Then
                                    If sHostIP = oTmpAddr.ToString Then
                                        bGood = True
                                        oAddress = oTmpAddr
                                        lMethod = 3
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            End If

            If oAddress Is Nothing AndAlso oEndPoint Is Nothing Then
                If oHost Is Nothing = False AndAlso oHost.AddressList Is Nothing = False AndAlso oHost.AddressList.Length > 0 Then
                    oAddress = oHost.AddressList(0)
                Else
                    RaiseEvent onError(SocketIndex, "Unable to resolve host. Check your internet connection." & vbCrLf & "If the problem persists, contact support@darkskyentertainment.com.")
                    Return
                End If
            End If

            If bGood = False Then
                RaiseEvent onError(SocketIndex, "Unable to resolve host. Check your internet connection." & vbCrLf & "If the problem persists, contact support@darkskyentertainment.com.")
                Return
            End If


            If oEndPoint Is Nothing Then oEndPoint = New IPEndPoint(oAddress, PortNumber)

            moSocket.BeginConnect(oEndPoint, AddressOf ConnectedCallback, moSocket)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description & " L: " & sLoc & ", M: " & lMethod)
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
                RaiseEvent ZeroReturnNeedRerequest()
                Return
                'moSocket.Shutdown(SocketShutdown.Both)
                'moSocket.Close()
                'RaiseEvent onDisconnect(SocketIndex)
                'Exit Sub
            End If
            ReDim oState.Buffer(32767)

            lPos = 0
            While lPos < lBytesRead
                If xlExpectedLen <> 0 Then
                    'ok, gotta take care of carryover first
                    If xlExpectedLen > lBytesRead Then
                        'ReDim Preserve xyLeftover(xyLeftover.Length + lBytesRead - 1)
                        Array.Copy(yData, 0, xyLeftover, xlLeftoverSize, lBytesRead)
                        lPos += lBytesRead
                        xlExpectedLen -= lBytesRead
                        xlLeftoverSize += lBytesRead
                    Else
                        'ReDim Preserve xyLeftover(xyLeftover.Length + xlExpectedLen - 1)
                        Array.Copy(yData, 0, xyLeftover, xlLeftoverSize, xlExpectedLen)
                        lPos += xlExpectedLen
                        xlExpectedLen = 0
                        yTemp = xyLeftover
                    End If
                Else
                    lLen = System.BitConverter.ToInt32(yData, lPos)
                    lPos += 4
                    If (lBytesRead - lPos) < lLen Then
                        'not enough bytes in this message, set up to wait for next message
                        xlLeftoverSize = lBytesRead - lPos
                        'ReDim xyLeftover(xlLeftoverSize - 1)
                        ReDim xyLeftover(lLen - 1) ' - 1)
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

	Public Function GetLocaleSpecificDT(ByVal sUSBased As String) As Date
		Dim lIdx As Int32 = sUSBased.IndexOf("/"c)
		Dim lIdx2 As Int32 = sUSBased.IndexOf("/"c, lIdx + 1)
		Dim sMM As String = sUSBased.Substring(0, lIdx)
		Dim sDD As String = sUSBased.Substring(lIdx + 1, lIdx2 - lIdx - 1)
		Dim sYYYY As String = sUSBased.Substring(lIdx2 + 1, 4)
		Dim sTime As String = sUSBased.Substring(lIdx2 + 5).Trim

		Try
			Dim dtDate As Date = CDate("January 5, 2012")
			Dim sDate As String = dtDate.ToShortDateString.Replace(".", "/").Replace("-", "/")
			Dim lVal As Int32 = CInt(Val(sDate))
			If lVal = 1 Then
				'MM
				'now what?
				lIdx = sDate.IndexOf("/"c)
				If lIdx <> -1 Then sDate = sDate.Substring(lIdx + 1) Else sDate = sDate.Substring(2)
				lVal = CInt(Val(sDate))
				If lVal = 5 Then
					'DD
					'so, MM/DD/YYYY
					Return CDate(sMM & "/" & sDD & "/" & sYYYY & " " & sTime)
				Else
					'YYYY
					'so, MM/YYYY/DD
					Return CDate(sMM & "/" & sYYYY & "/" & sDD & " " & sTime)
				End If
			ElseIf lVal = 5 Then
				'DD
				lIdx = sDate.IndexOf("/"c)
				If lIdx <> -1 Then sDate = sDate.Substring(lIdx + 1) Else sDate = sDate.Substring(3)
				lVal = CInt(Val(sDate))
				If lVal = 1 Then
					'MM
					'so, DD/MM/YYYY
					Return CDate(sDD & "/" & sMM & "/" & sYYYY & " " & sTime)
				Else
					'YYYY
					Return CDate(sDD & "/" & sYYYY & "/" & sMM & " " & sTime)
				End If
			Else
				'year
				lIdx = sDate.IndexOf("/"c)
				If lIdx <> -1 Then sDate = sDate.Substring(lIdx + 1) Else sDate = sDate.Substring(5)
				lVal = CInt(Val(sDate))
				If lVal = 1 Then
					'MM
					'so, YYYY/MM/DD
					Return CDate(sYYYY & "/" & sMM & "/" & sDD & " " & sTime)
				Else
					'DD
					'so, YYYY/DD/MM
					Return CDate(sYYYY & "/" & sDD & "/" & sMM & " " & sTime)
				End If
			End If
		Catch ex As Exception
			frmMain.AddTextLine("GetLocaleSpecificDT: " & ex.Message)
			Try
				Return CDate(sUSBased)
			Catch
			End Try
		End Try
	End Function

	Public Sub SetFileDateTime(ByVal sDate As String, ByVal sFilename As String)
		Dim tMyDate As Date

		'now update the datetime stamp
		'2/17/2003 11:56:41 AM
		tMyDate = GetLocaleSpecificDT(sDate)

		Dim oFSInfo As IO.FileInfo = New IO.FileInfo(sFilename)
		oFSInfo.CreationTime = tMyDate
		oFSInfo.LastAccessTime = tMyDate
		oFSInfo.LastWriteTime = tMyDate
		oFSInfo = Nothing

	End Sub

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