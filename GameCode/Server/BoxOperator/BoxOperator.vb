Option Strict On

Imports System.ServiceProcess

Public Enum ConnectionType As Int32
    eUnknown = -1
    eNoConnection = 0
    eBackupOperator = 1
    eClientConnection = 2
    ePrimaryServerApp = 3
    eRegionServerApp = 4
    ePathfindingServerApp = 5
    eEmailServerApp = 6
    eBoxOperator = 7
    eOperator = 8
End Enum

Public Class BoxOperator
    Inherits System.ServiceProcess.ServiceBase

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        ' This call is required by the Component Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call

    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' The main entry point for the process
    <MTAThread()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New BoxOperator}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        Me.ServiceName = "BoxOperator"
    End Sub

#End Region

    Private Const gl_OPERATOR_DIE_TIME As Int32 = 20000

    Private WithEvents moOperator As NetSock       'connection to operator

    Private mlBoxOperatorID As Int32 = -1

    Private mbActive As Boolean = False
    Private mlTimeout As Int32 = 10000
    Private mbConnectingToOperator As Boolean = False
    Private mbConnectedToOperator As Boolean = False

    Private msBackupOperator As String = ""
    Private msPathfinding As String = ""
    Private msPrimary As String = ""
    Private msRegion As String = ""
    Private msEmailServer As String = ""

    Private msOperatorIP As String = ""
	Private mlOperatorPort As Int32 = 0
	Private msBackupOperatorIP As String = ""
    Private mlBackupOperatorPort As Int32 = 0

    Public sExternalIPAddress As String = ""

	Private moProcesses() As Process = Nothing
	Private mlProcType() As ConnectionType = Nothing

	Private moKeepAliveSW As Stopwatch
	Private mbInConnectToBackup As Boolean = False

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        InitializeEventLogging()
		LoadProcessStrings()

		mbActive = True
        Dim oThread As Threading.Thread = New Threading.Thread(AddressOf AttemptConnectOperator)
		oThread.Start()

		Dim oMainThread As Threading.Thread = New Threading.Thread(AddressOf MainLoop)
		oMainThread.Start()
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        mbActive = False
        LogEvent(LogEventType.Informational, "Service closing")
        CloseEventLogger()
    End Sub

	Private Sub AttemptConnectOperator()
		Dim oEP As Net.EndPoint = Nothing
		Dim bCheckedIfIAmOperator As Boolean = False

		While mbActive = True AndAlso mbConnectedToOperator = False
			If mlBoxOperatorID = -1 Then
				If ConnectToOperator() = False Then
					Threading.Thread.Sleep(10000)

					If moOperator Is Nothing = False AndAlso bCheckedIfIAmOperator = False Then
						oEP = moOperator.GetLocalDetails()
						If oEP Is Nothing = False Then
							'ok, check our address
							bCheckedIfIAmOperator = True
							Dim sMyIP As String = CType(oEP, Net.IPEndPoint).Address.ToString
							If sMyIP = msOperatorIP Then
								'I am the operator box... check connected
								If ProcessRunning(ConnectionType.eOperator) = True Then
									Dim lResult As Int32 = SpawnProcess(ConnectionType.eOperator)
									If lResult = Int32.MinValue Then
										LogEvent(LogEventType.CriticalError, "Spawn Operator Failed")
									End If
								End If
							End If
                            If sMyIP = msBackupOperatorIP Then
                                If ProcessRunning(ConnectionType.eBackupOperator) = True Then
                                    Dim lResult As Int32 = SpawnProcess(ConnectionType.eBackupOperator)
                                    If lResult = Int32.MinValue Then
                                        LogEvent(LogEventType.CriticalError, "Spawn Backup Operator Failed")
                                    End If
                                End If
                            End If
						End If
					End If
				End If
			End If
		End While

		moKeepAliveSW = Stopwatch.StartNew

		'Ok, if we're here, we've connected
		'    Indicate IP and Port information as well as our type
		oEP = moOperator.GetLocalDetails()
		If oEP Is Nothing = False Then
			With CType(oEP, Net.IPEndPoint)
                Dim yMsg(45) As Byte
				Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eBoxOperator).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(Process.GetCurrentProcess.Id).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4     'indicates 1 connection specific
                StringToBytes(Mid$(.Address.ToString(), 1, 20)).CopyTo(yMsg, lPos) : lPos += 20     'use my local ip address

                Dim oIni As New InitFile()
                Dim lPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "OperatorListenerPort", "0")))
                If lPort = 0 Then Err.Raise(-1, "AttemptConnectOperator: Send Server Details", "Could not get Operator Listen Port Number from INI")
                System.BitConverter.GetBytes(lPort).CopyTo(yMsg, lPos) : lPos += 4
                oIni = Nothing
                System.BitConverter.GetBytes(ConnectionType.eOperator).CopyTo(yMsg, lPos) : lPos += 4

                moOperator.SendData(yMsg)
            End With
        End If
    End Sub

    Private Sub AttemptConnectBackupOperator()
        While mbActive = True AndAlso mbConnectedToOperator = False
            If ConnectToBackupOperator() = False Then
                Threading.Thread.Sleep(10000)
            End If
        End While

        moKeepAliveSW = Stopwatch.StartNew

        'Ok, if we're here, we've connected
        '    Indicate IP and Port information as well as our type
        Dim oEP As Net.EndPoint = moOperator.GetLocalDetails()
        If oEP Is Nothing = False Then
            With CType(oEP, Net.IPEndPoint)
                Dim yMsg(45) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(mlBoxOperatorID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eBoxOperator).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(Process.GetCurrentProcess.Id).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4     'indicates 1 connection specific
                StringToBytes(Mid$(.Address.ToString(), 1, 20)).CopyTo(yMsg, lPos) : lPos += 20     'use my local ip address

                Dim oIni As New InitFile()
                Dim lPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "OperatorListenerPort", "0")))
                If lPort = 0 Then Err.Raise(-1, "AttemptConnectOperator: Send Server Details", "Could not get Operator Listen Port Number from INI")
                System.BitConverter.GetBytes(lPort).CopyTo(yMsg, lPos) : lPos += 4
                oIni = Nothing
                System.BitConverter.GetBytes(ConnectionType.eOperator).CopyTo(yMsg, lPos) : lPos += 4

                moOperator.SendData(yMsg)
            End With
        End If

        'HandleRequestServerStatus(Nothing)

        mbInConnectToBackup = False
    End Sub

    Private Sub LoadProcessStrings()
        Dim oINI As New InitFile()
        Dim sProcessBaseDirectory As String = oINI.GetString("SETTINGS", "ProcessDirectory", AppDomain.CurrentDomain.BaseDirectory)
        If sProcessBaseDirectory.EndsWith("\") = False Then sProcessBaseDirectory &= "\"

        mlTimeout = CInt(Val(oINI.GetString("CONNECTION", "ConnectTimeout", "10000")))
        mlOperatorPort = CInt(Val(oINI.GetString("CONNECTION", "OperatorPort", "0")))
        msOperatorIP = oINI.GetString("CONNECTION", "OperatorIP", "")
        mlBackupOperatorPort = CInt(Val(oINI.GetString("CONNECTION", "BackupOperatorPort", "0")))
        msBackupOperatorIP = oINI.GetString("CONNECTION", "BackupOperatorIP", "")

        msBackupOperator = oINI.GetString("SETTINGS", "BackupOperator", "Operator\Operator.exe 1")
        msPathfinding = oINI.GetString("SETTINGS", "Pathfinding", "Pathfinding\EpicaPathfinding.exe")
        msPrimary = oINI.GetString("SETTINGS", "Primary", "Primary\EpicaPrimary.exe")
        msRegion = oINI.GetString("SETTINGS", "Region", "Region\RegionServer.exe")
        msEmailServer = oINI.GetString("SETTINGS", "MailServer", "MailServer\EpicaMailServer.exe")

        sExternalIPAddress = oINI.GetString("SETTINGS", "ExternalIP", "EXTERNALIPADDY")

        If msBackupOperator <> "" Then msBackupOperator = sProcessBaseDirectory & msBackupOperator
        If msPathfinding <> "" Then msPathfinding = sProcessBaseDirectory & msPathfinding
        If msPrimary <> "" Then msPrimary = sProcessBaseDirectory & msPrimary
        If msRegion <> "" Then msRegion = sProcessBaseDirectory & msRegion
        If msEmailServer <> "" Then msEmailServer = sProcessBaseDirectory & msEmailServer

        oINI = Nothing
    End Sub

    Private Sub AddProcess(ByRef oProcess As Process, ByVal lType As ConnectionType)
        If moProcesses Is Nothing Then ReDim moProcesses(-1)
        SyncLock moProcesses
            ReDim Preserve moProcesses(moProcesses.GetUpperBound(0) + 1)
            moProcesses(moProcesses.GetUpperBound(0)) = oProcess
            ReDim Preserve mlProcType(moProcesses.GetUpperBound(0))
            mlProcType(mlProcType.GetUpperBound(0)) = lType
        End SyncLock
    End Sub

    Private Function ProcessRunning(ByVal lType As ConnectionType) As Boolean
        If moProcesses Is Nothing Then Return False
        If mlProcType Is Nothing Then Return False

        Dim lUB As Int32 = Math.Min(moProcesses.GetUpperBound(0), mlProcType.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If mlProcType(X) = lType Then
                If moProcesses(X) Is Nothing = False Then
                    If moProcesses(X).Responding = True Then Return True
                End If
            End If
        Next X
        Return False
    End Function

    '#Region "Operator Listener"
    '	Private WithEvents moListener As NetSock	   'listens for the operator
    '	Private Sub moListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moListener.onConnectionRequest
    '		LogEvent(LogEventType.Informational, "Operator Server Connecting...")
    '		moOperator = New NetSock(oClient)
    '		moOperator.MakeReadyToReceive()
    '		moOperator.lSpecificID = -1
    '	End Sub
    '#End Region

#Region "Operator Server Handling"
    Private Sub moOperator_onConnect(ByVal Index As Integer) Handles moOperator.onConnect
        mbConnectedToOperator = True
        mbConnectingToOperator = False
    End Sub

    Private Sub moOperator_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moOperator.onConnectionRequest

    End Sub

    Private Sub moOperator_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moOperator.onDataArrival
        Dim iMsgID As Int16

        iMsgID = System.BitConverter.ToInt16(Data, 0)
        Select Case iMsgID
            Case GlobalMessageCode.eRegisterDomain
                HandleAcquireBoxOperatorID(Data)
            Case GlobalMessageCode.eAddObjectCommand
                HandleSpawnProcess(Data)
            Case GlobalMessageCode.eRequestObject
                HandleRequestServerStatus(Data)
            Case GlobalMessageCode.eBackupOperatorSyncMsg
                HandleBackupOperatorSyncMsg(Data)
        End Select
    End Sub

    Private Sub moOperator_onDisconnect(ByVal Index As Integer) Handles moOperator.onDisconnect
        LogEvent(LogEventType.Informational, "Operator Server Disconnected")
    End Sub

    Private Sub moOperator_onError(ByVal Index As Integer, ByVal Description As String) Handles moOperator.onError
        LogEvent(LogEventType.Informational, "Operator Server Error: " & Description)
    End Sub

    Private Sub moOperator_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moOperator.onSendComplete
        '
    End Sub
#End Region

#Region " Event Logging "
    Private moFileStream As System.IO.FileStream
    Private moFileWrite As System.IO.StreamWriter
    Public Enum LogEventType
        CriticalError = 1
        PossibleCheat = 2
        Warning = 3
        Informational = 4
    End Enum

    Public Function InitializeEventLogging() As Boolean
        Dim oINI As New InitFile()
        Dim sLog As String

        sLog = oINI.GetString("SETTINGS", "EventLogFilePath", "")
        If sLog = "" Then
            sLog = AppDomain.CurrentDomain.BaseDirectory()
            If Right$(sLog, 1) <> "\" Then sLog = sLog & "\"
            sLog = sLog & "Events.Log"
        End If
        moFileStream = System.IO.File.Open(sLog, IO.FileMode.Append)
        moFileWrite = New System.IO.StreamWriter(moFileStream)
        moFileWrite.AutoFlush = True
    End Function

    Public Sub LogEvent(ByVal lEventType As LogEventType, ByVal sValue As String)

        'Public Shared Function StringToBytes(ByVal sData As String) As Byte()
        '    Return System.Text.ASCIIEncoding.ASCII.GetBytes(sData)
        'End Function
        Select Case lEventType
            Case LogEventType.CriticalError
                sValue = "CRITICAL: " & sValue
            Case LogEventType.PossibleCheat
                sValue = "POSSIBLE CHEAT: " & sValue
            Case LogEventType.Warning
                sValue = "WARNING: " & sValue
        End Select

        If moFileStream Is Nothing = True OrElse moFileWrite Is Nothing = True Then InitializeEventLogging()

        If moFileStream Is Nothing = False Then
            If moFileStream.CanWrite() Then
                moFileWrite.WriteLine(sValue)
            End If
        End If
    End Sub

    Public Sub CloseEventLogger()
        moFileWrite.Close()
        moFileStream.Close()
    End Sub
#End Region

    Private Sub HandleBackupOperatorSyncMsg(ByVal yData() As Byte)
        'ok, backup op is telling us that the primary op has been brought up... let's reconnect to it
        While mbInConnectToBackup = True
            Threading.Thread.Sleep(10)
        End While

        'Ok,now,
        moKeepAliveSW = Nothing
        mbConnectedToOperator = False
        If moOperator Is Nothing = False Then moOperator.Disconnect()

        Dim lAttempt As Int32 = 0
        While mbActive = True AndAlso mbConnectedToOperator = False AndAlso lAttempt < 3
            lAttempt += 1
            If ConnectToOperator() = False Then
                Threading.Thread.Sleep(10000)
            End If
        End While

        moKeepAliveSW = Stopwatch.StartNew
    End Sub

    Private Sub HandleAcquireBoxOperatorID(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'msgcode
        mlBoxOperatorID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        lPos = 0
        'Send our specs for this box
        Dim yOutMsg(45) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eEnvironmentDomain).CopyTo(yOutMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlBoxOperatorID).CopyTo(yOutMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(ConnectionType.eBoxOperator).CopyTo(yOutMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(System.Environment.ProcessorCount).CopyTo(yOutMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(My.Computer.Info.TotalPhysicalMemory).CopyTo(yOutMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(My.Computer.Info.AvailablePhysicalMemory).CopyTo(yOutMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(My.Computer.Info.TotalVirtualMemory).CopyTo(yOutMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(My.Computer.Info.AvailableVirtualMemory).CopyTo(yOutMsg, lPos) : lPos += 8
        moOperator.SendData(yOutMsg)

        ''Now, disconnect from the operator
        'moOperator.Disconnect()

        '      'and set us up as a listener
        '      Dim oINI As New InitFile()
        '      Try
        '          Dim lPort As Int32 = CInt(Val(oINI.GetString("SETTINGS", "OperatorListenerPort", "0")))
        '          If lPort = 0 Then Err.Raise(-1, "HandleAcquireBoxOperatorID: Begin Listener", "Could not get Operator Listen Port Number from INI")
        '          If moListener Is Nothing Then moListener = New NetSock()
        '          moListener.PortNumber = lPort
        '          moListener.Listen()
        '      Catch
        '          LogEvent(LogEventType.CriticalError, "An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description)
        '      Finally
        '          oINI = Nothing
        '      End Try
    End Sub

    Private Sub HandleSpawnProcess(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lProcessType As ConnectionType = CType(System.BitConverter.ToInt32(yData, lPos), ConnectionType) : lPos += 4
        Dim lSpawnRequestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        LogEvent(LogEventType.Informational, "SpawnProcess Received: " & lProcessType.ToString)

        Dim lResult As Int32 = SpawnProcess(lProcessType)

        Dim yResp(13) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yResp, 0)
        System.BitConverter.GetBytes(lProcessType).CopyTo(yResp, 2)
        System.BitConverter.GetBytes(lSpawnRequestID).CopyTo(yResp, 6)
        System.BitConverter.GetBytes(lResult).CopyTo(yResp, 10)
        moOperator.SendData(yResp)
    End Sub

    Private Function ConnectToOperator() As Boolean
        Dim oINI As InitFile = New InitFile()
        Dim bRes As Boolean = False

        Try
            If mlOperatorPort = 0 Or msOperatorIP = "" Then
                Err.Raise(-1, "EstablishConnection", "Unable to connect to server: could not find connection credentials.")
            End If
            mbConnectingToOperator = True
            moOperator = New NetSock()
            moOperator.Connect(msOperatorIP, mlOperatorPort)
            Dim sw As Stopwatch = Stopwatch.StartNew
            While mbConnectedToOperator = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToOperator = True
                Threading.Thread.Sleep(10)
            End While
            sw.Stop()
            sw = Nothing
            mbConnectingToOperator = False
            If mbConnectedToOperator = False Then
                Err.Raise(-1, "EstablishConnection", "Unable to connect to server: Connection timed out!")
            End If
        Catch
            mbConnectedToOperator = False
        Finally
            bRes = mbConnectedToOperator
            oINI = Nothing
        End Try
        Return bRes
    End Function
    Private Function ConnectToBackupOperator() As Boolean
        Dim oINI As InitFile = New InitFile()
        Dim bRes As Boolean = False

        Try
            If mlBackupOperatorPort = 0 Or msBackupOperatorIP = "" Then
                Err.Raise(-1, "EstablishConnection", "Unable to connect to server: could not find backup operator connection credentials.")
            End If

            mbConnectingToOperator = True
            moOperator = Nothing
            moOperator = New NetSock()
            moOperator.Connect(msBackupOperatorIP, mlBackupOperatorPort)
            Dim sw As Stopwatch = Stopwatch.StartNew
            While mbConnectedToOperator = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToOperator = True
                Threading.Thread.Sleep(10)
            End While
            sw.Stop()
            sw = Nothing
            mbConnectingToOperator = False
            If mbConnectedToOperator = False Then
                Err.Raise(-1, "EstablishConnection", "Unable to connect to server: Connection timed out!")
            End If
        Catch
            mbConnectedToOperator = False
        Finally
            bRes = mbConnectedToOperator
            oINI = Nothing
        End Try
        Return bRes
    End Function


    Private Function BytesToString(ByVal yBytes() As Byte) As String
        Dim lLen As Int32
        Dim X As Int32

        lLen = yBytes.Length
        For X = 0 To yBytes.Length - 1
            If yBytes(X) = 0 Then
                lLen = X
                Exit For
            End If
        Next X
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yBytes), 1, lLen)
    End Function

    Private Function StringToBytes(ByVal sVal As String) As Byte()
        Return System.Text.ASCIIEncoding.ASCII.GetBytes(sVal)
    End Function

    Private Sub HandleRequestServerStatus(ByRef yData() As Byte)
        Dim yResp() As Byte

        Dim lPos As Int32 = 0
        Dim lUB As Int32 = -1

        If moProcesses Is Nothing = False Then lUB = moProcesses.GetUpperBound(0)

        If moKeepAliveSW Is Nothing = False Then
            moKeepAliveSW.Reset()
            moKeepAliveSW.Start()
        End If

        ReDim yResp(5 + (122 * (lUB + 1)))

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(lUB + 1).CopyTo(yResp, lPos) : lPos += 4

        For X As Int32 = 0 To lUB
            System.BitConverter.GetBytes(mlProcType(X)).CopyTo(yResp, lPos) : lPos += 4
            With moProcesses(X)
                System.BitConverter.GetBytes(.Id).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.BasePriority).CopyTo(yResp, lPos) : lPos += 4

                If .Responding = True Then yResp(lPos) = 1 Else yResp(lPos) = 0
                lPos += 1

                If .HasExited = True Then yResp(lPos) = 1 Else yResp(lPos) = 0
                lPos += 1

                Dim dSecondsRunning As Double = Now.Subtract(.StartTime).TotalSeconds
                System.BitConverter.GetBytes(dSecondsRunning).CopyTo(yResp, lPos) : lPos += 8

                System.BitConverter.GetBytes(.TotalProcessorTime.TotalSeconds).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.UserProcessorTime.TotalSeconds).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.PrivilegedProcessorTime.TotalSeconds).CopyTo(yResp, lPos) : lPos += 8

                System.BitConverter.GetBytes(.NonpagedSystemMemorySize64).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.PagedSystemMemorySize64).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.PagedMemorySize64).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.PrivateMemorySize64).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.VirtualMemorySize64).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.WorkingSet64).CopyTo(yResp, lPos) : lPos += 8

                System.BitConverter.GetBytes(.PeakPagedMemorySize64).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.PeakVirtualMemorySize64).CopyTo(yResp, lPos) : lPos += 8
                System.BitConverter.GetBytes(.PeakWorkingSet64).CopyTo(yResp, lPos) : lPos += 8

                System.BitConverter.GetBytes(.Threads.Count).CopyTo(yResp, lPos) : lPos += 4
            End With
        Next X

        moOperator.SendData(yResp)

    End Sub

	Private Sub MainLoop()
		While mbActive = True
			If mbConnectedToOperator = True Then
				If moKeepAliveSW Is Nothing = False Then
					If moKeepAliveSW.ElapsedMilliseconds > gl_OPERATOR_DIE_TIME Then
						'Ok, need to connect to backup operator
						mbConnectedToOperator = False
						mbConnectingToOperator = False
						moKeepAliveSW = Nothing

						If mbInConnectToBackup = False Then
							mbInConnectToBackup = True
							Dim oThread As Threading.Thread = New Threading.Thread(AddressOf AttemptConnectBackupOperator)
							oThread.Start()
						End If

					End If
				End If
			End If

			Threading.Thread.Sleep(1000)
		End While
	End Sub

	Private Function SpawnProcess(ByVal lProcessType As ConnectionType) As Int32
		Dim sFile As String = ""
		Dim lResult As Int32 = Int32.MinValue

		Select Case lProcessType
			Case ConnectionType.eBackupOperator, ConnectionType.eOperator
				'Spawn the Backup Operator application
				sFile = msBackupOperator
			Case ConnectionType.eEmailServerApp
				'Spawn Email Server App
				sFile = msEmailServer
			Case ConnectionType.ePathfindingServerApp
				'Spawn Pathfinding server
				sFile = msPathfinding
			Case ConnectionType.ePrimaryServerApp
				'Spawn Primary Server
				sFile = msPrimary
			Case ConnectionType.eRegionServerApp
				'spawn Region Server
				sFile = msRegion
		End Select

		If sFile <> "" Then
			Dim oPInfo As ProcessStartInfo = New ProcessStartInfo()
			With oPInfo
                .Arguments = mlBoxOperatorID.ToString & " " & msOperatorIP & " " & mlOperatorPort & " " & sExternalIPAddress
				.CreateNoWindow = False
				.FileName = sFile
				.UseShellExecute = False
				.WindowStyle = ProcessWindowStyle.Normal

				Dim lInstr As Int32 = sFile.LastIndexOf("\")
				If lInstr > -1 Then
					.WorkingDirectory = sFile.Substring(0, lInstr)
				End If
			End With

			Dim oProcess As Process = Process.Start(oPInfo)
			If oProcess Is Nothing = False Then
				AddProcess(oProcess, lProcessType)
				lResult = oProcess.Id
			End If
		End If
		Return lResult
	End Function
End Class
