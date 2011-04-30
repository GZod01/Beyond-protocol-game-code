Option Strict On

Public Structure ExternalConnData
    Public sIPAddress As String
    Public lPortNumber As Int32
    Public lConnectionType As ConnectionType
End Structure

Public Class ServerObject
    ''' <summary>
    ''' Connection info that the server listens on with the connection type
    ''' </summary>
    ''' <remarks></remarks>
    Public uListeners() As ExternalConnData

    Public bSocketConnected As Boolean = True   'set to true until otherwise told 

    ''' <summary>
    ''' The connection to the server app
    ''' </summary>
    ''' <remarks></remarks>
    Public oSocket As NetSock

    ''' <summary>
    ''' The connection type for this server object
    ''' </summary>
    ''' <remarks></remarks>
    Public lConnectionType As ConnectionType

    Public lProcessID As Int32 = 0

    ''' <summary>
    ''' Box operator ID
    ''' </summary>
    ''' <remarks></remarks>
    Public lBoxOperatorID As Int32
	Public oBoxOperator As BoxOperator

	Public sIPAddress As String

    'Public lSpawnRequestIdx As Int32 = -1       'index in the spawnrequest array of the spawn request that spawned me

    'Reported by primary servers only!!!
    Public lSpawnLocsAvail As Int32 = 0
    Public lRespawnLocsAvail As Int32 = 0
    'end of primary server only stuff

    Public lSpawnID() As Int32
    Public iSpawnTypeID() As Int16
    Public lSpawnUB As Int32 = -1
    Public oRelatedServer As ServerObject = Nothing

	Public dtLastMsgRecd As Date
	Public bResentKeepAlive As Boolean = False

	Private mlPendingSpawnRequestIdx() As Int32	'index in the spawnrequest array of the spawn request that spawned me
    Private mlPendingSpawnRequestUB As Int32 = -1
    Public Sub AddPendingSpawnRequestIdx(ByVal lIdx As Int32)
        Dim lFirstIdx As Int32 = -1
        For X As Int32 = 0 To mlPendingSpawnRequestUB
            If mlPendingSpawnRequestIdx(X) = lIdx Then
                Return
            ElseIf mlPendingSpawnRequestIdx(X) = -1 AndAlso lFirstIdx = -1 Then
                lFirstIdx = X
            End If
        Next X
        If lFirstIdx = -1 Then
            mlPendingSpawnRequestUB += 1
            ReDim Preserve mlPendingSpawnRequestIdx(mlPendingSpawnRequestUB)
            lFirstIdx = mlPendingSpawnRequestUB
        End If
        mlPendingSpawnRequestIdx(lFirstIdx) = lIdx
    End Sub
    Public Sub SendPendingSpawnRequests()
        For X As Int32 = 0 To mlPendingSpawnRequestUB
            If mlPendingSpawnRequestIdx(X) <> -1 Then
                ' Send spawn request
                Dim uSpawnRequest As SpawnRequest = uSpawnRequests(mlPendingSpawnRequestIdx(X))

                LogEvent(LogEventType.Informational, "Send Box Operator " & lBoxOperatorID & " spawn request for " & uSpawnRequest.lConnectionType.ToString)

                Dim yMsg(9) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(uSpawnRequest.lConnectionType).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(mlPendingSpawnRequestIdx(X)).CopyTo(yMsg, 6)
                oSocket.SendData(yMsg)

                mlPendingSpawnRequestIdx(X) = -1
            End If
        Next X
    End Sub

    Public Sub New()
		dtLastMsgRecd = Now
    End Sub

    Public Function GetSpawnIDScore() As Int32
        Dim lScore As Int32 = 0
        For X As Int32 = 0 To lSpawnUB
            If iSpawnTypeID(X) = ObjectType.eSolarSystem Then
                Dim oSys As SolarSystem = GetEpicaSystem(lSpawnID(X))
                If oSys Is Nothing = False Then
                    'Ok, lil smarter
                    ' spawnsystems are 5
                    ' respawnsystems are 4
                    ' hubs are 2
                    ' hub hubs are 1
                    ' aurelium is 10

                    Select Case CType(oSys.SystemType, GeoSpawner.elSystemType)
                        Case GeoSpawner.elSystemType.SpawnSystem
                            lScore += 5
                        Case GeoSpawner.elSystemType.TutorialSystem
                            lScore += 10
                        Case GeoSpawner.elSystemType.RespawnSystem, GeoSpawner.elSystemType.UnlockedSystem
                            lScore += 4
                        Case GeoSpawner.elSystemType.HubHubSystem
                            lScore += 1
                        Case GeoSpawner.elSystemType.HubSystem
                            lScore += 2
                    End Select
                End If
            End If
        Next X
        Return lScore
    End Function
End Class

Public Class BoxOperator
	Inherits ServerObject

	Public Class ServerProcess
		''' <summary>
		''' Spawn Requested Server Object through this Box Operator
		''' </summary>
		''' <remarks></remarks>
		Public oServerObject As ServerObject
		''' <summary>
		''' Process ID resulting from the process start command on the box operator
		''' </summary>
		''' <remarks></remarks>
		Public lProcessID As Int32
		''' <summary>
		''' Connection Type for this process
		''' </summary>
		''' <remarks></remarks>
		Public lType As ConnectionType
		''' <summary>
		''' Priority for the process on the box
		''' </summary>
		''' <remarks></remarks>
		Public BasePriority As Int32
		''' <summary>
		''' True or false indicating whether the process is responding
		''' </summary>
		''' <remarks></remarks>
		Public bResponding As Boolean
		''' <summary>
		''' True or False indicating whether the process has exited
		''' </summary>
		''' <remarks></remarks>
		Public bHasExited As Boolean
		''' <summary>
		''' Total number of seconds the process has ran (becomes invalid after the process exits)
		''' </summary>
		''' <remarks></remarks>
		Public dSecondsRan As Double
		''' <summary>
		''' Total number of seconds the process has utilized the processor
		''' </summary>
		''' <remarks></remarks>
		Public dTotalProcessorTime As Double
		''' <summary>
		''' Total number of seconds the process has utilized the processor for the process' internal code
		''' </summary>
		''' <remarks></remarks>
		Public dTotalUserProcessorTime As Double
		''' <summary>
		''' Total number of seconds the process has utilized the processor for the process' utilization of the kernel
		''' </summary>
		''' <remarks></remarks>
		Public dPrivilegedProcessorTime As Double

		''' <summary>
		''' UInt64 for amount of non paged system memory used for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulNonPagedSystemMemory As UInt64
		''' <summary>
		''' UInt64 for amount of Paged system memory used for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulPagedSystemMemory As UInt64
		''' <summary>
		''' UInt64 for amount of Paged Memory used for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulPagedMemory As UInt64
		''' <summary>
		''' UInt64 for amount of Private Memory used for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulPrivateMemory As UInt64
		''' <summary>
		''' UInt64 for amount of Virtual Memory used for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulVirtualMemory As UInt64
		''' <summary>
		''' UInt64 for amount of Working Set memory for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulWorkingSet As UInt64

		''' <summary>
		''' UInt64 for highest paged memory for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulPeakPagedMemory As UInt64
		''' <summary>
		''' UInt64 for highest Virtual Memory for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulPeakVirtualMemory As UInt64
		''' <summary>
		''' UInt64 for highest Working Set memory for this process
		''' </summary>
		''' <remarks></remarks>
		Public bulPeakWorkingSet As UInt64

		''' <summary>
		''' Number of threads this process currently has running
		''' </summary>
		''' <remarks></remarks>
		Public lThreadCount As Int32
	End Class

	Public lProcessorCount As Int32 = 0
	Public bulTotalPhysicalMemory As UInt64 = 0
	Public bulTotalVirtualMemory As UInt64 = 0
	Public bulAvailablePhysicalMemory As UInt64 = 0
	Public bulAvailableVirtualMemory As UInt64 = 0

	Private moProcs() As ServerProcess = Nothing

	Public Sub HandleBoxOperatorUpdate(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'for msgcode
		Dim lProcCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		For X As Int32 = 0 To lProcCnt - 1
			Dim lType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4


			Dim lBasePriority As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim bResponding As Boolean = yData(lPos) <> 0 : lPos += 1
			Dim bHasExited As Boolean = yData(lPos) <> 0 : lPos += 1
			Dim dSecondsRan As Double = System.BitConverter.ToDouble(yData, lPos) : lPos += 8
			Dim dTotalProcTime As Double = System.BitConverter.ToDouble(yData, lPos) : lPos += 8
			Dim dUserProcTime As Double = System.BitConverter.ToDouble(yData, lPos) : lPos += 8
			Dim dPrivProcTime As Double = System.BitConverter.ToDouble(yData, lPos) : lPos += 8

			Dim bulNonPagedSystem As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
			Dim bulPagedSystem As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
			Dim bulPagedMemory As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
			Dim bulPrivateMemory As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
			Dim bulVirtualMemory As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
			Dim bulWorkingSet As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8

			Dim bulPeakPaged As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
			Dim bulPeakVirtual As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
			Dim bulPeakWorkingSet As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8

			Dim lThreadCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			Dim oProc As ServerProcess = Nothing
			If moProcs Is Nothing = False Then
				For Y As Int32 = 0 To moProcs.GetUpperBound(0)
					If moProcs(Y) Is Nothing = False Then
						If moProcs(Y).lProcessID = lID AndAlso moProcs(Y).lType = lType Then
							oProc = moProcs(Y)
							Exit For
						End If
					End If
				Next Y
			End If
			If oProc Is Nothing = False Then
				With oProc
					.BasePriority = lBasePriority
					.bResponding = bResponding
					.bHasExited = bHasExited
					.dSecondsRan = dSecondsRan
					.dTotalProcessorTime = dTotalProcTime
					.dTotalUserProcessorTime = dUserProcTime
					.dPrivilegedProcessorTime = dPrivProcTime
					.bulNonPagedSystemMemory = bulNonPagedSystem
					.bulPagedMemory = bulPagedSystem
					.bulPagedMemory = bulPagedMemory
					.bulPrivateMemory = bulPrivateMemory
					.bulVirtualMemory = bulVirtualMemory
					.bulWorkingSet = bulWorkingSet
					.bulPeakPagedMemory = bulPeakPaged
					.bulPeakVirtualMemory = bulPeakVirtual
					.bulPeakWorkingSet = bulPeakWorkingSet
					.lThreadCount = lThreadCnt
				End With
			End If
		Next X
 

	End Sub

	Public Sub New(ByRef oServerObject As ServerObject)
		With oServerObject
			bResentKeepAlive = .bResentKeepAlive
			bSocketConnected = .bSocketConnected
			lBoxOperatorID = .lBoxOperatorID
			lConnectionType = .lConnectionType
            lSpawnLocsAvail = .lSpawnLocsAvail
            lRespawnLocsAvail = .lRespawnLocsAvail
			dtLastMsgRecd = .dtLastMsgRecd
            'lSpawnRequestIdx = .lSpawnRequestIdx
			oBoxOperator = .oBoxOperator
			oSocket = .oSocket
            uListeners = .uListeners
            sIPAddress = .sIPAddress
		End With
	End Sub

	Public Function CanHandleNewServer(ByVal lConnType As ConnectionType) As Boolean
		If moProcs Is Nothing Then Return True

		'Our estimated memory footprints
		Dim bulPhysicalMemoryFootprint As UInt64 = 0
		Dim bulVirtualMemoryFootprint As UInt64 = 0
		Dim fProcessorConsumption As Single = 0.0F

		Select Case lConnType
			Case ConnectionType.eEmailServerApp
				fProcessorConsumption = 0.2F
				bulPhysicalMemoryFootprint = 30000000
				bulVirtualMemoryFootprint = 30000000
			Case ConnectionType.ePathfindingServerApp
				fProcessorConsumption = 0.5F
				bulPhysicalMemoryFootprint = 500000000
				bulVirtualMemoryFootprint = bulPhysicalMemoryFootprint
			Case ConnectionType.ePrimaryServerApp
				fProcessorConsumption = 1.0F
				bulPhysicalMemoryFootprint = 500000000
				bulVirtualMemoryFootprint = bulPhysicalMemoryFootprint
			Case ConnectionType.eRegionServerApp
				fProcessorConsumption = 1.0F
				bulPhysicalMemoryFootprint = 1500000000
				bulVirtualMemoryFootprint = bulPhysicalMemoryFootprint
			Case Else
				Return True
		End Select

		'Ok, find out we already have a process like this running
		Dim fProcessorUsage As Single = 0.0F
		For X As Int32 = 0 To moProcs.GetUpperBound(0)
			If moProcs(X).lType = lConnType Then Return False
			fProcessorUsage += CSng(moProcs(X).dSecondsRan / moProcs(X).dTotalProcessorTime)
		Next X

		If fProcessorUsage + fProcessorConsumption > Me.lProcessorCount Then
			Return False
		End If

		If bulPhysicalMemoryFootprint > bulAvailablePhysicalMemory Then Return False
		If bulVirtualMemoryFootprint > bulAvailableVirtualMemory Then Return False

		Return True
    End Function

    Public Sub AddServerProc(ByRef oSrvr As ServerObject)
        Dim oProc As New ServerProcess
        With oProc
            .oServerObject = oSrvr
            .bHasExited = False
            .bResponding = True
            .lProcessID = oSrvr.lProcessID
            .lType = oSrvr.lConnectionType
        End With

        If moProcs Is Nothing Then ReDim moProcs(-1)
        SyncLock moProcs

            For X As Int32 = 0 To moProcs.GetUpperBound(0)
                If moProcs(X).lProcessID = oProc.lProcessID Then
                    moProcs(X) = oProc
                    Return
                End If
            Next X

            ReDim Preserve moProcs(moProcs.GetUpperBound(0) + 1)
            moProcs(moProcs.GetUpperBound(0)) = oProc
        End SyncLock
    End Sub
End Class