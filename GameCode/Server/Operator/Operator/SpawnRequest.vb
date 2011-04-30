Option Strict On

Public Structure SpawnRequest
    'Where the spawn request was sent to
    Public lBoxOperatorID As Int32
    Public oBoxOperator As ServerObject

    'What kind of connection to spawn
    Public lConnectionType As ConnectionType

    'When the spawn was requested
    Public dtRequest As DateTime

    Public ySpawnRequestState As Byte       '0 = unused, 1 = requested but not fulfilled, 2 = fulfilled and active (process running)

    'The GUID of the items this is being spawned for (they will be assigned upon connection)
    Public lSpawnID() As Int32
    Public iSpawnTypeID() As Int16
    Public lSpawnUB As Int32

    ''' <summary>
    ''' For Email Server, this is nothing. For Primary Server, this is an Email Server. For Pathfinding, this is a Primary. For Region, this is a Primary.
    ''' </summary>
    ''' <remarks></remarks>
    Public oRelatedServer As ServerObject

    Public Sub SendSpawnRequest(ByVal lSpawnRequestIdx As Int32)
        If oBoxOperator Is Nothing = True Then oBoxOperator = goMsgSys.GetBoxOperator(lBoxOperatorID)
        If oBoxOperator Is Nothing = True Then Return
        If oBoxOperator.uListeners Is Nothing = True Then Return

        oBoxOperator.AddPendingSpawnRequestIdx(lSpawnRequestIdx)

        If oBoxOperator.oSocket Is Nothing OrElse oBoxOperator.oSocket.IsConnected = False OrElse oBoxOperator.bSocketConnected = False Then
            Dim sIP As String = ""
            Dim lPort As Int32 = 0
            For X As Int32 = 0 To oBoxOperator.uListeners.GetUpperBound(0)
                If oBoxOperator.uListeners(X).lConnectionType = ConnectionType.eOperator Then
                    sIP = oBoxOperator.uListeners(X).sIPAddress
                    lPort = oBoxOperator.uListeners(X).lPortNumber
                    Exit For
                End If
            Next X
            If sIP <> "" AndAlso lPort <> 0 Then
                'Connect to the Box Operator
                LogEvent(LogEventType.Informational, "Connecting to Box Operator " & lBoxOperatorID & " @ " & sIP & ":" & lPort)
                goMsgSys.ConnectToBoxOperator(oBoxOperator.oSocket.SocketIndex, sIP, lPort, oBoxOperator)
            Else
                LogEvent(LogEventType.CriticalError, "Unable to determine Box Operator's Listener Credentials: Box Operator ID " & lBoxOperatorID)
            End If
        Else : oBoxOperator.SendPendingSpawnRequests()
        End If
    End Sub

End Structure
