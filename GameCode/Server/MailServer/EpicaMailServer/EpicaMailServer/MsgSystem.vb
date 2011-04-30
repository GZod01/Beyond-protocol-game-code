Public Class MsgSystem
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
    Private WithEvents moPrimaryListener As NetSock     'the one that listens for primary servers to connect
    Private moPrimary() As NetSock                      'the list of primary servers
    Private mlPrimaryUB As Int32 = -1
    Private WithEvents moOperator As NetSock

    Private mbAcceptingPrimarys As Boolean

    Public Property AcceptingPrimarys() As Boolean
        Get
            Return mbAcceptingPrimarys
        End Get
        Set(ByVal Value As Boolean)
            Dim lPort As Int32
            If mbAcceptingPrimarys <> Value Then
                mbAcceptingPrimarys = Value
                If mbAcceptingPrimarys Then
                    Dim oINI As New InitFile()
                    Try
                        lPort = CInt(Val(oINI.GetString("SETTINGS", "PrimaryListenerPort", "0")))
                        If lPort = 0 Then Err.Raise(-1, "AcceptingPrimarys", "Could not get Primary Listen Port Number from INI")
                        moPrimaryListener = Nothing
                        moPrimaryListener = New NetSock()
                        moPrimaryListener.PortNumber = lPort
                        moPrimaryListener.Listen()
                    Catch
                        MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
                        mbAcceptingPrimarys = False
                    Finally
                        oINI = Nothing
                    End Try
                Else
                    'ok, stop listening
                    moPrimaryListener.StopListening()
                End If
            End If
        End Set
    End Property

#Region "  Primary Listener Events  "
    Private Sub moPrimaryListener_onConnect(ByVal Index As Integer) Handles moPrimaryListener.onConnect
        '
    End Sub

    Private Sub moPrimaryListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moPrimaryListener.onConnectionRequest
        Dim X As Int32
        Dim lIdx As Int32 = -1

        If AcceptingPrimarys Then
            For X = 0 To mlPrimaryUB
                If moPrimary(X) Is Nothing Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                mlPrimaryUB += 1
                ReDim Preserve moPrimary(mlPrimaryUB)
                lIdx = mlPrimaryUB
            End If

            moPrimary(lIdx) = New NetSock(oClient)
            moPrimary(lIdx).SocketIndex = lIdx

            LogEvent("Primary Server Connected")

            'add event handlers
            AddHandler moPrimary(lIdx).onConnect, AddressOf moPrimary_onConnect
            AddHandler moPrimary(lIdx).onDataArrival, AddressOf moPrimary_onDataArrival
            AddHandler moPrimary(lIdx).onDisconnect, AddressOf moPrimary_onDisconnect
            AddHandler moPrimary(lIdx).onError, AddressOf moPrimary_onError

            'and then tell the socket to expect data
            moPrimary(lIdx).MakeReadyToReceive()
        Else
            LogEvent("Primary Connection Denied: Not Accepting!")
        End If
    End Sub

    Private Sub moPrimaryListener_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moPrimaryListener.onDataArrival
        '
    End Sub

    Private Sub moPrimaryListener_onDisconnect(ByVal Index As Integer) Handles moPrimaryListener.onDisconnect
        '
    End Sub

    Private Sub moPrimaryListener_onError(ByVal Index As Integer, ByVal Description As String) Handles moPrimaryListener.onError
        '
    End Sub

    Private Sub moPrimaryListener_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moPrimaryListener.onSendComplete
        '
    End Sub
#End Region

#Region "  Primary Sockets Events  "
    Private Sub moPrimary_onConnect(ByVal Index As Integer)
        '
    End Sub

    Private Sub moPrimary_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket)
        '
    End Sub

    Private Sub moPrimary_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
        'ok, gotta get the data... 
        Dim iMsgID As Int16
 
        iMsgID = System.BitConverter.ToInt16(Data, 0)
        Select Case iMsgID
            Case GlobalMessageCode.eChatMessage
                ChatManager.HandleChatMsg(Index, Data)
            Case GlobalMessageCode.eSendOutMailMsg
                HandleSendOutMailMsg(Data, Index)
            Case GlobalMessageCode.ePlayerLoginResponse
                HandlePlayerLogStatus(Index, Data)
            Case GlobalMessageCode.eChangingEnvironment
                HandlePlayerChangeEnvironments(Index, Data)
            Case GlobalMessageCode.eAddObjectCommand
                HandleAddObject(Data, Index)
            Case GlobalMessageCode.eEmailSettings
                HandleEmailSettings(Data, Index)
            Case GlobalMessageCode.eNewsItem
                Try
                    Dim oGNS As GNSMgr = GNSMgr.GetGNSMgr()
                    oGNS.ParseGNSMsg(Data)
                Catch ex As Exception
                    LogEvent("Receive News Item Error: " & ex.Message)
                End Try
            Case GlobalMessageCode.eSetGuildRel
                HandleSetPlayerGuildStatus(Index, Data)
            Case GlobalMessageCode.eRequestChannelDetails
                HandleRequestChannelDetails(Index, Data)
            Case GlobalMessageCode.eRequestChannelList
                Dim yResp() As Byte = ChatRoom.HandleRequestChannelList(Data)
                If yResp Is Nothing = False Then moPrimary(Index).SendData(yResp)
            Case GlobalMessageCode.eChatChannelCommand
                HandleChatChannelCommand(Index, Data)
            Case GlobalMessageCode.eUpdatePlayerDetails
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oPlayer As Player = GetPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    oPlayer.sPlayerName = GetStringFromBytes(Data, 50, 20)
                End If
        End Select
    End Sub

    Private Sub moPrimary_onDisconnect(ByVal Index As Integer)
        LogEvent("Primary " & Index & " disconnected.")
    End Sub

    Private Sub moPrimary_onError(ByVal Index As Integer, ByVal Description As String)
        LogEvent("Primary Connection Error (" & Index & "): " & Description)
    End Sub

    Private Sub moPrimary_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer)
        '
    End Sub
#End Region

#Region "  Operator Events  "
    Private mbConnectingToOperator As Boolean = False
    Private mbConnectedToOperator As Boolean = False
    Private moOperatorSW As Stopwatch = Nothing
    Private mcolOperatorFailQueue As Collection = Nothing
    Private msReconnectOperatorIP As String
    Private mlReconnectOperatorPort As Int32

    Public Function ConnectToOperator(ByVal sIP As String, ByVal lPort As Int32) As Boolean
        Dim mlTimeout As Int32 = 10000

        Try
            mbConnectingToOperator = True
            moOperator = New NetSock()
            moOperator.Connect(sIP, lPort)
            Dim sw As Stopwatch = Stopwatch.StartNew
            While mbConnectedToOperator = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToOperator = True
                Application.DoEvents()
            End While
            sw.Stop()
            sw = Nothing
            mbConnectingToOperator = False
            If mbConnectedToOperator = False Then
                Err.Raise(-1, "EstablishConnection", "Unable to connect to OPERATOR: Connection timed out!")
            End If
            moOperatorSW = Stopwatch.StartNew()
        Catch
            mbConnectedToOperator = False
        End Try
        Return mbConnectedToOperator
    End Function

    Private Sub moOperator_onConnect(ByVal Index As Integer) Handles moOperator.onConnect
        mbConnectedToOperator = True
        mbConnectingToOperator = False

        'Indicate to the Operator who I am... needs to indicate my server type
        Dim oEP As Net.EndPoint = moOperator.GetLocalDetails()
        If oEP Is Nothing = False Then
            With CType(oEP, Net.IPEndPoint)
                Dim yMsg(45) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(glBoxOperatorID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eEmailServerApp).CopyTo(yMsg, lPos) : lPos += 4

                Dim oMyProcess As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
                Dim lProcessID As Int32 = oMyProcess.Id
                System.BitConverter.GetBytes(lProcessID).CopyTo(yMsg, lPos) : lPos += 4
                oMyProcess = Nothing

                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4     'indicates 1 connection specifics

                'Get our Port number data
                Dim oIni As New InitFile()
                Dim lPrimaryPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "PrimaryListenerPort", "0")))
                If lPrimaryPort = 0 Then Err.Raise(-1, "moOperator.onConnect", "moOperator.onConnect: Unable to determine Primary Listen Port from INI")
                oIni = Nothing

                'now, we'll indicate our primary listener... so use our local ip address
                StringToBytes(Mid$(.Address.ToString(), 1, 20)).CopyTo(yMsg, lPos) : lPos += 20
                System.BitConverter.GetBytes(lPrimaryPort).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.ePrimaryServerApp).CopyTo(yMsg, lPos) : lPos += 4

                moOperator.SendData(yMsg)
            End With
        End If

        If mcolOperatorFailQueue Is Nothing = False Then
            For Each yAry() As Byte In mcolOperatorFailQueue
                moOperator.SendData(yAry)
            Next
            mcolOperatorFailQueue.Clear()
        End If
        mcolOperatorFailQueue = Nothing
    End Sub

    Private Sub moOperator_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moOperator.onConnectionRequest
        'should never happen
    End Sub

    Private Sub moOperator_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moOperator.onDataArrival
        Dim iMsgID As Int16
        iMsgID = System.BitConverter.ToInt16(Data, 0)

        Select Case iMsgID
            Case GlobalMessageCode.eRequestObject
                If moOperatorSW Is Nothing Then
                    moOperatorSW = Stopwatch.StartNew
                Else
                    moOperatorSW.Reset()
                    moOperatorSW.Start()
                End If
                moOperator.SendData(Data)
            Case GlobalMessageCode.eBackupOperatorSyncMsg
                HandleOperatorRebound(Data) 
        End Select
    End Sub

    Private Sub moOperator_onDisconnect(ByVal Index As Integer) Handles moOperator.onDisconnect
        LogEvent("Operator Disconnected")
    End Sub

    Private Sub moOperator_onError(ByVal Index As Integer, ByVal Description As String) Handles moOperator.onError
        LogEvent("moOperator Error: " & Description)
    End Sub

    Private Sub moOperator_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moOperator.onSendComplete
        'do nothing
    End Sub

    Public Sub SendReadyStateToOperator()
        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yMsg, 0)
        moOperator.SendData(yMsg)
    End Sub

    Public Function CheckOperatorConnection() As Boolean
        If moOperatorSW Is Nothing = False Then
            If moOperatorSW.ElapsedMilliseconds > 10000 Then
                If mbConnectedToOperator = True Then
                    moOperatorSW = Nothing
                    Return True
                End If
                Return False
            End If
        Else
            If mbConnectedToOperator = False Then
                'ok, not connected... should be... so let's connect to the backup operator
                If moOperator Is Nothing = False Then
                    Try
                        moOperator.Disconnect()
                    Catch
                    End Try
                End If
                moOperator = Nothing

                If ConnectToOperator(msReconnectOperatorIP, mlReconnectOperatorPort) = False Then
                    moOperatorSW = Stopwatch.StartNew()     'start a new stopwatch to wait 10 secs before trying again
                End If
            End If
        End If
        Return True
    End Function

    Public Sub OperatorFailure()
        msReconnectOperatorIP = gsBackupOperatorIP
        mlReconnectOperatorPort = glBackupOperatorPort
        mbConnectedToOperator = False
        moOperatorSW = Nothing
    End Sub

    Public Sub SendMsgToOperator(ByRef yData() As Byte)
        If mbConnectedToOperator = True Then
            moOperator.SendData(yData)
        Else
            If mcolOperatorFailQueue Is Nothing Then mcolOperatorFailQueue = New Collection
            mcolOperatorFailQueue.Add(yData)
        End If
    End Sub

    Public Sub HandleOperatorRebound(ByRef yData() As Byte)
        'the backup operator is telling us that the operator has returned and to connect to it
        msReconnectOperatorIP = gsOperatorIP
        mlReconnectOperatorPort = glOperatorPort

        'so, disconnect from the operator
        mbConnectedToOperator = False
        mbConnectingToOperator = False
        moOperator.Disconnect()
        moOperatorSW = Nothing

        'CheckOperatorConnection will connect us again
    End Sub
 
#End Region

    Private Sub HandleSendOutMailMsg(ByVal yData() As Byte, ByVal lIndex As Int32)
        'MsgCode(2)
        Dim lPos As Int32 = 2
        'iBaseAlertType(2)
        Dim iBaseAlertType As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'PlayerCommID(4)
        Dim lPC_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'PlayerID(4)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'lSubjectLen(4)
        Dim lSubjectLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'ySubject(varies)
        Dim sSubject As String = GetStringFromBytes(yData, lPos, lSubjectLen) : lPos += lSubjectLen
        'lBodyLen(4)
        Dim lBodyLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'yBody(varies)
        Dim sBody As String = GetStringFromBytes(yData, lPos, lBodyLen) : lPos += lBodyLen

        Dim oPlayer As Player = GetPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            'Dim lIdx As Int32 = AddNewOutboundMailMsg("fleetcommander@epicaonline.com", oPlayer.sEmailAddress, sSubject, sBody, lPC_ID, lPlayerID, iBaseAlertType)
            Dim lIdx As Int32 = AddNewOutboundMailMsg(gsOutEmailFrom, oPlayer.sEmailAddress, sSubject, sBody, lPC_ID, lPlayerID, iBaseAlertType)
            If lIdx <> -1 Then
                With goMailMsgs(lIdx)
                    Try
                        .lExtended1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lExtended2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lExtended3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lExtended4 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lExtended5 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lExtended6 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lPrimaryServerIdx = lIndex
                        .bCanHaveResponse = (.iBaseAlertType = GlobalMessageCode.eSetPlayerRel) OrElse _
                          (.iBaseAlertType = GlobalMessageCode.eColonyLowResources AndAlso .lExtended1 <> Int32.MinValue) OrElse _
                          (.iBaseAlertType = GlobalMessageCode.ePlayerAlert AndAlso .lExtended1 = PlayerAlertType.eUnderAttack) OrElse _
                          (.iBaseAlertType = GlobalMessageCode.eSetEntityProdSucceed AndAlso .lExtended2 = 1) OrElse _
                          (.iBaseAlertType = GlobalMessageCode.eRebuildAISetting) OrElse (.iBaseAlertType = GlobalMessageCode.eOutBidAlert) OrElse _
                          (.iBaseAlertType = GlobalMessageCode.eAddObjectCommand AndAlso .lExtended1 = ObjectType.ePlayerComm)

                    Catch ex As Exception
                        LogEvent(ex.Message)
                    End Try
                End With
                goOutQueue.AddNew(lIdx)
            End If
        End If

    End Sub

    Private Sub HandleAddObject(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2       'msgcode
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Ok, we have what we need to figure this out
        Select Case iTypeID
            Case ObjectType.ePlayer
                Dim lIdx As Int32 = -1
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To glPlayerUB
                    If glPlayerIdx(X) = lID Then
                        lIdx = X
                        Exit For
                    ElseIf lFirstIdx = -1 AndAlso glPlayerIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lIdx = -1 Then
                    If lFirstIdx = -1 Then
                        ReDim Preserve goPlayer(glPlayerUB + 1)
                        ReDim Preserve glPlayerIdx(glPlayerUB + 1)
                        glPlayerIdx(glPlayerUB + 1) = -1
                        glPlayerUB += 1
                        lIdx = glPlayerUB
                    Else : lIdx = lFirstIdx
                    End If
                    goPlayer(lIdx) = New Player()
                End If

                'This is a special case for add player object as the primary sends me only what I need
                With goPlayer(lIdx)
                    .PlayerID = lID
                    .bSystemAdmin = (lID = 1) OrElse (lID = 2) OrElse (lID = 6)
                    .sPlayerName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    .sCompareName = .sPlayerName.ToUpper
                    .sEmpireName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    .sRaceName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    .sEmailAddress = GetStringFromBytes(yData, lPos, 255) : lPos += 255
                    .yGender = yData(lPos) : lPos += 1
                    .lPrimaryIdx = lIndex

                    Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    If lGuildID > 0 Then
                        .oGuild = Guild.GetOrAddGuild(lGuildID)
                        .oGuild.AddMember(goPlayer(lIdx))
                    End If
                End With
                glPlayerIdx(lIdx) = lID
        End Select
    End Sub

    Private Sub HandleEmailSettings(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iValues As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Ok, get our player object
        Dim oPlayer As Player = Nothing
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                oPlayer = goPlayer(X)
                Exit For
            End If
        Next X
        If oPlayer Is Nothing Then Return

        'Ok, client changed  the email settings... get teh email address
        oPlayer.sEmailAddress = GetStringFromBytes(yData, lPos, 255).Trim : lPos += 255

        If lPos + 20 < yData.Length Then
            oPlayer.sPlayerName = GetStringFromBytes(yData, lPos, 20).Trim : lPos += 20
            oPlayer.sCompareName = oPlayer.sPlayerName.ToUpper
        End If
    End Sub

    Private Sub HandlePlayerLogStatus(ByVal lIndex As Int32, ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lAliasedAsID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yFlag As Byte = yData(lPos) : lPos += 1

        Dim oPlayer As Player = GetPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If yFlag = 0 Then
                oPlayer.PlayerLoggedOff()
            Else
                oPlayer.bIsOnline = True
                oPlayer.lPrimaryIdx = lIndex
                If lAliasedAsID > 0 Then
                    oPlayer.oAliasedAs = GetPlayer(lAliasedAsID)
                Else : oPlayer.oAliasedAs = Nothing
                End If
                If oPlayer.oAliasedAs Is Nothing = False Then
                    oPlayer.oAliasedAs.AddAliasLoggedOn(oPlayer)
                End If
                oPlayer.PlayerChangedPrimaryIndex()
                Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                oPlayer.PlayerChangeEnvironments(lEnvirID, iEnvirTypeID)

                If oPlayer.sPlayerName.StartsWith("VIP") = True Then
                    Dim lIdx As Int32 = AddNewOutboundMailMsg("support@darkskyentertainment.com", "7275345867@txt.att.net", "A VIP has logged in", oPlayer.sPlayerName & " has logged into the game.", -1, -1, -1)
                    If lIdx <> -1 Then goOutQueue.AddNew(lIdx)
                End If

                'Dim oroom As ChatRoom = ChatRoom.GetChatRoomByName("General")
                'If oroom Is Nothing = False Then
                '    oroom.AddMember(oPlayer)
                '    oPlayer.AddChannel(oroom)
                '    oroom.RemoveInvited(lPlayerID)
                'End If
            End If

            If oPlayer.oGuild Is Nothing = False AndAlso oPlayer.oAliasedAs Is Nothing = True Then
                Dim sOnOff As String
                If yFlag = 0 Then sOnOff = "Off" Else sOnOff = "On"
                Dim yMsg() As Byte = ChatManager.CreateChatMsg(-1, oPlayer.sPlayerName & " has logged " & sOnOff & ".", ChatManager.ChatMessageType.GuildChat, False, oPlayer.oGuild.lGuildID.ToString)
                Dim lPrimaryList() As Int32 = oPlayer.oGuild.GetPrimaryIdxList()
                If lPrimaryList Is Nothing = False Then
                    For X As Int32 = 0 To lPrimaryList.GetUpperBound(0)
                        Dim lRecipients() As Int32 = oPlayer.oGuild.GetRecipientList(lPrimaryList(X))
                        If lRecipients Is Nothing = False Then
                            Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                            SendToPrimary(lPrimaryList(X), yFinal)
                        End If
                    Next X
                End If
            End If
        Else
            LogEvent("ERROR: HandlePlayerLogOn PlayerID " & lPlayerID & " could not be found!")
            Return
        End If
    End Sub

    Private Sub HandleSetPlayerGuildStatus(ByVal lIndex As Int32, ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yType As Byte = yData(lPos) : lPos += 1

        Dim oPlayer As Player = GetPlayer(lPlayerID)
        If oPlayer Is Nothing Then
            LogEvent("ERROR: HandleSetPlayerGuildStatus PlayerID " & lPlayerID & " could not be found!")
            Return
        End If

        If yType = 0 Then
            'remove
            Dim oGuild As Guild = Guild.GetGuild(lGuildID)
            If oGuild Is Nothing = False Then oGuild.RemoveMember(lPlayerID)
            oPlayer.oGuild = Nothing
        Else
            'add
            Dim oGuild As Guild = Guild.GetOrAddGuild(lGuildID)
            If oGuild Is Nothing = False Then
                oPlayer.oGuild = oGuild
                oGuild.AddMember(oPlayer)
            End If
        End If
    End Sub

    Private Sub HandlePlayerChangeEnvironments(ByVal lIndex As Int32, ByVal yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oPlayer As Player = GetPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            oPlayer.PlayerChangeEnvironments(lEnvirID, iEnvirTypeID)
        Else
            LogEvent("ERROR: PlayerChangeEnvironments PlayerID " & lPlayerID & " could not be found!")
            Return
        End If
    End Sub

    Private Sub HandleRequestChannelDetails(ByVal lIndex As Int32, ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oPlayer As Player = GetPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return
        Dim sChannelName As String = GetStringFromBytes(yData, lPos, 30) : lPos += 30

        Dim oRoom As ChatRoom = ChatRoom.GetChatRoomByName(sChannelName)
        If oRoom Is Nothing Then Return
        If oRoom.PlayerInChatRoom(lPlayerID) = False Then Return

        Dim yResp() As Byte = oRoom.GetChatRoomDetailsMsg(lPlayerID)
        If yResp Is Nothing = False Then
            moPrimary(lIndex).SendData(yResp)
        End If
    End Sub

    Private Sub HandleChatChannelCommand(ByVal lIndex As Int32, ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yCmdType As Byte = yData(lPos) : lPos += 1

        Dim oPlayer As Player = GetPlayer(lPlayerID)
        If oPlayer Is Nothing Then
            LogEvent("Player does not exist in HandleChatChannelCommand. PlayerID: " & lPlayerID)
            Return
        End If

        Dim sChatRoomName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

        'Ok, determine the command type... the command type is bit-wise and will define how the rest of the msg is handled
        Dim oChatRoom As ChatRoom = Nothing
        If (yCmdType And eyChatRoomCommandType.AddNewChatRoom) <> 0 Then
            'ok, adding a new chat room... check if it already exists
            oChatRoom = ChatRoom.GetChatRoomByName(sChatRoomName)
            If oChatRoom Is Nothing = False Then
                Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "That channel name is already in use.", ChatManager.ChatMessageType.ChannelMessage, False, "")
                Dim lRecipients() As Int32 = {lPlayerID}
                Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                moPrimary(lIndex).SendData(yFinal)
                Return
            End If

            'Ok, chat room doesn't exist, create it
            oChatRoom = ChatRoom.GetOrAddChatRoomByName(sChatRoomName)
            If oChatRoom Is Nothing = False Then
                oChatRoom.AddMember(oPlayer)
                Dim bRes As Boolean = False
                oChatRoom.ToggleAdmin(oPlayer.PlayerID, bRes)
                oPlayer.AddChannel(oChatRoom)
            End If
        Else
            oChatRoom = ChatRoom.GetChatRoomByName(sChatRoomName)
        End If
        If oChatRoom Is Nothing Then
            Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "That channel does not exist or could not be created.", ChatManager.ChatMessageType.ChannelMessage, False, "")
            Dim lRecipients() As Int32 = {lPlayerID}
            Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
            moPrimary(lIndex).SendData(yFinal)
            Return
        End If

        If (yCmdType And eyChatRoomCommandType.LeaveChatRoom) <> 0 Then
            oChatRoom.RemoveMember(lPlayerID)
            oPlayer.RemoveChannel(oChatRoom)
            Return
        End If

        'go ahead and store whether the player is an admin of the channel
        Dim bIsAdmin As Boolean = oChatRoom.PlayerIsAdmin(lPlayerID)
        Dim bAdminMsgSend As Boolean = False

        'SetChannelPublic = 128 - additional parm of 1 byte
        If (yCmdType And eyChatRoomCommandType.SetChannelPublic) <> 0 Then
            Dim yNewVal As Byte = yData(lPos) : lPos += 1
            If bIsAdmin = False Then
                Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "You are not a channel admin.", ChatManager.ChatMessageType.ChannelMessage, False, "")
                Dim lRecipients() As Int32 = {lPlayerID}
                Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                moPrimary(lIndex).SendData(yFinal)
                bAdminMsgSend = True
            Else
                If yNewVal = 0 Then
                    'clear public
                    If (oChatRoom.yAttrs And eyChatRoomAttr.PublicChannel) <> 0 Then oChatRoom.yAttrs = oChatRoom.yAttrs Xor eyChatRoomAttr.PublicChannel
                    oChatRoom.SendToRoomMembers(lPlayerID, oPlayer.sPlayerName & " has set '" & oChatRoom.sChannelName & "' to private.", False)
                Else
                    'set public
                    oChatRoom.yAttrs = oChatRoom.yAttrs Or eyChatRoomAttr.PublicChannel
                    oChatRoom.SendToRoomMembers(lPlayerID, oPlayer.sPlayerName & " has set '" & oChatRoom.sChannelName & "' to public.", False)
                End If
            End If
        End If

        'GrantAdminRights = 4 - additional parm is playerid to be granted rights to
        'KickPlayer = 32 - additional parm is thep layer to kick
        If (yCmdType And (eyChatRoomCommandType.ToggleAdminRights Or eyChatRoomCommandType.KickPlayer)) <> 0 Then
            Dim lTargetID As Int32 = Math.Abs(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            If (yCmdType And eyChatRoomCommandType.ToggleAdminRights) <> 0 Then
                If bIsAdmin = False Then
                    If bAdminMsgSend = False Then
                        Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "You are not a channel admin.", ChatManager.ChatMessageType.ChannelMessage, False, "")
                        Dim lRecipients() As Int32 = {lPlayerID}
                        Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                        moPrimary(lIndex).SendData(yFinal)
                        bAdminMsgSend = True
                    End If
                Else
                    Dim bVal As Boolean = False
                    Dim sTargetName As String = oChatRoom.ToggleAdmin(lTargetID, bVal)
                    If sTargetName <> "" Then
                        If bVal = True Then oChatRoom.SendToRoomMembers(lPlayerID, oPlayer.sPlayerName & " has set " & sTargetName & " as a channel admin.", False) Else oChatRoom.SendToRoomMembers(lPlayerID, oPlayer.sPlayerName & " has removed " & sTargetName & " as a channel admin.", False)
                    End If
                End If
            End If
            If (yCmdType And eyChatRoomCommandType.KickPlayer) <> 0 Then
                If bIsAdmin = False Then
                    If bAdminMsgSend = False Then
                        Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "You are not a channel admin.", ChatManager.ChatMessageType.ChannelMessage, False, "")
                        Dim lRecipients() As Int32 = {lPlayerID}
                        Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                        moPrimary(lIndex).SendData(yFinal)
                        bAdminMsgSend = True
                    End If
                Else
                    'Kick the target player
                    Dim oTarget As Player = GetPlayer(lTargetID)
                    If oTarget Is Nothing = False Then
                        oChatRoom.SendToRoomMembers(lPlayerID, oPlayer.sPlayerName & " has kicked " & oTarget.sPlayerName & " from '" & oChatRoom.sChannelName & "'.", False)
                        oChatRoom.RemoveMember(lTargetID)
                        oTarget.RemoveChannel(oChatRoom)
                    End If
                End If
            End If
        End If

        'JoinChannel = 8 - additional parm is the password
        'SetChannelPassword = 16 - additional parm is hte password
        If (yCmdType And (eyChatRoomCommandType.JoinChannel Or eyChatRoomCommandType.SetChannelPassword)) <> 0 Then
            Dim sPassword As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            If (yCmdType And eyChatRoomCommandType.JoinChannel) <> 0 Then
                If oChatRoom.sChannelPassword Is Nothing OrElse oChatRoom.sChannelPassword = "" OrElse oChatRoom.sChannelPassword = sPassword Then
                    oChatRoom.AddMember(oPlayer)
                    oPlayer.AddChannel(oChatRoom)
                    If oChatRoom.sChannelName.ToUpper <> "GENERAL" AndAlso oChatRoom.sChannelName.ToUpper <> "EMPEROR" Then oChatRoom.SendToRoomMembers(-1, oPlayer.sPlayerName & " has joined '" & oChatRoom.sChannelName & "'.", False)
                    oChatRoom.RemoveInvited(lPlayerID)
                Else
                    Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "Incorrect password supplied for that channel.", ChatManager.ChatMessageType.ChannelMessage, False, "")
                    Dim lRecipients() As Int32 = {lPlayerID}
                    Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                    moPrimary(lIndex).SendData(yFinal)
                    Return
                End If
            End If
            If (yCmdType And eyChatRoomCommandType.SetChannelPassword) <> 0 Then
                If bIsAdmin = False Then
                    If bAdminMsgSend = False Then
                        Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "You are not a channel admin.", ChatManager.ChatMessageType.ChannelMessage, False, "")
                        Dim lRecipients() As Int32 = {lPlayerID}
                        Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                        moPrimary(lIndex).SendData(yFinal)
                        bAdminMsgSend = True
                    End If
                Else
                    If sPassword Is Nothing OrElse sPassword.Trim = "" Then
                        oChatRoom.sChannelPassword = ""
                        oChatRoom.SendToRoomMembers(lPlayerID, oPlayer.sPlayerName & " has cleared the password for '" & oChatRoom.sChannelName & "'.", False)
                    Else
                        oChatRoom.sChannelPassword = sPassword
                        oChatRoom.SendToRoomMembers(lPlayerID, oPlayer.sPlayerName & " has changed the password for '" & oChatRoom.sChannelName & "'.", False)
                    End If
                End If
            End If
        End If

        'InvitePlayer = 64 - additional parm is hte player to invite
        If (yCmdType And eyChatRoomCommandType.InvitePlayer) <> 0 Then
            Dim sPlayerToInvite As String = GetStringFromBytes(yData, lPos, 30) : lPos += 30
            Dim sSearch As String = sPlayerToInvite.ToUpper
            Dim bFound As Boolean = False
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) > -1 Then
                    Dim oInvited As Player = goPlayer(X)
                    If oInvited Is Nothing = False AndAlso oInvited.sCompareName = sSearch Then
                        Dim sMsgToSend As String = oPlayer.sPlayerName & " has invited you to join '" & oChatRoom.sChannelName & "'."
                        If oChatRoom.sChannelPassword Is Nothing = False AndAlso oChatRoom.sChannelPassword.Length > 0 Then
                            sMsgToSend &= " The password to join is " & oChatRoom.sChannelPassword & "."
                        End If
                        Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, sMsgToSend, ChatManager.ChatMessageType.ChannelMessage, False, "")
                        Dim lRecipients() As Int32 = {oInvited.PlayerID}
                        Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                        Dim lPrimIdx As Int32 = oInvited.lPrimaryIdx
                        If lPrimIdx > -1 AndAlso lPrimIdx <= mlPrimaryUB Then moPrimary(lPrimIdx).SendData(yFinal)
                        bFound = True

                        oChatRoom.AddInvited(oInvited.PlayerID)
                        Exit For
                    End If
                End If
            Next X

            If bFound = False Then
                Dim yMsg() As Byte = ChatManager.CreateChatMsg(lPlayerID, "Could not find a player by that name.", ChatManager.ChatMessageType.ChannelMessage, False, "")
                Dim lRecipients() As Int32 = {lPlayerID}
                Dim yFinal() As Byte = ChatManager.AppendRecipientList(yMsg, lRecipients)
                moPrimary(lIndex).SendData(yFinal)
            End If
        End If

    End Sub

    Public Function GetStringFromBytes(ByRef yData() As Byte, ByVal lPos As Int32, ByVal lDataLen As Int32) As String
        Dim yTemp(lDataLen - 1) As Byte
        Array.Copy(yData, lPos, yTemp, 0, lDataLen)
        Dim lLen As Int32 = yTemp.Length
        For Y As Int32 = 0 To yTemp.Length - 1
            If yTemp(Y) = 0 Then
                lLen = Y
                Exit For
            End If
        Next Y
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)
    End Function

    Public Sub DisconnectAll()
        For X As Int32 = 0 To mlPrimaryUB
            Try
                If moPrimary(X) Is Nothing = False Then
                    moPrimary(X).Disconnect()
                    moPrimary(X) = Nothing
                End If
            Catch
                'do nothing
            End Try
        Next X
    End Sub

    Public Sub SendToPrimary(ByVal lIndex As Int32, ByRef yData() As Byte)
        If yData Is Nothing Then Return
        If lIndex > -1 AndAlso lIndex <= mlPrimaryUB Then
            Try
                moPrimary(lIndex).SendData(yData)
            Catch ex As Exception
                LogEvent("Unable to send to primary " & lIndex & ": " & ex.Message)
            End Try
        End If
    End Sub

    Public Sub SendToAllPrimaries(ByVal yData() As Byte)
        If moPrimary Is Nothing = False Then
            For X As Int32 = 0 To moPrimary.GetUpperBound(0)
                If moPrimary(X) Is Nothing = False Then moPrimary(X).SendData(yData)
            Next X
        End If
    End Sub
End Class