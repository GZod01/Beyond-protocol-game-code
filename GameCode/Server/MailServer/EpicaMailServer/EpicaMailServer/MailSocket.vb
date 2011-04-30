Option Strict On

Imports System.Net
Imports System.Net.Sockets

'TODO: Need to make sure these get updated as needed!!!
Public Enum GlobalMessageCode As Short
    eRequestObject = 0
    eRequestObjectResponse

    eStopObjectRequest          'the client is requesting the object to stop
    eStopObjectCommand          'the server is commanding the stop object
    eStopObjectRequestDeny      'denying the stop object request
    eMoveObjectRequest          'the client is requesting the object to move
    eMoveObjectCommand          'this is the actual command
    eMoveObjectRequestDeny      'denying the move object request

    eUnitHPUpdate               'the specified unit's hitpoints have been updated

    eFireWeapon                 'Fire Weapon packet (Weapon hits)
    eFireWeaponMiss             'Fire weapon packet (weapon misses)

    eBurstEnvironmentRequest    'a client is requesting to receive an Environments data
    eBurstEnvironmentResponse   'the response with the environment data

    ePathfindingConnectionInfo  'the pathfinding server's connection information

    eAddObjectCommand           'add object thru a BurstMessage, default Add Object... Undock and Created use this too
    eAddObjectCommand_CE        'Add Object thru Change environments ONLY as the client will display it differently
    eRemoveObject               'object removal

    ePlayerLoginRequest         'player wishes to login with Username/Password
    ePlayerLoginResponse        'response to the login message, a negative value indicates failure, otherwise, it is the PlayerID

    eRequestPlayerList          'requests the ENTIRE player list. To be used only for region server initialization
    eRequestPlayerRelList       'requests all player rels. To be used only for region server initialization
    eDomainRequestEnvirObjects  'domain server requesting all objects from the primary server for an environment
    eRegisterDomain             'the domain server is telling the primary server what domain it controls
    eDomainServerReady          'the domain server is fully initialized and is ready

    eAddPlayerRelObject

    eRequestPlayerEnvironment   'the client is asking what environment the player passed in (ID) was last looking at
    ePlayerCurrentEnvironment   'the response to the RequestPlayerEnvironment message
    eRequestEnvironmentDomain   'requests the Domain Credentials for the passed in Environment
    eEnvironmentDomain          'the response to the RequestEnvironmentDomain message
    eChangingEnvironment        'the player is notifying a server that they are changing environments, contains new ID's

    eRequestGalaxyAndSystems    'the client is asking for the galaxy and system objects from the Primary
    eResponseGalaxyAndSystems   'the response from the Primary to the Client containing the Galaxy and System objects

    eRequestStarTypes           'the client is asking for the star type definitions from the Primary
    eResponseStarTypes          'the response from the Primary to the Client containing the Star Types

    eRequestSystemDetails       'the client needs the planet data for the items in the system
    eResponseSystemDetails      'the response from the primary to the client containing the details of the system
    eResponseSystemDetailsFail  'the response from the primary to the client indicating that the request failed

    eSetEntityProduction        'a request to the Primary Server to set the production of an entity
    eSetEntityProdFailed        'response to eSetEntityProduction request as a failure
    eSetEntityProdSucceed       'response to eSetEntityProduction request as a success

    eGetEntityProduction        'request and response to the Primary Server

    eRequestPlayerDetails       'client is requesting their player details

    eAddWaypointMsg             'a client is requesting to add a waypoint to a unit's destination list
    eDecelerationImminent       'domain server is notifying the pathfinding server that an entity is decelerating

    eEntityChangeEnvironment    'Change environment message indicating an entity is moving or has moved between environments

    eGetEntityDetails           'communication for getting/responding entity details...

    eRequestEntityContents      'used for request and response of an entity's contents (both cargo and hangar)

    eSetPrimaryTarget           'Client to Server, informs server to set primary target
    eRequestDock                'docking request
    eRequestDockFail            'docking request failed
    eRequestUndock              'request to undock unit
    eUpdatePlayerCredits        'Primary server telling client the player's current credits

    eSetPlayerRel               'Indicating a relationship change

    eEntityProductionStatus     'used to request and respond regarding the entity production status

    eSetMiningLoc               'client is telling unit to mine somewhere... 

    eSetEntityStatus            'to region server
    eMoveLockViolate            'move lock was violated

    eProductionCompleteNotice   'msg sent to clients that production was finished

    eHangarCargoBayDestroyed    'msg sent from Region to Primary to indicate either (or both) the Hangar/Cargo Bay was destroyed

    eGetAvailableResources      'msg for getting the available resources of an entity

    eDockCommand                'msg from primary to region, special

    eBombardFireMsg             'a special type of fire weapon event
    eRequestOrbitalBombard      'for requesting orbital bombardment
    eStopOrbitalBombard         'requesting to stop orbital bombardment

    eSetEntityAI                'sets the entity's AI settings

    eChatMessage                'a chat message

    eRequestSkillList           'get te skill list for agents
    eRequestGoalList            'get the goal list for agents

    eReloadWpnMsg               'sent to and from primary/domain

    eGetEntityName              'for client caching
    eSetEntityName

    eTransferContents           'for transferring cargo, ammo, fuel, minerals, ships between entities

    eGetColonyDetails           'for getting the colony details

    eRequestPlayerStartLoc      'from primary to region and back... gets a valid starting location for a player's first unit

    eSetEntityPersonnel         'from primary to client

    eServerShutdown             'from primary to region indicating a command to shut it down... from region to primary indicating domain is shut down... from primary to client, indicates server shutdown
    eUpdateEntityAndSave        'from region to primary

    eColonyLowResources         'some resource is low on the colony...

    eBugList                    'from client to primary indicates request for bug list... from primary to client indicates a bug list item
    eBugSubmission              'from client to primary indicates a new bug

    ePFRequestEntitys           'both way...

    eSubmitComponentPrototype       'from client to server
    eComponentPrototypeResponse     'from server to client

    eClientVersion
    eMineralDetailsMsg

    eSetColonyTaxRate

    eGetColonyChildList         'Get colony child list (facilities in the colony)

    eClearDestList              'from region to primary

    eUpdateEntityAttrs          'from region to client

    eUpdateCommandPoints        'from region to client, also serves as a keepalive

    eSetRallyPoint              'from Client to Primary

    eCreateFleet                'from client to primary
    eRemoveFromFleet            'from client to primary
    eAddToFleet                 'from client to primary
    eDeleteFleet                'from client to primary

    eUpdateFleetSpeed
    eSetFleetDest
    eFleetInterSystemMoving

    eUpdateEntityParent

    'Email related
    eAddPlayerCommFolder
    eSendEmail
    eSaveEmailDraft
    eDeleteEmailItem
    eMoveEmailToFolder
    eMarkEmailReadStatus

    ePlayerAlert

    eSubmitTrade

    eGetColonyList
    eGetTradePostTradeables
    eGetNonOwnerDetails         'gets an item's details based on the fact that the player is not the owner

    eRequestEntityDefenses
    eRepairOrder

    eRequestPlayerBudget

    eGetPirateStartLoc
    ePlacePirateAssets

    eUpdateEntityLoc

    eGetPlayerScores

    ePlayerTitleChange

    eGetEnvirConstructionObjects

    eMsgMonitorData

    ePlayerInitialSetup

    eDeathBudgetDeposit

    ePlayerIsDead

    eUpdateResearcherCnt

    eUpdateUnitExpLevel

    eMoveEngineer
    eSetDismantleTarget
    eSetRepairTarget
    eRepairCompleted
    eDeleteDesign

    eMissileFired
    eMissileDetonated
    eMissileImpact

    eRequestEntityAmmo
    eRequestLoadAmmo

    eSendOutMailMsg
    eEmailSettings

    eGetGTCList
    eGetTradeHistory
    eSubmitSellOrder
    eSubmitBuyOrder
    ePurchaseSellOrder
    eGetTradePostList
    eGetOrderSpecifics
    eAcceptBuyOrder
    eDeliverBuyOrder
    eGetTradeDeliveries

    eUpdatePlayerTechValue

    eFinalizeStopEvent  'Sent to and from Region/PF srvrs to determine final stop movements

    eAddPlayerTechKnow

    eSetPlayerSpecialAttribute

    eEntityDefCriticalHitChances

    eSetCounterAttack

    ePlayerDiscoversWormhole
    eJumpTarget
    eSetFleetReinforcer
    ePlayerAliasConfig
    eActOfWarNotice

    eFinalMoveCommand

    eDeinfiltrateAgent
    eSetInfiltrateSettings
    eSubmitMission
    eSetAgentStatus
    eGetAgentStatus
    eGetPMUpdate
    eSetSkipStatus
    eCaptureKillAgent

    eNewsItem

    eCameraPosUpdate
    ePlayerDomainAssignment
    eRemovePlayerRel

    eRequestPlayerDetailsPlayerMin
    eRequestPlayerDetailsAlert
    eRequestMineral
    eRequestEmailSummarys
    eRequestEmail
    eRequestAliasConfigs

    eRequestDXDiag

    eRouteMoveCommand
    eGetRouteList
    eSetRouteMineral
    eRemoveRouteItem
    eUpdateRouteStatus

    eGetSkillList
    eRebuildAISetting

    eAddFormation
    eRemoveFormation
    eMoveFormation
    eSetPrimaryTargetFormation

    eUpdateSlotStates
    eSetIronCurtain
    eGetItemIntelDetail
    eAlertDestinationReached
    eGetIntelSellOrderDetail

    eSetArchiveState
    eGetArchivedItems
    eGetColonyResearchList
    eSetColonyResearchQueue

    eAgentMissionCompleted

    eTutorialProdFinish
    eCreatePlanetInstance
    eSaveAndUnloadInstance
    eUpdatePlayerTutorialStep
    eAILaunchAll

    eBackupOperatorSyncMsg

    eUpdatePlayerTimer
    ePirateWaveSpawn

    eForcedMoveSpeedMove

    eUpdateMOTD
    eGuildBankTransaction
    eGuildRequestDetails
    eUpdatePlayerVote
    eUpdateGuildRank
    eUpdateRankPermission
    eGuildMemberStatus
    eInvitePlayerToGuild
    eUpdateGuildEvent
    eProposeGuildVote
    eSearchGuilds
    eRequestContactWithGuild
    eSetGuildRel
    eAddEventAttachment
    eUpdateGuildRecruitment
    eRemoveEventAttachment
    eUpdateGuildRelNotes
    eAdvancedEventConfig
    eCheckGuildName
    eCreateGuild
    eInviteFormGuild
    eGetGuildInvites
    eGetMyVoteValue
    eRequestGuildEvents
    eUpdateGuildTreasury

    eSetEntityTarget

    eGetGuildAssets
    ePlayerRestart

    eAddSenateProposalMessage
    eAddSenateProposal
    eGetSenateObjectDetails
    eGuildCreationUpdate
    eGuildCreationAcceptance

    eRequestChannelList
    eRequestChannelDetails
    eDeleteTradeHistoryItem
    eSetFleetFormation
    eGetImposedAgentEffects

    eUpdatePlayerDetails
    eUpdatePlayerCustomTitle

    eClearBudgetAlert
    eAbandonColony

    ePDS_AtMissileHit
    ePDS_AtMissileMiss
    ePDS_AtUnitHit
    ePDS_AtUnitMiss

    eTutorialGiveCredits

    eChatChannelCommand

    eForcePrimarySync
    eOperatorRequestPassThru
    eOperatorResponsePassThru
    eUpdatePlanetOwnership

    ePlayerConnectedPrimary
    ePlayerPrimaryOwner
    eForwardToPlayerAtPrimary
    eEntityChangingPrimary
    ePrimaryLoadSharedPlayerData
    eCrossPrimaryBudgetAdjust
    eBudgetResponse

    eRequestServerDetails
    eRequestGlobalPlayerScores
    eMineralBid
    eRequestBidStatus
    eSetMineralBid
    eUndockCommandFinished
    eOutBidAlert

    '===================================================
    eLastMsgCode                'MUST ALWAYS BE LAST!!!!
End Enum


Public Enum ObjectType As Short
    eUniverse = 0       'the grandaddy of all objects, this object is intangible but I figured I'd include it here
    eGalaxy             'contained in Universe objects
    eSolarSystem        'contained in Galaxy objects
    ePlanet             'contained in Solar System objects
    eColony             'contained in either a solar system or planet
    'eLandUnit           'contained only on planets or in hangars
    'eSpaceUnit          'contained only in space
    'eAeroUnit           'contained in space or in air of planets
    eAgent              'an agent/spy
    eAlloyTech          'an alloy technology design specification (not the alloy itself)
    eArmorTech          'an armor technology design specification (not the armor itself)
    eCorporation        'a corporation entity, NPC representing the collective interests of a corporation
    eEngineTech         'an Engine technology design specification (not the engine itself)
    eFacility           'an instance of a structure
    eFacilityDef        'definition of a structure's basic properties (not of an actual structure)
    eGNS                'a GNS Body, for example, USA Today would be GNS 1, and New York Times would be GNS 2
    eGoal               'an agent or mission goal
    eHangarTech         'a Hangar technology design specificaton (not the hangar itself)
    eHullTech           'a Hull technology design specification (not the hull itself)
    eMineral            'a mineral definition (not representing a cache of minerals)
    eMineralCache       'a cache of minerals
    eMission            'an agent mission
    eNebula             'contained in galaxy objects
    eObjectOrder        'an order assigned to an object
    ePlayer             'a player object
    ePlayerComm         'a player communication (email)
    ePrototype          'a unit design (not the actual unit)
    eRadarTech          'a radar technology design specification (not the radar itself)
    eSenate             'a senate body
    eSenateLaw          'a senate law
    eShieldTech         'a shield technology design specification (not the shield itself)
    eSkill              'an skill for agents, goals, etc...
    eUnit               'an instance of a unit
    eUnitDef            'a more precise version of a unit's definition
    eUnitGroup          'a group of units (move together, etc...)
    eUnitWeaponDef      'a weapon def to unit def relationship
    eWeaponDef          'a temporary, smaller, more need-to-know information about the weapon
    eWeaponTech         'a weapon technology design specification (not the weapon itself)
    eWormhole           'contained in solar system objects and link two objects together

    eProductionCost     'production cost item...

    'Intangibles...
    ePower              'Power Points
    eCredits            'Galactic Bartering Units
    eStock              'stock certificates
    eMorale             'morale points
    eColonists          'colonists
    eFood               'Food
    eMiningPower        'Mining
    eEnlisted           'General military personnel (infantry and crew are synonomous?)
    eOfficers           'officers required for commanding ships and large armies
    eProductionPoints   'TODO: are we sure of this one???
    eAmmunition
    eMineralProperty
    eMineralTech        'represents mineral researching

    'Component Cache
    eComponentCache     'like a mineral cache but it references a technological component

    eSpecialTech
    eSpecialTechPreq

    eTrade              'a trade object

    eRepairItem         'intangible

    eBudget

    eGuild
    eVoteProposal
    ePlayerIntel

    eGalacticTrade
End Enum


Public Class StateObject
    'Public workSocket As Socket = Nothing
    Public BufferSize As Int32 = 32767
    Public Buffer(32767) As Byte
    'Public sb As New StringBuilder()
End Class

Public Class NetSock
    'This class is inherited by NetSockClient and NetSockServer
    Public Event onConnect(ByVal Index As Int32)
    Public Event onError(ByVal Index As Int32, ByVal Description As String)
    Public Event onDataArrival(ByVal Index As Int32, ByVal Data As Byte(), ByVal TotalBytes As Int32)
    Public Event onDisconnect(ByVal Index As Int32)
    Public Event onSendComplete(ByVal Index As Int32, ByVal DataSize As Int32)
    Public Event onConnectionRequest(ByVal Index As Int32, ByVal oClient As Socket)

    Public SocketIndex As Int32 = 0     'this is important for when doing arrays of these classes

    Public PortNumber As Int32
    Private moListenThread As Threading.Thread
    Private Const ml_LISTEN_CHECK_INTERVAL As Int32 = 2000      'milliseconds for sleeper
    Private mbListening As Boolean

    Private moSocket As Socket      'the actual socket containment

    Public ReadOnly Property IsConnected() As Boolean
        Get
            If moSocket Is Nothing = False Then Return moSocket.Connected
            Return False
        End Get
    End Property

    Public Sub SetLocalBinding(ByVal iPort As Int16)
        ''#If Debug = True Then
        'Exit Sub
        ''#Else
        'If moSocket Is Nothing Then moSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

        'Dim oLocalhost As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName)
        'Dim oIP As IPAddress = oLocalhost.AddressList(0)

        ''moSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ExclusiveAddressUse, False)
        'moSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, 1)
        'moSocket.Bind(New System.Net.IPEndPoint(oIP, iPort))

        'oIP = Nothing
        'oLocalhost = Nothing
        ''#End If

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
            Dim yTemp(Data.Length + 1) As Byte
            System.BitConverter.GetBytes(CShort(Data.Length)).CopyTo(yTemp, 0)
            Data.CopyTo(yTemp, 2)
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
        Dim iLen As Int16
        Dim iOverrideLen As Int16 = -1

        Static xyLeftover() As Byte
        Static xlLeftoverSize As Int32
        Static xlExpectedLen As Int32
        Static xbPartialLenVal As Boolean

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
                    'Check if it is a partial length
                    If xbPartialLenVal = True Then
                        'Ok, this is a specific scenario... xlExpectedLen should equal 1 and we should have 1 byte in xyLeftover
                        Dim yTmp(1) As Byte
                        yTmp(0) = xyLeftover(0)
                        yTmp(1) = yData(lPos)
                        iOverrideLen = System.BitConverter.ToInt16(yTmp, 0)
                    Else
                        ReDim yTemp(xyLeftover.Length + xlExpectedLen)
                        xyLeftover.CopyTo(yTemp, 0)
                        Array.Copy(yData, 0, yTemp, xlLeftoverSize, xlExpectedLen)
                    End If
                    lPos += xlExpectedLen
                    xlExpectedLen = 0
                    xbPartialLenVal = False
                Else
                    If iOverrideLen < 1 AndAlso lBytesRead - (lPos + 2) < 0 Then
                        'Ok, what this is... this is a packet that got cut off at the LEN portion of the message... the length of the message takes two
                        '  bytes, if the len gets cut in half (1 byte in one packet, 1 byte in the other) then this occurs
                        'So what we do is basically the leftover packet deal...
                        xlLeftoverSize = 1
                        ReDim xyLeftover(xlLeftoverSize)
                        xlExpectedLen = 1
                        xyLeftover(0) = yData(lPos)

                        'and we set our boolean indicating partial length
                        xbPartialLenVal = True

                        lPos += 1
                    Else
                        'Check for our overridelen
                        If iOverrideLen < 1 Then
                            iLen = System.BitConverter.ToInt16(yData, lPos)
                            lPos += 2
                        Else
                            iLen = iOverrideLen
                            iOverrideLen = -1
                        End If

                        If (lBytesRead - lPos) < iLen AndAlso iLen <> 0 Then
                            'not enough bytes in this message, set up to wait for next message
                            xlLeftoverSize = lBytesRead - lPos
                            ReDim xyLeftover(xlLeftoverSize)
                            Array.Copy(yData, lPos, xyLeftover, 0, xlLeftoverSize)
                            xlExpectedLen = iLen - xlLeftoverSize
                            lPos += xlLeftoverSize
                        Else
                            'nope, all is here, let's process like normal
                            ReDim yTemp(iLen)
                            Array.Copy(yData, lPos, yTemp, 0, iLen)
                            lPos += iLen
                        End If
                    End If
                End If

                If xlExpectedLen = 0 AndAlso iOverrideLen < 1 Then RaiseEvent onDataArrival(SocketIndex, yTemp, iLen)
            End While
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
        End Try

        'This new Try...Catch block may hurt performance...
        Try
            'Set up our socket to receive the next data arrival
            If moSocket.Connected = True Then moSocket.BeginReceive(oState.Buffer, 0, oState.BufferSize, 0, AddressOf DataArrivalCallback, oState)
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
    Public Function GetLocalDetails() As EndPoint
        If moSocket Is Nothing = False Then Return moSocket.LocalEndPoint Else Return Nothing
    End Function
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
        Dim objResult As New System.Text.StringBuilder(2048)
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
