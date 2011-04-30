Option Strict On

Partial Class MsgSystem
	Private moClients() As NetSock
	Private mlClientUB As Int32 = -1
	Private mlClientPlayer() As Int32
	Private mlAliasedAs() As Int32
	Private mlAliasedRights() As Int32
	Private mlDroppedMsgs() As Int32
	Private mlLastMessageTime() As Int32
	Private mlClientStatusFlags() As Int32

	Private WithEvents moClientListener As NetSock

	Public Property AcceptingClients() As Boolean
		Get
			Return mbAcceptingClients
		End Get
		Set(ByVal Value As Boolean)
			Dim lPort As Int32
			If mbAcceptingClients <> Value Then
				mbAcceptingClients = Value
				If mbAcceptingClients Then
					Dim oINI As New InitFile()
					Try
						lPort = CInt(Val(oINI.GetString("SETTINGS", "ClientListenerPort", "0")))
						If lPort = 0 Then Err.Raise(-1, "AcceptingClients", "Could not get Client Listen Port Number from INI")
						moClientListener = Nothing
						moClientListener = New NetSock()
						moClientListener.PortNumber = lPort
						moClientListener.Listen()
					Catch
						MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
						mbAcceptingClients = False
					Finally
						oINI = Nothing
					End Try
				Else
					'ok, stop listening
					moClientListener.StopListening()
				End If
			End If
		End Set
	End Property

	Public Sub ForceDisconnectClients()
		Dim X As Int32
        Dim yMsg(1) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eServerShutdown).CopyTo(yMsg, 0)
        Dim yTemp(yMsg.Length + 1) As Byte
        System.BitConverter.GetBytes(CShort(yMsg.Length)).CopyTo(yTemp, 0)
        yMsg.CopyTo(yTemp, 2)

		For X = 0 To mlClientUB
            If moClients(X) Is Nothing = False Then
                moClients(X).SendLenAppendedData(yTemp)
                If moClients(X) Is Nothing = False Then moClients(X).Disconnect()
                Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(X))
                If oPlayer Is Nothing = False AndAlso (mlClientStatusFlags(X) And ClientStatusFlags.ePlayerLoggedIn) <> 0 Then oPlayer.TotalPlayTime += CInt(DateDiff(DateInterval.Second, oPlayer.LastLogin, Now))
                moClients(X) = Nothing
                If oPlayer Is Nothing = False Then oPlayer.oSocket = Nothing
            End If
		Next X
	End Sub

#Region "  Client Listener Handling  "
	Private Sub moClientListener_onConnect(ByVal Index As Integer) Handles moClientListener.onConnect
		'
	End Sub

	Private Sub moClientListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moClientListener.onConnectionRequest
		Dim X As Int32
		Dim lIdx As Int32 = -1
        Try
            If AcceptingClients Then
                If bDebug = True AndAlso oClient.RemoteEndPoint.ToString.StartsWith("192.168.1.") = False Then
                    LogEvent(LogEventType.Informational, "Client Connection Denied: Not Accepting Outside Connections")
                    oClient.Shutdown(Net.Sockets.SocketShutdown.Both)
                    Return
                End If

                For X = 0 To mlClientUB
                    If moClients(X) Is Nothing Then
                        lIdx = X
                        Exit For
                    End If
                Next X

                If lIdx = -1 Then
                    mlClientUB += 1
                    ReDim Preserve moClients(mlClientUB)
                    ReDim Preserve mlClientPlayer(mlClientUB)
                    ReDim Preserve mlAliasedAs(mlClientUB)
                    ReDim Preserve mlAliasedRights(mlClientUB)
                    ReDim Preserve mlDroppedMsgs(mlClientUB)
                    ReDim Preserve mlLastMessageTime(mlClientUB)
                    ReDim Preserve mlClientStatusFlags(mlClientUB)
                    lIdx = mlClientUB
                End If

                mlClientPlayer(lIdx) = -1
                moClients(lIdx) = New NetSock(oClient)
                moClients(lIdx).SocketIndex = lIdx
                moClients(lIdx).lAppType = MsgMonitor.eMM_AppType.ClientConnection
                moClients(lIdx).lSpecificID = -1
                mlAliasedAs(lIdx) = -1
                mlAliasedRights(lIdx) = 0
                mlClientStatusFlags(lIdx) = 0

                If moClients(lIdx) Is Nothing Then
                    If oClient Is Nothing = False Then oClient.Shutdown(Net.Sockets.SocketShutdown.Both)
                    Return
                End If
                LogEvent(LogEventType.Informational, "Client Connected")

                'add event handlers
                AddHandler moClients(lIdx).onConnect, AddressOf moClients_onConnect
                AddHandler moClients(lIdx).onDataArrival, AddressOf moClients_onDataArrival
                AddHandler moClients(lIdx).onDisconnect, AddressOf moClients_onDisconnect
                AddHandler moClients(lIdx).onError, AddressOf moClients_onError
                AddHandler moClients(lIdx).onSocketClosed, AddressOf moClients_onSocketClosed

                'and then tell the socket to expect data
                moClients(lIdx).MakeReadyToReceive()
            Else
                LogEvent(LogEventType.Informational, "Client Connection Denied: Not Accepting")
                oClient.Disconnect(True)
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Client.OnConnectionRequest: " & ex.Message)
            Try
                If moClients(lIdx) Is Nothing = False Then moClients(lIdx).Disconnect()
            Catch
            End Try
            moClients(lIdx) = Nothing
        End Try
    End Sub

	Private Sub moClientListener_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moClientListener.onDataArrival
		'
	End Sub

	Private Sub moClientListener_onDisconnect(ByVal Index As Integer) Handles moClientListener.onDisconnect
		'
	End Sub

	Private Sub moClientListener_onError(ByVal Index As Integer, ByVal Description As String) Handles moClientListener.onError
		'
	End Sub

	Private Sub moClientListener_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moClientListener.onSendComplete
		'
	End Sub
#End Region

#Region "  Client Connections Handling  "
	Private Sub moClients_onConnect(ByVal Index As Integer)
		'
	End Sub

	Private Sub moClients_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket)
		'
	End Sub

	Private Sub moClients_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
		'What to do with the message??? hmmm?
		Dim iMsgID As Int16
		Dim X As Int32
		Dim lTemp As Int32

		iMsgID = System.BitConverter.ToInt16(Data, 0)

		If glCurrentCycle - mlLastMessageTime(Index) < 1 Then
			mlDroppedMsgs(Index) += 1
		Else
            If mlDroppedMsgs(Index) > 100 Then
                LogEvent(LogEventType.PossibleCheat, "Spam Attack: " & mlDroppedMsgs(Index) & " msgs, Player: " & mlClientPlayer(Index))
            End If
			mlDroppedMsgs(Index) = 0
			mlLastMessageTime(Index) = glCurrentCycle
		End If

		'MSC - 11/23/2007 - added to stop people that are not logged in from making invalid requests to the server
		If mlClientPlayer(Index) = -1 Then
            If iMsgID <> GlobalMessageCode.eClientVersion AndAlso iMsgID <> GlobalMessageCode.ePlayerLoginRequest AndAlso iMsgID <> GlobalMessageCode.ePathfindingConnectionInfo Then
                Dim oEndPoint As Net.EndPoint = moClients(Index).GetRemoteDetails()
                If oEndPoint Is Nothing = False Then
                    LogEvent(LogEventType.PossibleCheat, "Invalid MsgCode received before login processed: " & iMsgID & ". IP: " & CType(oEndPoint, Net.IPEndPoint).Address.ToString)
                End If
                moClients(Index).Disconnect()
                moClients(Index) = Nothing
                Return
            End If
		End If

		'LogEvent(LogEventType.Informational, "Received Msg From Client " & Index)
		If mb_MONITOR_MSG_ACTIVITY = True Then moMonitor.AddInMsg(iMsgID, MsgMonitor.eMM_AppType.ClientConnection, Data.Length, mlClientPlayer(Index))

        Select Case iMsgID
            Case GlobalMessageCode.ePlayerActivityReport
                HandlePlayerActivityReport(Data, Index)
            Case GlobalMessageCode.eRequestPlayerEnvironment
                'this is a request for what environment a player is in (the player is part of the message)
                'LogEvent(LogEventType.Informational, "Client Requesting Player Environment")
                lTemp = System.BitConverter.ToInt32(Data, 2)
                For X = 0 To glPlayerUB
                    If glPlayerIdx(X) = lTemp Then
                        Dim yTemp() As Byte = GetPlayerEnvirMsg(lTemp, goPlayer(X).lLastViewedEnvir, goPlayer(X).iLastViewedEnvirType)
                        If yTemp Is Nothing = False Then
                            moClients(Index).SendData(yTemp)
                            moClients(Index).SendData(CreateChatMsg(-1, gsMOTD, ChatMessageType.eSysAdminMessage))
                        End If
                        Exit For
                    End If
                Next X
            Case GlobalMessageCode.eRequestEnvironmentDomain
                'a client is asking for which domain does an environment belong to...
                'LogEvent(LogEventType.Informational, "Client Requesting Environment's Domain")
                HandleRequestEnvironmentDomain(Data, moClients(Index))
            Case GlobalMessageCode.eRequestStarTypes
                'LogEvent(LogEventType.Informational, "Client Requesting Star Types")
                HandleRequestStarTypes(Index, ConnectionType.eClientConnection)
            Case GlobalMessageCode.eRequestGalaxyAndSystems
                'LogEvent(LogEventType.Informational, "Client Requesting Galaxy Map")
                HandleRequestGalaxy(Index, ConnectionType.eClientConnection)
            Case GlobalMessageCode.eRequestSystemDetails
                'LogEvent(LogEventType.Informational, "Client Requesting System Details")
                HandleRequestSystemDetails(Index, ConnectionType.eClientConnection, Data)
            Case GlobalMessageCode.eChangingEnvironment
                'LogEvent(LogEventType.Informational, "Client Changing Environment")
                HandleChangeEnvironment(Data, Index)
            Case GlobalMessageCode.eRequestPlayerDetails
                'LogEvent(GlobalVars.LogEventType.Informational, "Client Requesting Player Details")
                HandleRequestPlayerDetails(moClients(Index), Data, Index)
            Case GlobalMessageCode.ePlayerLoginRequest
                'ok contains the temporary player ID Message
                LogEvent(LogEventType.Informational, "Client Requesting Login")
                HandleLoginRequest(Data, Index)
            Case GlobalMessageCode.eSetPlayerRel
                'LogEvent(LogEventType.Informational, "Received Set Player Rel")
                HandleSetPlayerRel(Data, Index)
            Case GlobalMessageCode.eRequestEntityContents
                'LogEvent(LogEventType.Informational, "Received Request Entity Contents")
                HandleRequestEntityContents(Data, moClients(Index))
            Case GlobalMessageCode.eEntityProductionStatus
                'LogEvent(LogEventType.Informational, "Received Entity Production Status")
                Dim oPlayer As Player = Nothing
                If mlAliasedAs(Index) > 0 Then GetEpicaPlayer(mlAliasedAs(Index)) Else oPlayer = GetEpicaPlayer(mlClientPlayer(Index))
                Dim yResp() As Byte = HandleRequestProductionStatus(Data)
                If yResp Is Nothing = False Then moClients(Index).SendData(yResp)
            Case GlobalMessageCode.eSetEntityProduction
                'LogEvent(LogEventType.Informational, "Received Set Entity Production")
                HandleSetEntityProduction(Index, Data)
            Case GlobalMessageCode.eGetEntityProduction
                'LogEvent(LogEventType.Informational, "Received Get Entity Production")
                Dim oPlayer As Player = Nothing
                If mlAliasedAs(Index) > 0 Then GetEpicaPlayer(mlAliasedAs(Index)) Else oPlayer = GetEpicaPlayer(mlClientPlayer(Index))
                'HandleGetEntityProduction(moClients(Index), Data)
                Dim yResp() As Byte = HandleGetEntityProduction(Data)
                If yResp Is Nothing = False Then moClients(Index).SendData(yResp)
            Case GlobalMessageCode.eGetAvailableResources
                'LogEvent(LogEventType.Informational, "Received Get Available Resources")
                HandleGetAvailableResources(Index, Data)
            Case GlobalMessageCode.eChatMessage
                'LogEvent(LogEventType.Informational, "Received Chat Msg")
                HandleChatMsg(Index, Data)
            Case GlobalMessageCode.eRequestSkillList
                'LogEvent(LogEventType.Informational, "Received Request Skill List")
                moClients(Index).SendData(GetSkillListResponse)
            Case GlobalMessageCode.eRequestGoalList
                'LogEvent(LogEventType.Informational, "Received Request Goal List")
                'moClients(Index).SendData(GetGoalListResponse)
            Case GlobalMessageCode.eSetEntityStatus
                'LogEvent(LogEventType.Informational, "Received Client Set Entity Status")
                HandleClientSetEntityStatus(Data, Index)
            Case GlobalMessageCode.eTransferContents
                'LogEvent(LogEventType.Informational, "Received Transfer Contents")
                HandleTransferContents(Data, Index)
            Case GlobalMessageCode.eSetEntityName
                'LogEvent(LogEventType.Informational, "Received Set Entity Name")
                HandleSetEntityName(Data, Index)
            Case GlobalMessageCode.eGetColonyDetails
                'logevent(LogEventType.Informational, "Received Get Colony Details")
                HandleGetColonyDetails(Data, Index)
            Case GlobalMessageCode.eAddPlayerRelObject
                HandleAddPlayerRelMsg(Data, Index)
            Case GlobalMessageCode.eBugList
                'player requesting bug list
                HandleRequestBugList(Index)
            Case GlobalMessageCode.eBugSubmission
                HandleBugSubmission(Data, Index)
            Case GlobalMessageCode.eGetEntityName
                HandleGetEntityName(Data, Index)
            Case GlobalMessageCode.eClientVersion
                HandleClientVersion(Data, Index)
            Case GlobalMessageCode.eSubmitComponentPrototype
                HandleSubmitComponentDesignMsg(Data, Index)
            Case GlobalMessageCode.eSetColonyTaxRate
                HandleSetTaxRate(Data, Index)
            Case GlobalMessageCode.eSetEntityPersonnel
                HandleSetEntityPersonnel(Data, Index)
            Case GlobalMessageCode.eGetColonyChildList
                HandleGetColonyChildList(Data, Index)
            Case GlobalMessageCode.eSetRallyPoint
                HandleSetRallyPoint(Data, Index)
            Case GlobalMessageCode.eAddToFleet
                HandleAddToFleet(Data, Index)
            Case GlobalMessageCode.eCreateFleet
                HandleCreateFleet(Data, Index)
            Case GlobalMessageCode.eDeleteFleet
                HandleDeleteFleet(Data, Index)
            Case GlobalMessageCode.eRemoveFromFleet
                HandleRemoveFromFleet(Data, Index)
            Case GlobalMessageCode.eSetFleetDest
                HandleSetFleetDest(Data, Index)
            Case GlobalMessageCode.eAddPlayerCommFolder
                HandleAddPlayerCommFolder(Data, Index)
            Case GlobalMessageCode.eSendEmail, GlobalMessageCode.eSaveEmailDraft
                HandleSendEmail(Data, Index)
            Case GlobalMessageCode.eDeleteEmailItem
                HandleDeleteEmailItem(Data, Index)
            Case GlobalMessageCode.eMoveEmailToFolder
                HandleMoveEmailToFolder(Data, Index)
            Case GlobalMessageCode.eMarkEmailReadStatus
                HandleMarkEmailReadStatus(Data, Index)
            Case GlobalMessageCode.eSubmitTrade
                HandleSubmitTrade(Data, Index)
            Case GlobalMessageCode.eGetColonyList
                HandleGetColonyList(Data, Index)
            Case GlobalMessageCode.eGetNonOwnerDetails
                HandleGetNonOwnerDetails(Data, Index)
            Case GlobalMessageCode.eRequestEntityDefenses
                HandleRequestEntityDefenses(Data, Index)
            Case GlobalMessageCode.eRepairOrder
                HandleRepairOrder(Data, Index)
            Case GlobalMessageCode.eRequestPlayerBudget
                HandleRequestPlayerBudget(Data, Index)
            Case GlobalMessageCode.eGetPlayerScores
                HandleGetPlayerScores(Data, Index)
            Case GlobalMessageCode.eGetEnvirConstructionObjects
                HandleGetEnvirConstructionObjects(Data, Index)
            Case GlobalMessageCode.ePlayerInitialSetup
                HandlePlayerInitialSetup(Data, Index)
            Case GlobalMessageCode.eDeathBudgetDeposit
                HandleDeathDeposit(Data, Index)
            Case GlobalMessageCode.ePlayerIsDead
                HandlePlayerIsDead(Data, Index)
            Case GlobalMessageCode.eGetEntityDetails
                HandleGetEntityDetails(Data, Index)
            Case GlobalMessageCode.eDeleteDesign
                HandleDeleteDesign(Data, Index)
            Case GlobalMessageCode.eRequestEntityAmmo
                HandleRequestEntityAmmo(Data, Index)
            Case GlobalMessageCode.eRequestLoadAmmo
                HandleRequestLoadAmmo(Data, Index)
            Case GlobalMessageCode.eEmailSettings
                HandleEmailSettings(Data, Index)
            Case GlobalMessageCode.eGetGTCList
                If goGTC Is Nothing = False Then
                    If mlAliasedAs(Index) > 0 Then
                        If (mlAliasedRights(Index) And AliasingRights.eViewTrades) <> 0 Then
                            goGTC.HandleGetGTCList(Data, mlAliasedAs(Index))
                        End If
                    Else : goGTC.HandleGetGTCList(Data, mlClientPlayer(Index))
                    End If
                End If
            Case GlobalMessageCode.ePurchaseSellOrder
                If goGTC Is Nothing = False Then goGTC.HandlePurchaseSellOrder(Data)
            Case GlobalMessageCode.eSubmitSellOrder
                If goGTC Is Nothing = False Then
                    If goGTC.HandleSubmitSellOrder(Data) = False Then
                        'TODO: Return a reason
                    End If
                End If
            Case GlobalMessageCode.eGetTradePostList
                HandleGetTradePostList(Data, Index)
            Case GlobalMessageCode.eGetTradePostTradeables
                HandleGetTradePostTradeables(Data, Index)
            Case GlobalMessageCode.eGetOrderSpecifics
                HandleGetOrderSpecifics(Data, Index)
            Case GlobalMessageCode.eGetTradeHistory
                HandleGetTradeHistory(Data, Index)
            Case GlobalMessageCode.eSubmitBuyOrder
                If goGTC Is Nothing = False Then
                    If mlAliasedAs(Index) > 0 Then
                        If (mlAliasedRights(Index) And AliasingRights.eAlterTrades) <> 0 Then
                            goGTC.HandleSubmitBuyOrder(Data, mlAliasedAs(Index))
                        End If
                    Else : goGTC.HandleSubmitBuyOrder(Data, mlClientPlayer(Index))
                    End If
                End If
            Case GlobalMessageCode.eAcceptBuyOrder
                If goGTC Is Nothing = False Then
                    If mlAliasedAs(Index) > 0 Then
                        If (mlAliasedRights(Index) And AliasingRights.eAlterTrades) <> 0 Then
                            goGTC.HandleAcceptBuyOrder(Data, mlAliasedAs(Index))
                        End If
                    Else : goGTC.HandleAcceptBuyOrder(Data, mlClientPlayer(Index))
                    End If
                End If
            Case GlobalMessageCode.eDeliverBuyOrder
                If goGTC Is Nothing = False Then
                    If mlAliasedAs(Index) > 0 Then
                        If (mlAliasedRights(Index) And AliasingRights.eAlterTrades) <> 0 Then
                            goGTC.HandleDeliverBuyOrder(Data, mlAliasedAs(Index))
                        End If
                    Else : goGTC.HandleDeliverBuyOrder(Data, mlClientPlayer(Index))
                    End If
                End If
            Case GlobalMessageCode.eGetTradeDeliveries
                HandleGetTradeDeliveries(Data, Index)
            Case GlobalMessageCode.eSetFleetReinforcer
                HandleSetFleetReinforcer(Data, Index)
            Case GlobalMessageCode.ePlayerAliasConfig
                HandlePlayerAliasConfig(Data, Index)
            Case GlobalMessageCode.eDeinfiltrateAgent
                HandleDeinfiltrateAgent(Data, Index)
            Case GlobalMessageCode.eSetInfiltrateSettings
                HandleSetInfiltrateSettings(Data, Index)
            Case GlobalMessageCode.eSetAgentStatus
                HandleSetAgentStatus(Data, Index)
            Case GlobalMessageCode.eSubmitMission
                HandleSubmitMission(Data, Index)
            Case GlobalMessageCode.eGetAgentStatus
                HandleGetAgentStatus(Data, Index)
            Case GlobalMessageCode.eGetPMUpdate
                HandleGetPMUpdate(Data, Index)
            Case GlobalMessageCode.eSetSkipStatus
                HandleSetSkipStatus(Data, Index)
            Case GlobalMessageCode.eRequestMineral
                HandleRequestMineral(Data, Index)
            Case GlobalMessageCode.eRequestEmailSummarys
                HandleRequestEmailSummarys(Data, Index)
            Case GlobalMessageCode.eRequestEmail
                HandleRequestEmail(Data, Index)
            Case GlobalMessageCode.eRequestAliasConfigs
                HandleRequestAliasConfigs(Data, Index)
            Case GlobalMessageCode.eRequestDXDiag
                HandleRequestDXDiag(Data, Index)
            Case GlobalMessageCode.eGetRouteList
                HandleGetRouteList(Data, Index)
            Case GlobalMessageCode.eSetRouteMineral
                HandleSetRouteMineral(Data, Index)
            Case GlobalMessageCode.eUpdateRouteStatus
                HandleUpdateRouteStatus(Data, Index)
            Case GlobalMessageCode.eRemoveRouteItem
                HandleRemoveRouteItem(Data, Index)
            Case GlobalMessageCode.eGetSkillList
                HandleGetSkillList(Data, Index)
            Case GlobalMessageCode.eAddFormation
                HandleAddFormation(Data, Index)
            Case GlobalMessageCode.eRemoveFormation
                HandleRemoveFormation(Data, Index)
            Case GlobalMessageCode.eSetIronCurtain
                HandleSetIronCurtain(Data, Index)
            Case GlobalMessageCode.eUpdateSlotStates
                HandleUpdateSlotStates(Data, Index)
            Case GlobalMessageCode.eCaptureKillAgent
                HandleCaptureKillAgent(Data, Index)
            Case GlobalMessageCode.eGetItemIntelDetail
                HandleGetItemIntelDetail(Data, Index)
            Case GlobalMessageCode.eGetIntelSellOrderDetail
                HandleGetIntelSellOrderDetail(Data, Index)
            Case GlobalMessageCode.eGetArchivedItems
                HandleGetArchivedItems(Data, Index)
            Case GlobalMessageCode.eSetArchiveState
                HandleSetArchiveState(Data, Index)
            Case GlobalMessageCode.eGetColonyResearchList
                HandleGetColonyResearchList(Data, Index)
            Case GlobalMessageCode.eSetColonyResearchQueue
                HandleSetColonyResearchQueue(Data, Index)
            Case GlobalMessageCode.eTutorialProdFinish
                HandleTutorialProdFinish(Data, Index)
            Case GlobalMessageCode.eUpdatePlayerTutorialStep
                HandleUpdatePlayerTutorialStep(Data, Index)
            Case GlobalMessageCode.eUpdateMOTD
                HandleUpdateMOTD(Data, Index)
            Case GlobalMessageCode.eGuildBankTransaction
                HandleGuildBankTrans(Data, Index)
            Case GlobalMessageCode.eGuildRequestDetails
                HandleGuildRequestDetails(Data, Index)
            Case GlobalMessageCode.eUpdatePlayerVote
                HandleUpdatePlayerVote(Data, Index)
            Case GlobalMessageCode.eUpdateGuildRank
                HandleUpdateGuildRank(Data, Index)
            Case GlobalMessageCode.eUpdateRankPermission
                HandleUpdateRankPermission(Data, Index)
            Case GlobalMessageCode.eGuildMemberStatus
                HandleGuildMemberStatus(Data, Index)
            Case GlobalMessageCode.eInvitePlayerToGuild
                HandleInvitePlayerToGuild(Data, Index)
            Case GlobalMessageCode.eUpdateGuildEvent
                HandleUpdateGuildEvent(Data, Index)
            Case GlobalMessageCode.eProposeGuildVote
                HandleProposeGuildVote(Data, Index)
            Case GlobalMessageCode.eSearchGuilds
                HandleSearchGuilds(Data, Index)
            Case GlobalMessageCode.eRequestContactWithGuild
                HandleRequestContactWithGuild(Data, Index)
            Case GlobalMessageCode.eSetGuildRel
                HandleSetGuildRel(Data, Index)
            Case GlobalMessageCode.eAddEventAttachment
                HandleAddEventAttachment(Data, Index)
            Case GlobalMessageCode.eUpdateGuildRecruitment
                HandleUpdateGuildRecruitment(Data, Index)
            Case GlobalMessageCode.eRemoveEventAttachment
                HandleRemoveEventAttachment(Data, Index)
            Case GlobalMessageCode.eRequestGuildVoteProposals
                HandleRequestGuildVoteProposals(Data, Index)
            Case GlobalMessageCode.eAdvancedEventConfig
                HandleAdvancedEventConfig(Data, Index)
            Case GlobalMessageCode.eCheckGuildName
                HandleCheckGuildName(Data, Index)
            Case GlobalMessageCode.eCreateGuild
                HandleCreateGuild(Data, Index)
            Case GlobalMessageCode.eInviteFormGuild
                HandleInviteFormGuild(Data, Index)
            Case GlobalMessageCode.eGetGuildInvites
                HandleGetGuildInvites(Data, Index)
            Case GlobalMessageCode.eGetMyVoteValue
                HandleGetMyVoteValue(Data, Index)
            Case GlobalMessageCode.eRequestGuildEvents
                HandleRequestGuildEvents(Data, Index)
            Case GlobalMessageCode.eUpdateGuildTreasury
                HandleUpdateGuildTreasury(Data, Index)
            Case GlobalMessageCode.eGetGuildAssets
                HandleGetGuildAssets(Data, Index)
            Case GlobalMessageCode.eAddSenateProposal
                'Dim yResp() As Byte = Senate.HandleCreateProposalMsg(Data, mlClientPlayer(Index))
                'If yResp Is Nothing = False Then moClients(Index).SendData(yResp)
                System.BitConverter.GetBytes(mlClientPlayer(Index)).CopyTo(Data, 2)
                moOperator.SendData(Data)
            Case GlobalMessageCode.eAddSenateProposalMessage
                'Senate.HandleAddProposalMessage(Data, mlClientPlayer(Index))
                System.BitConverter.GetBytes(mlClientPlayer(Index)).CopyTo(Data, 2)
                moOperator.SendData(Data)
            Case GlobalMessageCode.eGetSenateObjectDetails
                System.BitConverter.GetBytes(mlClientPlayer(Index)).CopyTo(Data, 2)
                moOperator.SendData(Data)
            Case GlobalMessageCode.eGuildCreationUpdate
                HandleGuildCreationUpdate(Data, Index)
            Case GlobalMessageCode.eGuildCreationAcceptance
                HandleGuildCreationAcceptance(Data, Index)
            Case GlobalMessageCode.eSetDismantleTarget
                HandleClientDismantleObject(Data, Index)        'be sure to delete DeleteStationChild code block as it is obsolete
            Case GlobalMessageCode.ePlayerRestart
                HandlePlayerRestart(Data, Index)
            Case GlobalMessageCode.eRequestChannelDetails
                HandleRequestChannelDetails(Data, Index)
            Case GlobalMessageCode.eRequestChannelList
                'Dim yResp() As Byte = HandleRequestChannelList(mlClientPlayer(Index))
                'If yResp Is Nothing = False Then moClients(Index).SendData(yResp)
                HandleRequestChannelList(Data, Index)
            Case GlobalMessageCode.eDeleteTradeHistoryItem
                HandleDeleteTradeHistoryItem(Data, Index)
            Case GlobalMessageCode.eSetFleetFormation
                HandleSetFleetFormation(Data, Index)
            Case GlobalMessageCode.eGetImposedAgentEffects
                HandleGetImposedAgentEffects(Data, Index)
            Case GlobalMessageCode.eUpdatePlayerCustomTitle
                HandleUpdatePlayerCustomTitle(Data, Index)
            Case GlobalMessageCode.eAbandonColony
                HandleAbandonColony(Data, Index)
            Case GlobalMessageCode.eClearBudgetAlert
                HandleClearBudgetAlert(Data, Index)
            Case GlobalMessageCode.eTutorialGiveCredits
                HandleTutorialGiveCredits(Data, Index)
            Case GlobalMessageCode.eChatChannelCommand
                HandleChatChannelCommand(Data, Index)
            Case GlobalMessageCode.eRequestBidStatus
                HandleRequestBidStatus(Data, Index)
            Case GlobalMessageCode.eSetMineralBid
                HandleSetMineralBid(Data, Index)
            Case GlobalMessageCode.eMineralBid
                HandleMineralBidMsg(Data, Index)
            Case GlobalMessageCode.eRestartTutorial
                HandleRestartPlayerTutorial(Data, Index)
            Case GlobalMessageCode.eClearEntityProdQueue
                HandleClearUnitProdQueue(Data, Index)
            Case GlobalMessageCode.eUpdatePlanetOwnership
                HandleRequestPlanetOwnership(Data, Index)
            Case GlobalMessageCode.eSetSpecialTechThrowback
                HandleSetSpecialTechThrowback(Data, Index)
            Case GlobalMessageCode.eGetGuildBillboards
                HandleGetGuildBillboards(Data, Index)
            Case GlobalMessageCode.ePlaceGuildBillboardBid
                HandlePlaceGuildBillboardBid(Data, Index)
            Case GlobalMessageCode.eGetSpecTechGuaranteeList
                HandleGetSpecialTechGuaranteeList(Data, Index)
            Case GlobalMessageCode.eSetSpecTechGuarantee
                HandleSetSpecTechGuarantee(Data, Index)
            Case GlobalMessageCode.eForcefulDismantle
                HandleForcefulDismantle(Data, Index)
            Case GlobalMessageCode.eClaimItem
                HandleClaimItem(Data, Index)
            Case GlobalMessageCode.eUpdatePlayerDetails
                Dim lPlayerID As Int32 = mlClientPlayer(Index)
                If mlAliasedAs(Index) > 0 AndAlso mlAliasedAs(Index) <> lPlayerID Then Return
                Dim lIcon As Int32 = System.BitConverter.ToInt32(Data, 2)
                If lIcon <> 0 Then
                    Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                    If oPlayer Is Nothing = False Then
                        oPlayer.lPlayerIcon = lIcon
                        oPlayer.blCredits -= 1000000
                        oPlayer.SaveObject(True)
                    End If
                End If
            Case GlobalMessageCode.eRequestObject
                HandleClientRequestObject(Data, Index)
            Case GlobalMessageCode.eGetCPPenaltyList
                HandleGetCPPenaltyList(Data, Index)
            Case GlobalMessageCode.eEditRouteTemplate
                HandleEditRouteTemplate(Data, Index)
            Case GlobalMessageCode.eGetRouteTemplates
                HandleGetRouteTemplates(Data, Index)
            Case GlobalMessageCode.eApplyRouteToEntities
                HandleApplyRouteToEntities(Data, Index)
            Case GlobalMessageCode.eBeginRouteForEntities
                HandleApplyRouteToEntities(Data, Index)
            Case GlobalMessageCode.eExceptionReport
                HandleExceptionReport(Data, Index)
            Case GlobalMessageCode.eSenateStatusReport  '293
                System.BitConverter.GetBytes(mlClientPlayer(Index)).CopyTo(Data, 2)
                SendMsgToOperator(Data)
            Case GlobalMessageCode.eRequestMaxWpnRng
                HandleRequestMaxWpnRng(Data, Index)
            Case GlobalMessageCode.eAddTransport
                HandleAddTransport(Data, Index)
            Case GlobalMessageCode.eDecommissionTransport
                HandleDecommissionTransport(Data, Index)
            Case GlobalMessageCode.eRequestColonyTransportPotentials
                HandleRequestColonyTransportPotentials(Data, Index)
            Case GlobalMessageCode.eRequestTransports
                HandleRequestTransports(Data, Index)
            Case GlobalMessageCode.eRequestTransportName
                HandleRequestTransportName(Data, Index)
            Case GlobalMessageCode.eRequestTransportDetails
                HandleRequestTransportDetails(Data, Index)
            Case GlobalMessageCode.eRequestTransportRouteDestList
                HandleRequestTransportRouteDestList(Data, Index)
            Case GlobalMessageCode.eSetTransportStatus
                HandleSetTransportStatus(Data, Index)
            Case GlobalMessageCode.eGetDAValues
                HandleGetDAValues(Data, Index)
            Case GlobalMessageCode.eRespawnSelection
                HandleRespawnSelection(Data, Index)
                'Case GlobalMessageCode.eWPTimers
                '    HandleClientRequestWPTimers(Data, Index)
            Case GlobalMessageCode.eGetGuildShareAssets
                HandleClientRequestGuildSharedAssets(Data, Index)
            Case GlobalMessageCode.eRequestUniversalAssets
                HandleClientRequestUniversalAssets(Data, Index)
        End Select
	End Sub

	Private Sub moClients_onDisconnect(ByVal Index As Integer)
		Dim X As Int32
		Dim sPlayerName As String = ""
		For X = 0 To glPlayerUB
			If glPlayerIdx(X) <> -1 Then
				If goPlayer(X).oSocket Is Nothing = False Then
					If goPlayer(X).oSocket.SocketIndex = Index Then
						sPlayerName = BytesToString(goPlayer(X).PlayerName)
						goPlayer(X).oSocket = Nothing
                        goPlayer(X).lAliasingPlayerID = -1

                        If (mlClientStatusFlags(Index) And ClientStatusFlags.ePlayerLoggedIn) <> 0 Then goPlayer(X).TotalPlayTime += CInt(DateDiff(DateInterval.Second, goPlayer(X).LastLogin, Now))

                        Dim yOpMsg(9) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerConnectedPrimary).CopyTo(yOpMsg, 0)
                        System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yOpMsg, 2)
                        System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yOpMsg, 6)
                        moOperator.SendData(yOpMsg)
                        BroadcastToDomains(yOpMsg)

                        'tell our email/chat server that the player is offline
                        Dim yEmail(10) As Byte
                        Dim lAID As Int32 = mlAliasedAs(Index)
                        If lAID < 1 OrElse lAID = mlClientPlayer(Index) Then lAID = -1
                        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yEmail, 0)
                        System.BitConverter.GetBytes(mlClientPlayer(Index)).CopyTo(yEmail, 2)
                        System.BitConverter.GetBytes(lAID).CopyTo(yEmail, 6)
                        yEmail(10) = 0
                        SendToEmailSrvr(yEmail)                        'moEmailSrvr.SendData(yEmail)

						Dim lActualPlayerID As Int32 = mlClientPlayer(Index)
						If mlAliasedAs(Index) > -1 Then lActualPlayerID = mlAliasedAs(Index)
						CancelIronCurtainEvents(lActualPlayerID)
						If goPlayer(X).yPlayerPhase = eyPlayerPhase.eInitialPhase Then
							AddToQueue(glCurrentCycle + 30, QueueItemType.eSaveAndUnloadInstancedPlanet, lActualPlayerID, 0, 0, 0, 0, 0, 0, 0)
						Else
                            If goPlayer(X).yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso goPlayer(X).AccountStatus = AccountStatusType.eActiveAccount Then
                                goPlayer(X).yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnSelectedPlanet
                                AddToQueue(glCurrentCycle + 108000, QueueItemType.eIronCurtainRaise, lActualPlayerID, 0, 0, 0, 0, 0, 0, 0)
                            End If
						End If
						Exit For
					End If
				End If
			End If
		Next X
		moClients(Index) = Nothing

        'If sPlayerName.Length <> 0 Then
        '	Dim yLogOff() As Byte = CreateChatMsg(-1, sPlayerName & " has logged off.", ChatMessageType.eSysAdminMessage)
        '	For X = 0 To glPlayerUB
        '		If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).oSocket Is Nothing = False Then goPlayer(X).oSocket.SendData(yLogOff)
        '	Next X
        'End If

		LogEvent(LogEventType.Informational, "Client " & Index & " Disconnected")
	End Sub

	Private Sub moClients_onError(ByVal Index As Integer, ByVal Description As String)

		On Error Resume Next
		LogEvent(LogEventType.Informational, "Client Connection Error (" & Index & "): " & Description)
		If InStr(Description.ToLower, "an existing connection was forcibly closed by the remote host", CompareMethod.Binary) <> 0 Then
			moClients(Index).Disconnect()
			If mlClientPlayer(Index) > 0 Then
				Dim oP As Player = GetEpicaPlayer(mlClientPlayer(Index))
				If oP Is Nothing = False Then
					oP.oSocket = Nothing
                    If (mlClientStatusFlags(Index) And ClientStatusFlags.ePlayerLoggedIn) <> 0 Then oP.TotalPlayTime += CInt(DateDiff(DateInterval.Second, oP.LastLogin, Now)) 'oP.TotalPlayTime += CInt(DateDiff(DateInterval.Second, oP.LastLogin, Now))

                    Dim yOpMsg(9) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerConnectedPrimary).CopyTo(yOpMsg, 0)
                    System.BitConverter.GetBytes(oP.ObjectID).CopyTo(yOpMsg, 2)
                    System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yOpMsg, 6)
                    moOperator.SendData(yOpMsg)
                    BroadcastToDomains(yOpMsg)

                    'tell our email/chat server that the player is offline
                    Dim yEmail(10) As Byte
                    Dim lAID As Int32 = mlAliasedAs(Index)
                    If lAID < 1 OrElse lAID = mlClientPlayer(Index) Then lAID = -1
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yEmail, 0)
                    System.BitConverter.GetBytes(mlClientPlayer(Index)).CopyTo(yEmail, 2)
                    System.BitConverter.GetBytes(lAID).CopyTo(yEmail, 6)
                    yEmail(10) = 0
                    SendToEmailSrvr(yEmail) 'moEmailSrvr.SendData(yEmail)

					Dim lActualPlayerID As Int32 = mlClientPlayer(Index)
					If mlAliasedAs(Index) > -1 Then lActualPlayerID = mlAliasedAs(Index)
					CancelIronCurtainEvents(lActualPlayerID)
					If oP.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
						AddToQueue(glCurrentCycle + 30, QueueItemType.eSaveAndUnloadInstancedPlanet, lActualPlayerID, 0, 0, 0, 0, 0, 0, 0)
					Else
                        If oP.yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso oP.AccountStatus = AccountStatusType.eActiveAccount Then
                            oP.yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnSelectedPlanet
                            AddToQueue(glCurrentCycle + 108000, QueueItemType.eIronCurtainRaise, lActualPlayerID, 0, 0, 0, 0, 0, 0, 0)
                        End If
					End If
				End If
			End If
			moClients(Index) = Nothing
		ElseIf Err.Number = 5 Then
			If moClients(Index) Is Nothing = False Then moClients(Index).Disconnect()
			If mlClientPlayer(Index) > 0 Then
				Dim oP As Player = GetEpicaPlayer(mlClientPlayer(Index))
				If oP Is Nothing = False Then
					oP.oSocket = Nothing
                    If (mlClientStatusFlags(Index) And ClientStatusFlags.ePlayerLoggedIn) <> 0 Then oP.TotalPlayTime += CInt(DateDiff(DateInterval.Second, oP.LastLogin, Now)) 'oP.TotalPlayTime += CInt(DateDiff(DateInterval.Second, oP.LastLogin, Now))

                    Dim yOpMsg(9) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerConnectedPrimary).CopyTo(yOpMsg, 0)
                    System.BitConverter.GetBytes(oP.ObjectID).CopyTo(yOpMsg, 2)
                    System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yOpMsg, 6)
                    moOperator.SendData(yOpMsg)
                    BroadcastToDomains(yOpMsg)

                    'tell our email/chat server that the player is offline
                    Dim yEmail(10) As Byte
                    Dim lAID As Int32 = mlAliasedAs(Index)
                    If lAID < 1 OrElse lAID = mlClientPlayer(Index) Then lAID = -1
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yEmail, 0)
                    System.BitConverter.GetBytes(mlClientPlayer(Index)).CopyTo(yEmail, 2)
                    System.BitConverter.GetBytes(lAID).CopyTo(yEmail, 6)
                    yEmail(10) = 0
                    SendToEmailSrvr(yEmail) 'moEmailSrvr.SendData(yEmail)

					Dim lActualPlayerID As Int32 = mlClientPlayer(Index)
					If mlAliasedAs(Index) > -1 Then lActualPlayerID = mlAliasedAs(Index)
					CancelIronCurtainEvents(lActualPlayerID)
					If oP.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
						AddToQueue(glCurrentCycle + 30, QueueItemType.eSaveAndUnloadInstancedPlanet, lActualPlayerID, 0, 0, 0, 0, 0, 0, 0)
					Else
                        If oP.yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso oP.AccountStatus = AccountStatusType.eActiveAccount Then
                            oP.yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnSelectedPlanet
                            AddToQueue(glCurrentCycle + 108000, QueueItemType.eIronCurtainRaise, lActualPlayerID, 0, 0, 0, 0, 0, 0, 0)
                        End If
					End If
				End If
			End If
			moClients(Index) = Nothing
		End If
	End Sub

    Private Sub moClients_onSocketClosed(ByVal lIndex As Int32)
        moClients(lIndex) = Nothing
    End Sub
#End Region

	Private Function DecBytes(ByVal yBytes() As Byte) As Byte()
		Const ml_ENCRYPT_SEED As Int32 = 777

		'Now, we do the exact opposite...
		Dim lLen As Int32 = yBytes.GetUpperBound(0)
		Dim lKey As Int32
		'Dim lOffset As Int32
		Dim X As Int32
		Dim yFinal(lLen - 1) As Byte
		Dim lChrCode As Int32
		Dim lMod As Int32

		'Get our key value...
		lKey = yBytes(0)

		'set up our seed
		Rnd(-1)
		Call Randomize(ml_ENCRYPT_SEED + lKey)

		For X = 1 To lLen
			'Now, find out what we got here...
			lChrCode = yBytes(X)
			'now, subtract our value... 1 to 5
			lMod = CInt(Int(Rnd() * 5) + 1)
			lChrCode = lChrCode - lMod
			If lChrCode < 0 Then lChrCode = 256 + lChrCode
			yFinal(X - 1) = CByte(lChrCode)
		Next X
		DecBytes = yFinal
	End Function

#Region "  Message Handlers  "
    Private Sub HandleAbandonColony(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim lPos As Int32 = 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)
            If lColonyIdx > -1 Then
                If glColonyIdx(lColonyIdx) > -1 Then
                    Dim oColony As Colony = goColony(lColonyIdx)
                    If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = lPlayerID Then

                        LogEvent(LogEventType.Informational, BytesToString(oPlayer.PlayerName) & " is abandoning colony ID " & oColony.ObjectID & " on " & lEnvirID & " of " & iEnvirTypeID)

                        If iEnvirTypeID = ObjectType.eFacility Then
                            Dim oParentFac As Facility = CType(oColony.ParentObject, Facility)
                            If oParentFac Is Nothing = False Then
                                'ok, special case... delete the colony object
                                oColony.DeleteColony(Colony.ColonyLostReason.Abandoned)
                                'clear the parent facility's parent colony
                                CType(oColony.ParentObject, Facility).ParentColony = Nothing
                                'now, destroy all children facilities
                                oColony.DestroyAllChildrenFacilities()
                                If oParentFac.ServerIndex > -1 Then
                                    If glFacilityIdx(oParentFac.ServerIndex) = oParentFac.ObjectID Then
                                        glFacilityIdx(oParentFac.ServerIndex) = -1
                                        RemoveLookupFacility(oParentFac.ObjectID, oParentFac.ObjTypeID)
                                        goFacility(oParentFac.ServerIndex) = Nothing
                                    End If
                                    oParentFac.DeleteEntity(oParentFac.ServerIndex)
                                End If
                            End If
                        End If
                        oColony.Population = 0
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub HandleAddEventAttachment(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2    'for msgcode
        Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEvent As GuildEvent = oPlayer.oGuild.GetEvent(lEventID)
        If oEvent Is Nothing Then Return

        Dim yAttachPos As Byte = yData(lPos) : lPos += 1

        Dim lIDPos As Int32 = lPos
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oSketchPad As SketchPad = Nothing
        If lID > 0 Then oSketchPad = oEvent.GetAttachment(lID)
        If oSketchPad Is Nothing Then
            oSketchPad = New SketchPad()
        End If

        With oSketchPad
            .lID = lID
            ReDim .yName(19)
            Array.Copy(yData, lPos, .yName, 0, 20) : lPos += 20
            .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .ViewID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraAtX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraAtY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraAtZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lItemCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            For X As Int32 = 0 To lItemCnt - 1
                Dim yType As Byte = yData(lPos) : lPos += 1
                Dim yClrVal As Byte = yData(lPos) : lPos += 1
                Dim fPtAX As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim fPtAY As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim fPtBX As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim fPtBY As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                Dim sText As String = ""
                If yType = SketchPad.eySketchShapes.Text Then
                    Dim lTextLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    sText = GetStringFromBytes(yData, lPos, lTextLen) : lPos += lTextLen
                End If
                .AddSketchPadItem(yType, fPtAX, fPtAY, fPtBX, fPtBY, yClrVal, sText)
            Next X
            If .SaveObject(oEvent.EventID) = False Then Return
        End With
        oEvent.AddOrSetAttachment(oSketchPad)

        System.BitConverter.GetBytes(oSketchPad.lID).CopyTo(yData, lIDPos)

        oPlayer.oGuild.SendMsgToGuildMembers(yData)

    End Sub
	Private Sub HandleAddFormation(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2	'for msgcode

		Dim lFormationID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim oFormation As FormationDef = Nothing

		If lFormationID > 0 Then
			'change/update... find out if the formation belongs to the player
			oFormation = GetEpicaFormation(lFormationID)
			If oFormation Is Nothing = False Then
				If oFormation.lOwnerID <> mlClientPlayer(lIndex) Then
					LogEvent(LogEventType.PossibleCheat, "Owner of formation is not correct: " & mlClientPlayer(lIndex))
					Return
				Else
					oFormation.lOwnerID = mlClientPlayer(lIndex)
                    If oFormation.FillFromMsg(yData, False) = False Then Return
				End If
			Else : Return
			End If
		Else
			'insert
			oFormation = New FormationDef()
			oFormation.lOwnerID = mlClientPlayer(lIndex)
            If oFormation.FillFromMsg(yData, False) = False Then Return
		End If
		If oFormation Is Nothing Then Return
		If oFormation.SaveObject() = False Then Return

		If lFormationID < 1 Then
			'need to add it to the global arrays and such
            Dim lIdx As Int32 = AddFormationToGlobalArray(oFormation)
		End If

		Dim yForward() As Byte = oFormation.GetAsAddMsg()

		'ok, we need to send this to all domains, pathfinding and back at the client...

        'need to inform other primarys too via operator
        moOperator.SendData(yForward)

		For X As Int32 = 0 To mlDomainUB
			moDomains(X).SendData(yForward)
		Next X
		moPathfinding.SendData(yForward)
		moClients(lIndex).SendData(yForward)
	End Sub
	Private Sub HandleAddPlayerCommFolder(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		If mlClientPlayer(lIndex) <> lPlayerID AndAlso (mlAliasedAs(lIndex) <> lPlayerID OrElse (mlAliasedRights(lIndex) And AliasingRights.eAlterEmail) = 0) Then
			LogEvent(LogEventType.PossibleCheat, "Add Player Comm Folder, Player ID doesn't match!")
			Return
		End If

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

        If oPlayer Is Nothing = False Then

            If oPlayer.InMyDomain = False Then
                SendPassThruMsg(yData, lPlayerID, lPlayerID, ObjectType.ePlayer)
                Return
            End If

            Dim yResp() As Byte = DoAddPlayerCommFolder(oPlayer, yData, 0)
            If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
        End If
	End Sub
	Private Sub HandleAddPlayerRelMsg(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim lOther As Int32 = System.BitConverter.ToInt32(yData, 6)
		Dim X As Int32
		Dim oPlayer1 As Player = Nothing
		Dim oPlayer2 As Player = Nothing
		Dim oTmpRel_A_to_B As PlayerRel = Nothing
		Dim oTmpRel_B_to_A As PlayerRel = Nothing

		Dim bAdd_A_to_B As Boolean = False
		Dim bAdd_B_to_A As Boolean = False

		Dim yMsg_A_to_B() As Byte = Nothing
		Dim yMsg_B_to_A() As Byte = Nothing

		'TODO: check if the player owning the socket index is the right player

		For X = 0 To glPlayerUB
			If glPlayerIdx(X) = lPlayerID Then
				oPlayer1 = goPlayer(X)
			ElseIf glPlayerIdx(X) = lOther Then
				oPlayer2 = goPlayer(X)
			End If
			If oPlayer1 Is Nothing = False AndAlso oPlayer2 Is Nothing = False Then Exit For
		Next X

		If oPlayer1 Is Nothing OrElse oPlayer2 Is Nothing Then Return

		oTmpRel_A_to_B = oPlayer1.GetPlayerRel(lOther)
		oTmpRel_B_to_A = oPlayer2.GetPlayerRel(lPlayerID)

		If oTmpRel_A_to_B Is Nothing Then
			oTmpRel_A_to_B = New PlayerRel()
			bAdd_A_to_B = True
			With oTmpRel_A_to_B
				.oPlayerRegards = oPlayer1
				.oThisPlayer = oPlayer2
				.WithThisScore = elRelTypes.eNeutral
			End With
            oPlayer1.SetPlayerRel(lOther, oTmpRel_A_to_B, True)

			yMsg_A_to_B = GetAddPlayerRelMessage(oTmpRel_A_to_B)

			oPlayer1.SendPlayerMessage(yMsg_A_to_B, False, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)
		End If

		If oTmpRel_B_to_A Is Nothing Then
			oTmpRel_B_to_A = New PlayerRel
			bAdd_B_to_A = True
			With oTmpRel_B_to_A
				.oPlayerRegards = oPlayer2
				.oThisPlayer = oPlayer1
				.WithThisScore = elRelTypes.eNeutral
			End With
            oPlayer2.SetPlayerRel(lPlayerID, oTmpRel_B_to_A, True)

			yMsg_B_to_A = GetAddPlayerRelMessage(oTmpRel_B_to_A)

			oPlayer2.SendPlayerMessage(yMsg_B_to_A, False, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)
		End If

		If bAdd_A_to_B OrElse bAdd_B_to_A Then
			For X = 0 To mlDomainUB
				If bAdd_A_to_B Then
					moDomains(X).SendData(yMsg_A_to_B)
				End If
				If bAdd_B_to_A Then
					moDomains(X).SendData(yMsg_B_to_A)
				End If
			Next X
		End If

	End Sub
	Private Sub HandleAddToFleet(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oFleet As UnitGroup = CType(GetEpicaObject(lFleetID, ObjectType.eUnitGroup), UnitGroup)

		If oFleet Is Nothing = False Then

			If oFleet.CanAlterComposition = True Then
				If oFleet.oOwner.ObjectID = mlClientPlayer(lIndex) OrElse (oFleet.oOwner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eModifyBattleGroups) <> 0) Then
					If oFleet.oOwner.oSpecials.yMaxBattleGroupUnits > oFleet.GetElementCount Then
						For X As Int32 = 0 To lCnt - 1
							Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

							For lUnitIdx As Int32 = 0 To glUnitUB
								If glUnitIdx(lUnitIdx) = lID Then
                                    oFleet.AddUnit(lUnitIdx, True)
									Exit For
								End If
							Next lUnitIdx
						Next X
					End If
				Else
					LogEvent(LogEventType.PossibleCheat, "HandleAddToFleet Owner Mismatch: " & mlClientPlayer(lIndex))
				End If
			Else
				moClients(lIndex).SendData(GetAddObjectMessage(oFleet, GlobalMessageCode.eAddObjectCommand))
			End If
		End If

    End Sub
    Private Sub HandleAddTransport(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > -1 Then
            LogEvent(LogEventType.PossibleCheat, "Attempting to AddTransport as Alias: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim oUnit As Unit = GetEpicaUnit(lUnitID)
        If oUnit Is Nothing Then Return
        If oUnit.Owner Is Nothing Then Return
        If oUnit.Owner.ObjectID <> mlClientPlayer(lIndex) Then Return

        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing Then Return

        Dim yResp() As Byte
        If oPlayer.TransportCount() > 99 Then
            ReDim yResp(6)
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eAddTransport).CopyTo(yResp, lPos) : lPos += 2
            yResp(lPos) = 1 : lPos += 1
            System.BitConverter.GetBytes(lUnitID).CopyTo(yResp, lPos) : lPos += 4
            moClients(lIndex).SendData(yResp)
            Return
        Else
            Dim oTrans As Transport = Transport.CreateFromUnit(oUnit)
            If oTrans Is Nothing Then
                ReDim yResp(6)
                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eAddTransport).CopyTo(yResp, lPos) : lPos += 2
                yResp(lPos) = 2 : lPos += 1
                System.BitConverter.GetBytes(lUnitID).CopyTo(yResp, lPos) : lPos += 4
                moClients(lIndex).SendData(yResp)
                Return
            End If

            'Ok, remove the unit
            Dim lIdx As Int32 = -1
            Dim lID As Int32 = oUnit.ObjectID
            For X As Int32 = 0 To glUnitUB
                If glUnitIdx(X) = lID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If oUnit.ParentObject Is Nothing = False Then
                If CType(oUnit.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    With CType(oUnit.ParentObject, Facility)
                        Try
                            For X As Int32 = 0 To .lHangarUB
                                If .lHangarIdx(X) = oUnit.ObjectID Then
                                    If .oHangarContents(X) Is Nothing = False AndAlso .oHangarContents(X).ObjTypeID = ObjectType.eUnit AndAlso .oHangarContents(X).ObjectID = oUnit.ObjectID Then
                                        .lHangarIdx(X) = -1
                                        .oHangarContents(X) = Nothing
                                    End If
                                End If
                            Next X
                        Catch
                        End Try
                    End With
                End If
            End If

            oUnit.DeleteEntity(lIdx)
            If lIdx > -1 Then
                glUnitIdx(lIdx) = -1
                goUnit(lIdx) = Nothing
            End If

            oPlayer.AddTransport(oTrans)

            ReDim yResp(18)
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eAddTransport).CopyTo(yResp, lPos) : lPos += 2
            yResp(lPos) = 0 : lPos += 1
            System.BitConverter.GetBytes(lUnitID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oTrans.TransportID).CopyTo(yResp, lPos) : lPos += 4
            yResp(lPos) = oTrans.TransFlags : lPos += 1
            System.BitConverter.GetBytes(oTrans.LocationID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oTrans.LocationTypeID).CopyTo(yResp, lPos) : lPos += 2

            moClients(lIndex).SendData(yResp)
            End If
    End Sub
    Private Sub HandleAdvancedEventConfig(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2    'for msgcode

        Dim yType As Byte = yData(lPos) : lPos += 1
        Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEvent As GuildEvent = oPlayer.oGuild.GetEvent(lEventID)
        If oEvent Is Nothing Then Return

        If yType = 1 Then
            'save race config
            If oEvent.lPostedBy <> lPlayerID Then
                LogEvent(LogEventType.PossibleCheat, "Player configure race without permission: " & mlClientPlayer(lIndex))
                Return
            End If
            Dim oRC As New RaceConfig()
            With oRC
                .lEventID = oEvent.EventID
                .blEntryFee = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .lMinHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lMaxHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yMinRacers = yData(lPos) : lPos += 1
                .yFirstPlace = yData(lPos) : lPos += 1
                .ySecondPlace = yData(lPos) : lPos += 1
                .yThirdPlace = yData(lPos) : lPos += 1
                .yGuildTake = yData(lPos) : lPos += 1
                .yLaps = yData(lPos) : lPos += 1
                .yGroundOnly = yData(lPos) : lPos += 1

                .lWPUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
                ReDim .uWP(.lWPUB)
                For X As Int32 = 0 To .lWPUB
                    .uWP(X).EnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .uWP(X).EnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .uWP(X).lX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .uWP(X).lZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Next X
                If oRC.VerifyData() = False Then
                    LogEvent(LogEventType.PossibleCheat, "Player race configuration was invalid: " & mlClientPlayer(lIndex))
                    Return
                End If
            End With
            oEvent.oAdvancedConfig = oRC
            oPlayer.oGuild.SendMsgToGuildMembers(oRC.GetMsg)
        ElseIf yType = 2 Then
            'save tournament config
            If oEvent.lPostedBy <> lPlayerID Then
                LogEvent(LogEventType.PossibleCheat, "Player configure tournament without permission: " & mlClientPlayer(lIndex))
                Return
            End If
            Dim oT As New TournamentConfig
            With oT
                .lEventID = lEventID
                .blEntryFee = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .lMapID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yMaxPlayers = yData(lPos) : lPos += 1
                .yMaxUnits = yData(lPos) : lPos += 1
                .yMaxGround = yData(lPos) : lPos += 1
                .yGuildTake = yData(lPos) : lPos += 1
                .yMaxAir = yData(lPos) : lPos += 1
            End With
            oEvent.oAdvancedConfig = oT
            oPlayer.oGuild.SendMsgToGuildMembers(oT.GetMsg)
        Else
            'requesting config
            If oEvent.oAdvancedConfig Is Nothing = False Then
                moClients(lIndex).SendData(oEvent.oAdvancedConfig.GetMsg)
            End If
        End If

    End Sub
    Private Sub HandleApplyRouteToEntities(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 0
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Ok, get the player
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eMoveUnits Or AliasingRights.eViewUnitsAndFacilities)) <> (AliasingRights.eMoveUnits Or AliasingRights.eViewUnitsAndFacilities) Then
                Return
            Else
                lPlayerID = mlAliasedAs(lIndex)
            End If
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim lTemplateID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oTemplate As RouteTemplate = oPlayer.GetRouteTemplate(lTemplateID)
        If oTemplate Is Nothing Then Return

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For X As Int32 = 0 To lCnt - 1
            Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim oUnit As Unit = GetEpicaUnit(lUnitID)
            If oUnit Is Nothing = False Then
                If oUnit.Owner.ObjectID <> lPlayerID Then
                    LogEvent(LogEventType.PossibleCheat, "Player attempting to Apply route to non-owned unit: " & mlClientPlayer(lIndex))
                Else
                    oTemplate.AssignRouteToUnit(oUnit)
                    If iMsgCode = GlobalMessageCode.eBeginRouteForEntities Then
                        oUnit.lCurrentRouteIdx = -1
                        oUnit.bRoutePaused = False
                        oUnit.bRunRouteOnce = False
                        oUnit.ProcessNextRouteItem()
                    End If
                End If
            End If
        Next X
         

    End Sub
	Private Sub HandleBugSubmission(ByVal yData() As Byte, ByVal lSocketIndex As Int32)

		Dim lPos As Int32 = 2 'msgcode
		lPos += 4		'for bug id (no longer used)

		Dim oBug As New EpicaBug()
		With oBug
			lPos += 4		'created by

			.yCategory = yData(lPos) : lPos += 1
			.ySubCat = yData(lPos) : lPos += 1
			.yOccurs = yData(lPos) : lPos += 1
			.yPriority = yData(lPos) : lPos += 1
			.yStatus = yData(lPos) : lPos += 1

			.lAssignedTo = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			'Problem Desc
			Dim lLen As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim yTemp(lLen - 1) As Byte
			Array.Copy(yData, lPos, yTemp, 0, lLen)
			Dim sTemp As String = System.Text.ASCIIEncoding.ASCII.GetString(yTemp)
			.sProblemDesc = sTemp
			lPos += lLen

			lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			ReDim yTemp(lLen - 1)
			Array.Copy(yData, lPos, yTemp, 0, lLen)
			sTemp = System.Text.ASCIIEncoding.ASCII.GetString(yTemp)
			.sStepsToProduce = sTemp
			lPos += lLen

			lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			ReDim yTemp(lLen - 1)
			Array.Copy(yData, lPos, yTemp, 0, lLen)
			sTemp = System.Text.ASCIIEncoding.ASCII.GetString(yTemp)
			.sDevNotes = sTemp
			lPos += lLen

			.DataChanged()
			'save the object now...
			If .SaveObject() = False Then
				LogEvent(LogEventType.CriticalError, "Unable to save Bug Report!")
			End If
			If .NewSaveObject() = False Then
				LogEvent(LogEventType.CriticalError, "Unable to save Bug Report to Website!")
			End If
		End With

		'Dim lBugID As Int32 = System.BitConverter.ToInt32(yData, 2)
		'Dim oBug As EpicaBug = Nothing
		'Dim X As Int32
		'Dim lIdx As Int32 = -1
		'Dim lLen As Int32
		'Dim lPos As Int32
		'Dim yTemp() As Byte

		'Dim bUpdatePlayer As Boolean = True

		'Dim lUpdates As Int32 = 0
		'Dim sTemp As String

		'If lBugID <> -1 Then
		'    'ok, find our bug
		'    For X = 0 To glBugUB
		'        If goBugs(X).lBugID = lBugID Then
		'            lIdx = X
		'            Exit For
		'        End If
		'    Next X
		'End If

		'If lIdx = -1 Then
		'    glBugUB += 1
		'    ReDim Preserve goBugs(glBugUB)
		'    goBugs(glBugUB) = New EpicaBug
		'    goBugs(glBugUB).lBugID = lBugID
		'    lIdx = glBugUB
		'End If

		'With goBugs(lIdx)
		'    If .lBugID < 1 Then .lCreatedBy = System.BitConverter.ToInt32(yData, 6)

		'    bUpdatePlayer = (mlClientPlayer(lSocketIndex) <> .lCreatedBy)

		'    If .yCategory <> yData(10) Then lUpdates = lUpdates Or 1
		'    .yCategory = yData(10)
		'    If .ySubCat <> yData(11) Then lUpdates = lUpdates Or 2
		'    .ySubCat = yData(11)
		'    If .yOccurs <> yData(12) Then lUpdates = lUpdates Or 4
		'    .yOccurs = yData(12)
		'    If .yPriority <> yData(13) Then lUpdates = lUpdates Or 8
		'    .yPriority = yData(13)
		'    If .yStatus <> yData(14) Then lUpdates = lUpdates Or 16
		'    .yStatus = yData(14)
		'    lPos = 15

		'    .lAssignedTo = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		'    'Problem Desc
		'    lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		'    ReDim yTemp(lLen - 1)
		'    Array.Copy(yData, lPos, yTemp, 0, lLen)
		'    sTemp = System.Text.ASCIIEncoding.ASCII.GetString(yTemp)
		'    If .sProblemDesc <> sTemp Then lUpdates = lUpdates Or 32
		'    .sProblemDesc = sTemp
		'    lPos += lLen

		'    lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		'    ReDim yTemp(lLen - 1)
		'    Array.Copy(yData, lPos, yTemp, 0, lLen)
		'    sTemp = System.Text.ASCIIEncoding.ASCII.GetString(yTemp)
		'    If .sStepsToProduce <> sTemp Then lUpdates = lUpdates Or 64
		'    .sStepsToProduce = sTemp
		'    lPos += lLen

		'    lLen = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		'    ReDim yTemp(lLen - 1)
		'    Array.Copy(yData, lPos, yTemp, 0, lLen)
		'    sTemp = System.Text.ASCIIEncoding.ASCII.GetString(yTemp)
		'    If .sDevNotes <> sTemp Then lUpdates = lUpdates Or 128
		'    .sDevNotes = sTemp
		'    lPos += lLen

		'    .DataChanged()
		'    If .SaveObject() = False Then
		'        LogEvent(LogEventType.CriticalError, "Unable to save Bug Report!")
		'    ElseIf bUpdatePlayer = True Then
		'        'Create our message...
		'        Dim oSB As New System.Text.StringBuilder()
		'        oSB.AppendLine("Bug Report " & .lBugID & " has been updated and may require your attention.")
		'        oSB.AppendLine("Bug Description: " & .sProblemDesc)
		'        oSB.AppendLine()
		'        oSB.AppendLine()
		'        oSB.AppendLine("Altered Items:")
		'        If (lUpdates And 1) <> 0 Then oSB.AppendLine("New Category: " & EpicaBug.GetCategoryText(.yCategory))
		'        If (lUpdates And 2) <> 0 Then oSB.AppendLine("New Sub Category: " & EpicaBug.GetSubCategoryText(.yCategory, .ySubCat))
		'        If (lUpdates And 4) <> 0 Then oSB.AppendLine("New Occurrence: " & EpicaBug.GetOccurrenceText(.yOccurs))
		'        If (lUpdates And 8) <> 0 Then oSB.AppendLine("New Priority: " & EpicaBug.GetPriorityText(.yPriority))
		'        If (lUpdates And 16) <> 0 Then oSB.AppendLine("New Status: " & EpicaBug.GetStatusText(.yStatus))
		'        If (lUpdates And 32) <> 0 Then oSB.AppendLine(vbCrLf & "Updated Description: " & .sProblemDesc)
		'        If (lUpdates And 64) <> 0 Then oSB.AppendLine(vbCrLf & "Updated Steps to Produce: " & .sStepsToProduce)
		'        If (lUpdates And 128) <> 0 Then oSB.AppendLine(vbCrLf & "Update Dev Notes: " & .sDevNotes)

		'        Dim oPlayer As Player = GetEpicaPlayer(.lCreatedBy)
		'        If oPlayer Is Nothing = False Then
		'            Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Bug Report Updated", mlClientPlayer(lSocketIndex), GetDateAsNumber(Now), False, BytesToString(oPlayer.PlayerName))
		'            If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
		'        End If
		'        oPlayer = Nothing
		'        If .lAssignedTo > 0 AndAlso .lAssignedTo <> .lCreatedBy Then
		'            oPlayer = GetEpicaPlayer(.lAssignedTo)
		'            If oPlayer Is Nothing = False Then
		'                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Bug Report Updated", mlClientPlayer(lSocketIndex), GetDateAsNumber(Now), False, BytesToString(oPlayer.PlayerName))
		'                If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
		'            End If
		'        End If
		'    End If
		'End With

	End Sub
	Private Sub HandleCaptureKillAgent(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim yActionID As Byte = yData(6)

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) = 0 Then Return
		End If

		Dim oAgent As Agent = GetEpicaAgent(lAgentID)
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oAgent Is Nothing Then Return
		If oPlayer Is Nothing Then Return
		If oAgent.lCapturedBy = lPlayerID Then
			If (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
				Select Case yActionID
					Case 1 'execute
						'ok, kill the agent
						oAgent.KillMe()
						'remove the agent from the capturer's hand
						With oPlayer.oSecurity
							For X As Int32 = 0 To .lCapturedAgentUB
								If .lCapturedAgentIdx(X) = lAgentID Then
									.lCapturedAgentIdx(X) = -1
									.oCapturedAgents(X) = Nothing
								End If
							Next X
						End With
						oPlayer.SendPlayerMessage(yData, False, AliasingRights.eViewAgents)

						If (oAgent.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                            Dim oPC As PlayerComm = oAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                                BytesToString(oAgent.AgentName) & " was publicly executed by " & oPlayer.sPlayerNameProper & ".", "Agent Executed", _
                                oAgent.oOwner.ObjectID, GetDateAsNumber(Now), False, oAgent.oOwner.sPlayerNameProper, Nothing)
							If oPC Is Nothing = False Then
								oAgent.oOwner.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
							End If
						End If
					Case 2 'interrogate
						Dim lNewID As Int32 = System.BitConverter.ToInt32(yData, 7)
						If lNewID <> -1 Then
							Dim oInterrogator As Agent = GetEpicaAgent(lNewID)
							If oInterrogator Is Nothing = False AndAlso oInterrogator.oOwner.ObjectID = lPlayerID Then
								oAgent.lInterrogatorID = lNewID
								oAgent.yInterrogationState = 1
								goAgentEngine.CancelAllAgentEvents(oAgent.ObjectID)
								goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.InterrogateTest, oAgent, Nothing, glCurrentCycle + AgentEngine.ml_MISSION_CHECK_DELAY)

								oInterrogator.oOwner.SendPlayerMessage(oAgent.GetCapturedAgentMsg(), False, AliasingRights.eViewAgents)
							Else
								LogEvent(LogEventType.PossibleCheat, "HandleCaptureKillAgent: Interrogator is nothing or owner mismatch. PlayerID: " & mlClientPlayer(lIndex))
								Return
							End If
						End If
					Case Else
						If yActionID <> 3 Then LogEvent(LogEventType.PossibleCheat, "HandleCaptureKillAgent: Invalid ActionID passed in (" & yActionID & "). PlayerID: " & mlClientPlayer(lIndex))
						'release
						'remove the agent from the capturer's hand
						With oPlayer.oSecurity
							For X As Int32 = 0 To .lCapturedAgentUB
								If .lCapturedAgentIdx(X) = lAgentID Then
									.lCapturedAgentIdx(X) = -1
									.oCapturedAgents(X) = Nothing
								End If
							Next X
						End With
						oPlayer.SendPlayerMessage(yData, False, AliasingRights.eViewAgents)

						'Now, send the agent to home...
						oAgent.yHealth = 100
                        oAgent.lCapturedBy = -1
                        oAgent.lCapturedOn = -1
                        oAgent.lPrisonTestCycles = 0
						If (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then oAgent.lAgentStatus = oAgent.lAgentStatus Xor AgentStatus.HasBeenCaptured
						oAgent.lInterrogatorID = -1
						oAgent.yInterrogationState = 0
						goAgentEngine.CancelAllAgentEvents(oAgent.ObjectID)
						oAgent.ReturnHome()

						If (oAgent.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                            Dim oPC As PlayerComm = oAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, BytesToString(oAgent.AgentName) & " has been released by " & oPlayer.sPlayerNameProper & " and is returning home.", "Agent Released", oPlayer.ObjectID, GetDateAsNumber(Now), False, oAgent.oOwner.sPlayerNameProper, Nothing)
							If oPC Is Nothing = False Then
								oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
							End If
						End If
						oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)


				End Select
			Else
				LogEvent(LogEventType.PossibleCheat, "HandleCaptureKillAgent: Captured agent is not captured! PlayerID: " & mlClientPlayer(lIndex))
			End If
		Else
			LogEvent(LogEventType.PossibleCheat, "HandleCaptureKillAgent: Captured By is not player! PlayerID: " & mlClientPlayer(lIndex))
		End If
	End Sub
	Private Sub HandleChangeEnvironment(ByVal yData() As Byte, ByVal lIndex As Int32)
		'has PlayerID, New EnvirID, New EnvirTypeID
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 6)
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 10)

		If mlClientPlayer(lIndex) <> lPlayerID AndAlso mlAliasedAs(lIndex) <> lPlayerID Then
			LogEvent(LogEventType.PossibleCheat, "HandleChangeEnvironment: Player Mismatch. Player: " & mlClientPlayer(lIndex))
			Return
		End If

		If iEnvirTypeID = ObjectType.ePlanet OrElse iEnvirTypeID = ObjectType.eSolarSystem Then
			Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
			If oPlayer Is Nothing = False Then
				oPlayer.lLastViewedEnvir = lEnvirID
				oPlayer.iLastViewedEnvirType = iEnvirTypeID
				oPlayer.DataChanged()

				'Ok, now, send the player's environment...
                Dim yTemp() As Byte = GetPlayerEnvirMsg(lPlayerID, lEnvirID, iEnvirTypeID)
                If yTemp Is Nothing = False Then moClients(lIndex).SendData(yTemp)
                moOperator.SendData(yData)
            End If

            System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yData, 2)
            SendToEmailSrvr(yData) 'moEmailSrvr.SendData(yData)
		End If
    End Sub
    Private Sub HandleChatChannelCommand(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        'client left room for the playerid
        System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yData, lPos)
        'now, send to the email server
        moEmailSrvr.SendData(yData)
    End Sub
	Private Sub HandleChatMsg(ByVal lIndex As Int32, ByVal yData() As Byte)
		'Ok, 
		Dim X As Int32
        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim sMsg As String = GetStringFromBytes(yData, 6, lLen)
		Dim sPlayerName As String = ""

		If sMsg.ToLower.StartsWith("/addplayer") = True Then
			If mlClientPlayer(lIndex) = 1 OrElse mlClientPlayer(lIndex) = 6 OrElse mlClientPlayer(lIndex) = 2 OrElse mlClientPlayer(lIndex) = 7 Then
				X = AddPlayerFromCommand(sMsg)
				If X = -1 Then
					moClients(lIndex).SendData(CreateChatMsg(-1, "Unable to create player, check server for msg.", ChatMessageType.eSysAdminMessage))
				ElseIf X = -2 Then
					moClients(lIndex).SendData(CreateChatMsg(-1, "Syntax: /addplayer <DisplayName>, <UserName>, <Password>", ChatMessageType.eSysAdminMessage))
					moClients(lIndex).SendData(CreateChatMsg(-1, "  Each parameter cannot exceed 20 characters.", ChatMessageType.eSysAdminMessage))
				ElseIf X = -3 Then
					moClients(lIndex).SendData(CreateChatMsg(-1, "That Username is already in use.", ChatMessageType.eSysAdminMessage))
				Else : moClients(lIndex).SendData(CreateChatMsg(-1, "Player created as ID: " & X, ChatMessageType.eSysAdminMessage))
				End If
			Else
				moClients(lIndex).SendData(CreateChatMsg(-1, "Invalid Command, command unrecognized.", ChatMessageType.eSysAdminMessage))
			End If
			Return
		ElseIf sMsg.ToLower.StartsWith("/addagent") = True Then
			If mlClientPlayer(lIndex) = 1 Then
				Dim oAgent As Agent = Agent.GenerateNewAgent(GetEpicaPlayer(mlClientPlayer(lIndex)))
				moClients(lIndex).SendData(GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand))
				Return
			Else
				moClients(lIndex).SendData(CreateChatMsg(-1, "Invalid Command, command unrecognized.", ChatMessageType.eSysAdminMessage))
            End If
            'ElseIf sMsg.ToLower.StartsWith("/toggletest") = True Then
            '    gb_IS_TEST_SERVER = Not gb_IS_TEST_SERVER
            '    Dim sOnOff As String = ""
            '    If gb_IS_TEST_SERVER = True Then sOnOff = "ON" Else sOnOff = "OFF"
            '    sPlayerName = GetEpicaPlayer(mlClientPlayer(lIndex)).sPlayerNameProper
            '    Dim yAllMsg() As Byte = CreateChatMsg(-1, sPlayerName & " has set Test Server Mode to " & sOnOff, ChatMessageType.eSysAdminMessage)
            '    For X = 0 To glPlayerUB
            '        If glPlayerIdx(X) > -1 Then
            '            goPlayer(X).SendPlayerMessage(yAllMsg, False, 0)
            '        End If
            '    Next X
            '    Return
            'ElseIf sMsg.ToLower.StartsWith("/givememoney") = True Then
            '    Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
            '    If oPlayer Is Nothing Then Return
            '    oPlayer.blCredits += 2000000000
            '    Return
        ElseIf sMsg.ToLower.StartsWith("/settutorial") = True Then
            If mlClientPlayer(lIndex) = 1 Then
                sPlayerName = sMsg.Substring(12).Trim.ToUpper
                Dim lTmpIdx As Int32 = sPlayerName.IndexOf(" "c)
                If lTmpIdx > -1 Then
                    Dim lStepID As Int32 = CInt(Val(sPlayerName.Substring(lTmpIdx).Trim))
                    sPlayerName = sPlayerName.Substring(0, lTmpIdx).Trim.ToUpper

                    For X = 0 To glPlayerUB
                        If glPlayerIdx(X) <> -1 Then
                            If goPlayer(X).sPlayerName = sPlayerName Then
                                LogEvent(LogEventType.Informational, "Setting " & sPlayerName & " to tutorial step: " & lStepID)
                                goPlayer(X).lTutorialStep = lStepID

                                moClients(lIndex).SendData(CreateChatMsg(-1, goPlayer(X).sPlayerNameProper & " set to step " & lStepID & ".", ChatMessageType.eSysAdminMessage))

                                Return
                            End If
                        End If
                    Next X

                    moClients(lIndex).SendData(CreateChatMsg(-1, "Could not find player.", ChatMessageType.eSysAdminMessage))
                    Return
                End If
            End If
        ElseIf sMsg.ToLower.StartsWith("/disconnect") = True Then
            sPlayerName = sMsg.Substring(11).Trim.ToUpper
            Dim oCmdPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
            If oCmdPlayer Is Nothing = True Then Return
            LogEvent(LogEventType.Informational, oCmdPlayer.sPlayerNameProper & " is disconnecting " & sPlayerName)
            If oCmdPlayer.ObjectID <> 1 AndAlso oCmdPlayer.ObjectID <> 6 AndAlso oCmdPlayer.ObjectID <> 7 Then Return
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    If goPlayer(X).sPlayerName = sPlayerName Then
                        'disconnect the player from the game
                        For Y As Int32 = 0 To mlClientUB
                            If mlClientPlayer(Y) = goPlayer(X).ObjectID Then
                                Try
                                    Dim yTmp(1) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eServerShutdown).CopyTo(yTmp, 0)
                                    moClients(Y).SendData(yTmp)

                                    moClients(Y).Disconnect()
                                    If (mlClientStatusFlags(Y) And ClientStatusFlags.ePlayerLoggedIn) <> 0 Then goPlayer(X).TotalPlayTime += CInt(DateDiff(DateInterval.Second, goPlayer(X).LastLogin, Now))
                                    moClients(Y) = Nothing
                                    goPlayer(X).oSocket = Nothing
                                    goPlayer(X).lConnectedPrimaryID = -1
                                    mlClientPlayer(Y) = -1
                                Catch
                                End Try
                                Exit For
                            End If
                        Next Y
                        Exit For
                    End If
                End If
            Next X
            Return
        ElseIf sMsg.ToLower.StartsWith("/addrel") = True Then
            If mlClientPlayer(lIndex) <> 1 Then Return
            sPlayerName = sMsg.Substring(8).Trim().ToUpper
            Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
            If oPlayer Is Nothing = True Then Return
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    If goPlayer(X).sPlayerName = sPlayerName Then
                        'ok, the player doing this wants to add a rel
                        Dim oRel As PlayerRel = goPlayer(X).GetPlayerRel(mlClientPlayer(lIndex))
                        If oRel Is Nothing = True Then
                            oRel = New PlayerRel()
                            oRel.lCycleOfNextScoreUpdate = -1
                            oRel.oPlayerRegards = goPlayer(X)
                            oRel.oThisPlayer = oPlayer
                            oRel.TargetScore = elRelTypes.eNeutral
                            oRel.WithThisScore = oRel.TargetScore
                            goPlayer(X).SetPlayerRel(mlClientPlayer(lIndex), oRel, True)
                        End If
                        oRel = Nothing

                        oRel = oPlayer.GetPlayerRel(goPlayer(X).ObjectID)
                        If oRel Is Nothing = True Then
                            oRel = New PlayerRel()
                            oRel.lCycleOfNextScoreUpdate = -1
                            oRel.oPlayerRegards = oPlayer
                            oRel.oThisPlayer = goPlayer(X)
                            oRel.TargetScore = elRelTypes.eNeutral
                            oRel.WithThisScore = oRel.TargetScore
                            oPlayer.SetPlayerRel(goPlayer(X).ObjectID, oRel, True)
                        End If
                        oRel = Nothing

                        Exit For
                    End If
                End If
            Next X
            Return

        ElseIf sMsg.ToLower.StartsWith("/played") = True Then
            If mlClientPlayer(lIndex) = 1 OrElse mlClientPlayer(lIndex) = 6 OrElse mlClientPlayer(lIndex) = 2 OrElse mlClientPlayer(lIndex) = 7 Then
                sPlayerName = sMsg.Substring(8).Trim().ToUpper
                For X = 0 To glPlayerUB
                    If glPlayerIdx(X) <> -1 Then
                        If BytesToString(goPlayer(X).PlayerName).Trim.ToUpper = sPlayerName Then
                            sMsg = "Total time played outside of tutorial for " & sPlayerName & " is "
                            Dim lSeconds As Int32 = goPlayer(X).TotalPlayTime

                            If goPlayer(X).lConnectedPrimaryID > -1 Then
                                lSeconds += CInt(Now.Subtract(goPlayer(X).LastLogin).TotalSeconds)
                            End If
                            'If goPlayer(X).PlayedTimeInTutorialOne <> Int32.MinValue Then lSeconds += goPlayer(X).PlayedTimeInTutorialOne
                            If goPlayer(X).yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then lSeconds = 0

                            Dim lTotalSecs As Int32 = lSeconds

                            Dim lMinutes As Int32 = lSeconds \ 60
                            lSeconds -= (lMinutes * 60)
                            Dim lHours As Int32 = lMinutes \ 60
                            lMinutes -= (lHours * 60)
                            Dim lDays As Int32 = lHours \ 24
                            lHours -= (lDays * 24)
                            If lDays <> 0 Then
                                sMsg &= lDays & " days, " & lHours & " hours, " & lMinutes & " minutes and " & lSeconds & " seconds"
                            ElseIf lHours <> 0 Then
                                sMsg &= lHours & " hours, " & lMinutes & " minutes and " & lSeconds & " seconds"
                            Else
                                sMsg &= lMinutes & " minutes and " & lSeconds & " seconds"
                            End If
                            sMsg &= " (" & lTotalSecs.ToString("#,##0") & ")"
                            moClients(lIndex).SendData(CreateChatMsg(-1, sMsg, ChatMessageType.eSysAdminMessage))
                            Return
                        End If
                    End If
                Next X
                moClients(lIndex).SendData(CreateChatMsg(-1, "Could not find player.", ChatMessageType.eSysAdminMessage))
                Return
            Else
                Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
                If oPlayer Is Nothing = False Then
                    sPlayerName = oPlayer.sPlayerNameProper
                    sMsg = "Total time played outside of tutorial for " & sPlayerName & " is "
                    Dim lSeconds As Int32 = oPlayer.TotalPlayTime

                    If oPlayer.lConnectedPrimaryID > -1 Then
                        lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
                    End If
                    'If oPlayer.PlayedTimeInTutorialOne <> Int32.MinValue Then lSeconds += oPlayer.PlayedTimeInTutorialOne

                    If oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then lSeconds = 0

                    Dim lTotalSecs As Int32 = lSeconds

                    Dim lMinutes As Int32 = lSeconds \ 60
                    lSeconds -= (lMinutes * 60)
                    Dim lHours As Int32 = lMinutes \ 60
                    lMinutes -= (lHours * 60)
                    Dim lDays As Int32 = lHours \ 24
                    lHours -= (lDays * 24)
                    If lDays <> 0 Then
                        sMsg &= lDays & " days, " & lHours & " hours, " & lMinutes & " minutes and " & lSeconds & " seconds"
                    ElseIf lHours <> 0 Then
                        sMsg &= lHours & " hours, " & lMinutes & " minutes and " & lSeconds & " seconds"
                    Else
                        sMsg &= lMinutes & " minutes and " & lSeconds & " seconds"
                    End If
                    sMsg &= " (" & lTotalSecs.ToString("#,##0") & ")"
                    moClients(lIndex).SendData(CreateChatMsg(-1, sMsg, ChatMessageType.eSysAdminMessage))
                    Return
                End If

            End If

		ElseIf sMsg.ToLower.StartsWith("/testpirates") = True Then
			If mlClientPlayer(lIndex) = 1 OrElse mlClientPlayer(lIndex) = 2 OrElse mlClientPlayer(lIndex) = 6 Then
				Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
				oPlayer.PirateStartLocX = 0

				Dim yPSLMsg(17) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eGetPirateStartLoc).CopyTo(yPSLMsg, 0)
				System.BitConverter.GetBytes(oPlayer.lStartedEnvirID).CopyTo(yPSLMsg, 2)
				System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yPSLMsg, 6)
				System.BitConverter.GetBytes(oPlayer.lStartLocX).CopyTo(yPSLMsg, 10)
				System.BitConverter.GetBytes(oPlayer.lStartLocZ).CopyTo(yPSLMsg, 14)

				Dim oPlanet As Planet = GetEpicaPlanet(oPlayer.lStartedEnvirID)
				oPlanet.oDomain.DomainSocket.SendData(yPSLMsg)
				Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/who") = True Then
            Dim lSenderID As Int32 = mlClientPlayer(lIndex)
            If lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse mlClientPlayer(lIndex) = 7 OrElse mlClientPlayer(lIndex) = 21296 OrElse mlClientPlayer(lIndex) = 28005 Then 'lSenderID = 221 OrElse lSenderID = 131 OrElse lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse lSenderID = 2067 OrElse lSenderID = 3253 OrElse lSenderID = 7 OrElse lSenderID = 653 OrElse lSenderID = 2076 OrElse lSenderID = 1780 Then
                moClients(lIndex).SendData(CreateChatMsg(-1, "Players Currently Online:", ChatMessageType.eSysAdminMessage))
                Dim lCnt As Int32 = 0
                For X = 0 To glPlayerUB
                    If glPlayerIdx(X) > -1 AndAlso (goPlayer(X).oSocket Is Nothing = False OrElse goPlayer(X).lConnectedPrimaryID > -1) Then
                        moClients(lIndex).SendData(CreateChatMsg(goPlayer(X).ObjectID, "  " & goPlayer(X).sPlayerNameProper, ChatMessageType.eSysAdminMessage))
                        lCnt += 1
                    End If
                Next X
                moClients(lIndex).SendData(CreateChatMsg(-1, "Total Online: " & lCnt, ChatMessageType.eSysAdminMessage))
                Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/csrloggedin") = True Then
            Dim lSenderID As Int32 = mlClientPlayer(lIndex)
            If lSenderID = 221 OrElse lSenderID = 131 OrElse lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse lSenderID = 2067 OrElse lSenderID = 3253 OrElse lSenderID = 7 OrElse lSenderID = 80 OrElse lSenderID = 653 OrElse lSenderID = 2076 OrElse lSenderID = 1780 Then
                sMsg = "/csrloggedin " & lSenderID.ToString
                moOperator.SendData(CreateChatMsg(lSenderID, sMsg, ChatMessageType.eSysAdminMessage))
                Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/csrloggedout") = True Then
            Dim lSenderID As Int32 = mlClientPlayer(lIndex)
            If lSenderID = 221 OrElse lSenderID = 131 OrElse lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse lSenderID = 2067 OrElse lSenderID = 3253 OrElse lSenderID = 7 OrElse lSenderID = 80 OrElse lSenderID = 653 OrElse lSenderID = 2076 OrElse lSenderID = 1780 Then
                sMsg = "/csrloggedout " & lSenderID.ToString
                moOperator.SendData(CreateChatMsg(lSenderID, sMsg, ChatMessageType.eSysAdminMessage))
                Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/aurelium") = True Then
            Dim lSenderID As Int32 = mlClientPlayer(lIndex)
            If lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse mlClientPlayer(lIndex) = 7 Then 'lSenderID = 221 OrElse lSenderID = 131 OrElse lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse lSenderID = 2067 OrElse lSenderID = 3253 OrElse lSenderID = 7 OrElse lSenderID = 80 OrElse lSenderID = 653 OrElse lSenderID = 2076 OrElse lSenderID = 1780 Then
                moClients(lIndex).SendData(CreateChatMsg(-1, "Players Currently Online in Aurelium:", ChatMessageType.eSysAdminMessage))
                Dim lCnt As Int32 = 0
                For X = 0 To glPlayerUB
                    If glPlayerIdx(X) > -1 AndAlso (goPlayer(X).oSocket Is Nothing = False OrElse goPlayer(X).lConnectedPrimaryID > -1) Then
                        If goPlayer(X).yPlayerPhase = eyPlayerPhase.eSecondPhase Then
                            Dim sStatus As String = ""
                            If goPlayer(X).AccountStatus = AccountStatusType.eMondelisActive Then
                                sStatus = "Mondelis"
                            Else
                                sStatus = "Paying: " & (goPlayer(X).AccountStatus = AccountStatusType.eActiveAccount).ToString
                            End If
                            moClients(lIndex).SendData(CreateChatMsg(goPlayer(X).ObjectID, "  " & goPlayer(X).sPlayerNameProper & " (" & sStatus & ")", ChatMessageType.eSysAdminMessage))
                            lCnt += 1
                        End If
                    End If
                Next X
                moClients(lIndex).SendData(CreateChatMsg(-1, "Total Online Aurelium Players: " & lCnt, ChatMessageType.eSysAdminMessage))
                moClients(lIndex).SendData(CreateChatMsg(-1, "-------", ChatMessageType.eSysAdminMessage))
                moClients(lIndex).SendData(CreateChatMsg(-1, "Players Currently Online in Tutorial 1:", ChatMessageType.eSysAdminMessage))
                lCnt = 0
                For X = 0 To glPlayerUB
                    If glPlayerIdx(X) > -1 AndAlso (goPlayer(X).oSocket Is Nothing = False OrElse goPlayer(X).lConnectedPrimaryID > -1) Then
                        If goPlayer(X).yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                            moClients(lIndex).SendData(CreateChatMsg(goPlayer(X).ObjectID, "  " & goPlayer(X).sPlayerNameProper & " (on step: " & goPlayer(X).lTutorialStep & ")", ChatMessageType.eSysAdminMessage))
                            lCnt += 1
                        End If
                    End If
                Next X
                moClients(lIndex).SendData(CreateChatMsg(-1, "Total Online Tutorial 1: " & lCnt, ChatMessageType.eSysAdminMessage))

                Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/mondelis") = True Then
            Dim lSenderID As Int32 = mlClientPlayer(lIndex)
            If lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse mlClientPlayer(lIndex) = 7 Then 'lSenderID = 221 OrElse lSenderID = 131 OrElse lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse lSenderID = 2067 OrElse lSenderID = 3253 OrElse lSenderID = 7 OrElse lSenderID = 80 OrElse lSenderID = 653 OrElse lSenderID = 2076 OrElse lSenderID = 1780 Then
                moClients(lIndex).SendData(CreateChatMsg(-1, "Mondelis Players:", ChatMessageType.eSysAdminMessage))
                Dim lCnt As Int32 = 0
                For X = 0 To glPlayerUB
                    If glPlayerIdx(X) > -1 AndAlso (goPlayer(X).oSocket Is Nothing = False OrElse goPlayer(X).lConnectedPrimaryID > -1) Then
                        If goPlayer(X).AccountStatus = AccountStatusType.eMondelisActive Then
                            moClients(lIndex).SendData(CreateChatMsg(goPlayer(X).ObjectID, "  " & goPlayer(X).sPlayerNameProper, ChatMessageType.eSysAdminMessage))
                            lCnt += 1
                        End If
                    End If
                Next X
                moClients(lIndex).SendData(CreateChatMsg(-1, "Total Online Mondelis Players: " & lCnt, ChatMessageType.eSysAdminMessage))

                Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/instabuildentity") = True Then
            Dim lSenderID As Int32 = mlClientPlayer(lIndex)
            If lSenderID = 1 Then
                'instabuildentity <entitydefid>, <entitydeftypeid>, <builtByentityID>, <builtByentitytypeid>
                sMsg = sMsg.Substring(17).Trim
                Dim sParms() As String = Split(sMsg, ",")
                If sParms Is Nothing = False AndAlso sParms.GetUpperBound(0) = 3 Then
                    Dim lDefID As Int32 = CInt(Val(sParms(0)))
                    Dim iDefTypeID As Int16 = CShort(Val(sParms(1)))
                    Dim lBuiltByID As Int32 = CInt(Val(sParms(2)))
                    Dim iBuiltByTypeID As Int16 = CShort(Val(sParms(3)))

                    'Ok... get the built by
                    Dim oCreator As Epica_Entity = GetEpicaEntity(lBuiltByID, iBuiltByTypeID)
                    If oCreator Is Nothing = False Then

                        If iDefTypeID = ObjectType.eUnitDef Then
                            Dim lUnitIndex As Int32 = AddUnit(oCreator, GetEpicaUnitDef(lDefID), 0)
                            If lUnitIndex < 0 Then Return
                            Dim oUnit As Unit = goUnit(lUnitIndex)

                            With oCreator
                                For lHDIdx As Int32 = 0 To .lHangarUB
                                    If .lHangarIdx(lHDIdx) = oUnit.ObjectID AndAlso .oHangarContents(lHDIdx).ObjTypeID = ObjectType.eUnit Then
                                        'Now, remove the item from the Hangar...
                                        .lHangarIdx(lHDIdx) = -1
                                        .oHangarContents(lHDIdx) = Nothing
                                    End If
                                Next lHDIdx

                                'add our hangar cap back
                                '.Hangar_Cap += oUnit.EntityDef.HullSize

                                'set our parent object

                                'Ensure that our parentobject is the creator... the region server requires this
                                oUnit.LocX = .LocX
                                oUnit.LocZ = .LocZ
                                oUnit.DataChanged()

                                'Send off our data... add... to our parent's parent...
                                Dim iTemp As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID
                                If iTemp = ObjectType.ePlanet Then
                                    'CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eRequestUndock))
                                    CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(oUnit, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, oCreator.LocX, oCreator.LocZ))
                                ElseIf iTemp = ObjectType.eSolarSystem Then
                                    'CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eRequestUndock))
                                    CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(oUnit, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, oCreator.LocX, oCreator.LocZ))
                                End If

                                'Now, set the ParentObject to the Creator's ParentObject because by us sending that msg
                                '  the region server will make it so
                                oUnit.ParentObject = .ParentObject
                                'One more call to datachanged
                                oUnit.DataChanged()

                                'Tell it where to go
                                .SendUndockFirstWaypoint(CType(oUnit, Epica_Entity))
                            End With

                            moClients(lIndex).SendData(CreateChatMsg(-1, "Unit ID " & oUnit.ObjectID & " spawned.", ChatMessageType.eSysAdminMessage))
                        ElseIf iDefTypeID = ObjectType.eFacilityDef Then
                            moClients(lIndex).SendData(CreateChatMsg(-1, "Not implemented yet!", ChatMessageType.eSysAdminMessage))
                        End If

                    Else
                        moClients(lIndex).SendData(CreateChatMsg(-1, "Cannot find the built by.", ChatMessageType.eSysAdminMessage))
                    End If

                Else
                    moClients(lIndex).SendData(CreateChatMsg(-1, "Syntax: /instabuildentity <entitydefid>, <entitydeftypeid>, <builtByentityID>, <builtByentitytypeid>", ChatMessageType.eSysAdminMessage))
                End If
                Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/setmotd ") = True Then
            Dim lSenderID As Int32 = mlClientPlayer(lIndex)
            If lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 Then  'lSenderID = 221 OrElse lSenderID = 131 OrElse lSenderID = 2 OrElse lSenderID = 6 OrElse lSenderID = 1 OrElse lSenderID = 2067 OrElse lSenderID = 3253 OrElse lSenderID = 7 OrElse lSenderID = 80 OrElse lSenderID = 653 OrElse lSenderID = 2076 OrElse lSenderID = 1780 Then
                gfrmDisplayForm.txtMOTD.Text = sMsg.Substring(9)
                gsMOTD = gfrmDisplayForm.txtMOTD.Text
                moClients(lIndex).SendData(CreateChatMsg(-1, "New MOTD: " & gsMOTD, ChatMessageType.eSysAdminMessage))
                Return
            End If
        ElseIf sMsg.ToLower.StartsWith("/warpoints ") = True Then
            Try
                Dim lIdx As Int32 = sMsg.IndexOf(","c)
                If lIdx < 12 Then Return
                Dim lID As Int32 = CInt(Val(sMsg.Substring(11, lIdx - 11)))
                Dim iTypeID As Int16 = CShort(Val(sMsg.Substring(lIdx + 1)))

                Dim oEntity As Epica_Entity = GetEpicaEntity(lID, iTypeID)
                If oEntity Is Nothing = False Then
                    If oEntity.Owner Is Nothing OrElse oEntity.Owner.ObjectID <> mlClientPlayer(lIndex) Then
                        moClients(lIndex).SendData(CreateChatMsg(-1, "That unit does not belong to you.", ChatMessageType.eSysAdminMessage))
                        Return
                    End If
                    Dim fWarpoints As Single = 0.0F
                    Dim lUpkeep As Int32 = 0
                    Dim lCR As Int32 = 0
                    If iTypeID = ObjectType.eUnit Then
                        lCR = CType(oEntity, Unit).EntityDef.CombatRating
                    ElseIf iTypeID = ObjectType.eFacility Then
                        lCR = CType(oEntity, Facility).EntityDef.CombatRating
                    End If
                    lUpkeep = Math.Max(lCR \ 10000, 1)  'max(int(UnitCR/10000),1)
                    fWarpoints = lCR * 0.01F  'unit CR/100 

                    If mlClientPlayer(lIndex) = 1 OrElse mlClientPlayer(lIndex) = 3510 OrElse mlClientPlayer(lIndex) = 20611 Then
                        moClients(lIndex).SendData(CreateChatMsg(-1, "Base Warpoints: " & fWarpoints.ToString("#,##0.#0") & ". Upkeep: " & lUpkeep.ToString("#,##0") & ". Combat Rating: " & lCR.ToString("#,##0"), ChatMessageType.eSysAdminMessage))
                    Else
                        moClients(lIndex).SendData(CreateChatMsg(-1, "Base Warpoints: " & fWarpoints.ToString("#,##0.#0") & ". Upkeep: " & lUpkeep.ToString("#,##0") & ".", ChatMessageType.eSysAdminMessage))
                    End If

                Else
                    moClients(lIndex).SendData(CreateChatMsg(-1, "Unable to find that unit/facility.", ChatMessageType.eSysAdminMessage))
                End If
            Catch
            End Try
            Return
        End If

        Dim yForward(10 + sMsg.Length) As Byte
        'we need to send the email server...
        'msgcode (2), senderid (4), len (4), msg ()
        System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yForward, 0)
        System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yForward, 2)
        Array.Copy(yData, 2, yForward, 6, yData.Length - 2)

        moEmailSrvr.SendData(yForward)
	End Sub
	Private Sub HandleCheckGuildName(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim sCompareName As String = GetStringFromBytes(yData, 2, 50).Trim.ToUpper
        Dim yValue As Byte = 0

        If sCompareName <> FilterBadWords(sCompareName) Then
            yValue = 1
        End If

        If yValue = 0 Then
            For X As Int32 = 0 To glGuildUB
                If glGuildIdx(X) <> -1 Then
                    Dim oGuild As Guild = goGuild(X)
                    If oGuild Is Nothing = False Then
                        If oGuild.sSearchableName = sCompareName Then
                            yValue = 1
                            Exit For
                        End If
                    End If
                End If
            Next X
        End If
        If yValue = 0 Then
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) > 0 Then
                    Dim oTmpPlayer As Player = goPlayer(X)
                    If oTmpPlayer Is Nothing = False Then
                        If oTmpPlayer.sPlayerName = sCompareName Then
                            yValue = 1
                            Exit For
                        End If
                    End If
                End If
            Next X
        End If

        Dim yMsg(52) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eCheckGuildName).CopyTo(yMsg, 0)
        Array.Copy(yData, 2, yMsg, 2, 50)
        yMsg(52) = yValue
        moClients(lIndex).SendData(yMsg)
    End Sub
    Private Sub HandleClaimItem(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lOffer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = Nothing

        If lID = 724000123 AndAlso iTypeID = 12305 AndAlso lOffer = 661110104 Then
            'Ok, player got the winnebago
            Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
            If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
            oPlayer = GetEpicaPlayer(lPlayerID)

            'ok, see if hte player already has the winnebago
            Dim bFound As Boolean = False
            For X As Int32 = 0 To oPlayer.Claimables.GetUpperBound(0)
                With oPlayer.Claimables(X)
                    If .lID = 147 AndAlso .iTypeID = ObjectType.eHullTech Then
                        If (.yClaimFlag And eyClaimFlags.eClaimed) <> 0 Then Return
                        .yClaimFlag = .yClaimFlag Or eyClaimFlags.eClaimed
                        .SaveObject()
                        bFound = True
                        Exit For
                    End If
                End With
            Next X
            If bFound = False Then
                Dim oClaim As New Claimable()
                With oClaim
                    .lOfferCode = 0
                    .lID = 147
                    .iTypeID = ObjectType.eHullTech
                    .lPlayerID = oPlayer.ObjectID
                    .yItemName = StringToBytes("DSE Company Car")
                    .yClaimFlag = eyClaimFlags.eClaimed
                    .SaveObject()
                End With
                ReDim Preserve oPlayer.Claimables(oPlayer.Claimables.GetUpperBound(0) + 1)
                oPlayer.Claimables(oPlayer.Claimables.GetUpperBound(0)) = oClaim
            End If
        ElseIf lID = -1 OrElse iTypeID = -1 OrElse lOffer = -1 Then
            'ok requesting
            Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
            If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then lPlayerID = mlAliasedAs(lIndex)
            oPlayer = GetEpicaPlayer(lPlayerID) 
        Else
            'ok, claiming
            Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
            If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
            oPlayer = GetEpicaPlayer(lPlayerID)

            'First, verify nothing in that offercode is already claimed
            Dim bClaimed As Boolean = False
            For X As Int32 = 0 To oPlayer.Claimables.GetUpperBound(0)
                With oPlayer.Claimables(X)
                    If .lOfferCode = lOffer Then
                        If (.yClaimFlag And eyClaimFlags.eClaimed) <> 0 Then
                            bClaimed = True
                            Exit For
                        End If
                    End If
                End With
            Next X
            If bClaimed = False Then
                For X As Int32 = 0 To oPlayer.Claimables.GetUpperBound(0)
                    With oPlayer.Claimables(X)
                        If .lID = lID AndAlso .iTypeID = iTypeID AndAlso .lOfferCode = lOffer Then
                            'Ok, check if the typeid = specialtech
                            If .iTypeID = ObjectType.eSpecialTech Then
                                'First, check if the blQuantity is 0
                                If .blQuantity > 0 Then
                                    'Ok, let's get the extended ID value which is the lab to apply the bonus to
                                    Dim lExtID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    Dim oFac As Facility = GetEpicaFacility(lExtID)

                                    Dim bFacGood As Boolean = False
                                    If oFac Is Nothing = False AndAlso oFac.Owner Is Nothing = False AndAlso oFac.Owner.ObjectID = oPlayer.ObjectID Then
                                        If oFac.yProductionType = ProductionType.eResearch Then
                                            If oFac.CurrentProduction Is Nothing = False AndAlso oFac.bProducing = True Then
                                                Dim iVal As Int16 = oFac.CurrentProduction.ProductionTypeID
                                                If iVal = ObjectType.eAlloyTech OrElse iVal = ObjectType.eArmorTech OrElse iVal = ObjectType.eEngineTech OrElse _
                                                   iVal = ObjectType.eHullTech OrElse iVal = ObjectType.ePrototype OrElse iVal = ObjectType.eRadarTech OrElse _
                                                   iVal = ObjectType.eShieldTech OrElse iVal = ObjectType.eSpecialTech OrElse iVal = ObjectType.eWeaponTech Then
                                                    bFacGood = True
                                                End If
                                            End If
                                        End If
                                    End If

                                    Dim oTech As Epica_Tech = oPlayer.GetTech(oFac.CurrentProduction.ProductionID, oFac.CurrentProduction.ProductionTypeID)
                                    If oTech Is Nothing Then bFacGood = False

                                    If bFacGood = False Then
                                        oPlayer.SendPlayerMessage(CreateChatMsg(-1, "You must select one of your research facilities that is currently researching a project.", ChatMessageType.eSysAdminMessage), False, 0)
                                        Return
                                    End If

                                    'Now, apply it... if the result is no more points then the item is fully claimed
                                    .blQuantity = oTech.AdjustProductionFromClaim(.blQuantity)
                                    If .blQuantity < 1 Then
                                        .yClaimFlag = .yClaimFlag Or eyClaimFlags.eClaimed
                                    End If
                                Else
                                    .yClaimFlag = .yClaimFlag Or eyClaimFlags.eClaimed
                                End If
                            Else
                                .yClaimFlag = .yClaimFlag Or eyClaimFlags.eClaimed
                            End If

                            .SaveObject()
                            Exit For
                        End If
                    End With
                Next X
            End If

        End If

        If oPlayer Is Nothing Then Return

        Dim yMsg(5 + (oPlayer.Claimables.Length * 31)) As Byte
        lPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eClaimItem).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(oPlayer.Claimables.Length).CopyTo(yMsg, lPos) : lPos += 4
        For X As Int32 = 0 To oPlayer.Claimables.GetUpperBound(0)
            With oPlayer.Claimables(X)
                System.BitConverter.GetBytes(.lID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.iTypeID).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(.lOfferCode).CopyTo(yMsg, lPos) : lPos += 4
                .yItemName.CopyTo(yMsg, lPos) : lPos += 20
                yMsg(lPos) = .yClaimFlag : lPos += 1
            End With
        Next X
        oPlayer.SendPlayerMessage(yMsg, False, 0)

    End Sub
    Private Sub HandleClearBudgetAlert(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim lPos As Int32 = 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then oPlayer.oBudget.SetEnvironmentInConflict(lEnvirID, iEnvirTypeID, False)
    End Sub
    Private Sub HandleClearUnitProdQueue(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2

        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2


        If iEntityTypeID <> ObjectType.eUnit Then Return
        Dim oUnit As Unit = GetEpicaUnit(lEntityID)
        If oUnit Is Nothing Then Return

        If oUnit.Owner.ObjectID <> mlClientPlayer(lIndex) AndAlso (mlAliasedAs(lIndex) <> oUnit.Owner.ObjectID OrElse (mlAliasedRights(lIndex) And (AliasingRights.eCancelProduction Or AliasingRights.eMoveUnits)) = 0) Then
            LogEvent(LogEventType.PossibleCheat, "Player Attempting to clear prod queue without permission: " & mlClientPlayer(lIndex))
            Return
        End If

        oUnit.ClearProdQueue()
    End Sub
    Private Sub HandleClientDismantleObject(ByVal yData() As Byte, ByVal lIndex As Int32)

        Dim lPos As Int32 = 2
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iParentTypeID <> ObjectType.eFacility AndAlso iParentTypeID <> ObjectType.eUnit Then
            LogEvent(LogEventType.PossibleCheat, "ClientDismantleObject, parenttype is not an entity: " & mlClientPlayer(lIndex))
            Return
        End If

        If iObjTypeID <> ObjectType.eFacility AndAlso iObjTypeID <> ObjectType.eUnit Then
            LogEvent(LogEventType.PossibleCheat, "ClientDismantleObject, objecttype is not an entity: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim oParent As Epica_Entity = GetEpicaEntity(lParentID, iParentTypeID)
        If oParent Is Nothing Then Return

        Dim oEntity As Epica_Entity = GetEpicaEntity(lObjID, iObjTypeID)
        If oEntity Is Nothing Then Return

        If oEntity.ParentObject Is Nothing Then Return
        If oParent.ParentObject Is Nothing Then Return

        With CType(oEntity.ParentObject, Epica_GUID)
            If .ObjectID <> oParent.ObjectID OrElse .ObjTypeID <> oParent.ObjTypeID Then
                LogEvent(LogEventType.PossibleCheat, "ClientDismantleObject, entity not in parent: " & mlClientPlayer(lIndex))
                Return
            End If
        End With
        If oEntity.Owner.ObjectID <> oParent.Owner.ObjectID Then
            LogEvent(LogEventType.PossibleCheat, "ClientDismantleObject, entity owner mismatch: " & mlClientPlayer(lIndex))
            Return
        End If

        If (oEntity.Owner.ObjectID <> mlClientPlayer(lIndex) AndAlso (oEntity.Owner.ObjectID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eDismantle) = 0)) Then
            'Sending the response back to the client indicates a failed command
            moClients(lIndex).SendData(yData)
            Return
        End If

        'Cost: originally in the senate proposal, removed for sake of ease and understanding
        '1) 25 % the cost to build the structure unit  
        '2) 10% credit cost of all componets on the ship reguardless if they are recoved or not.  
        '3) Minerals used in production would be lost.

        If iParentTypeID = ObjectType.eFacility AndAlso iObjTypeID = ObjectType.eFacility Then          'station dismantling station child
            CType(oParent, Facility).DismantleChildFacility(CType(oEntity, Facility))
            Return
        ElseIf iParentTypeID = ObjectType.eFacility AndAlso iObjTypeID = ObjectType.eUnit Then          'facility dismantling child unit
            CType(oParent, Facility).DismantleChildUnit(CType(oEntity, Unit))
        ElseIf iParentTypeID = ObjectType.eUnit AndAlso iObjTypeID = ObjectType.eUnit Then              'unit dismantling child unit
            CType(oParent, Unit).DismantleChildUnit(CType(oEntity, Unit))
        End If
        'ElseIf iParentTypeID = ObjectType.eUnit AndAlso iObjTypeID = ObjectType.eFacility Then          'unit dismantling facility - handled in HandleSetMaintenanceTarget

        DestroyEntity(oEntity, False, 0, False, "ClientDismantleObject")

    End Sub
    'Private Sub HandleClientRequestWPTimers(ByVal yData() As Byte, ByVal lIndex As Int32)
    '    If (mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewTreasury) = 0) Then
    '        Return
    '    End If

    '    Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
    '    If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then lPlayerID = mlAliasedAs(lIndex)

    '    Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
    '    If oPlayer Is Nothing = False Then
    '        Dim yMsg(17) As Byte
    '        Dim lPos As Int32 = 0
    '        System.BitConverter.GetBytes(GlobalMessageCode.eWPTimers).CopyTo(yMsg, lPos) : lPos += 2
    '        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

    '        Dim dtTemp As DateTime = DateTime.SpecifyKind(GetDateFromNumber(oPlayer.lLastWPUpkeepTime), DateTimeKind.Local).ToUniversalTime
    '        Dim lTemp As Int32 = GetDateAsNumber(dtTemp)
    '        System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, lPos) : lPos += 4

    '        dtTemp = DateTime.SpecifyKind(GetDateFromNumber(oPlayer.lLastGuildShareUpkeep), DateTimeKind.Local).ToUniversalTime
    '        lTemp = GetDateAsNumber(dtTemp)
    '        System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, lPos) : lPos += 4

    '        Dim lSeconds As Int32 = oPlayer.TotalPlayTime
    '        If oPlayer.lConnectedPrimaryID > -1 Then
    '            lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
    '        End If
    '        lTemp = lSeconds - oPlayer.lLastWPUpkeep
    '        System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, lPos) : lPos += 4

    '        moClients(lIndex).SendData(yMsg)
    '    End If
    'End Sub
    Private Sub HandleClientRequestGuildSharedAssets(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If lPlayerID <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
            If (mlAliasedRights(lIndex) And AliasingRights.eViewUnitsAndFacilities) = 0 Then Return
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return
        If oPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 6
        Dim lCnt As Int32 = 0
        Dim lOverPacket As Int32 = 0
        Dim yResp(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildShareAssets).CopyTo(yResp, 0)
        For x As Int32 = 0 To glUnitUB
            If glUnitIdx(x) <> -1 Then
                Dim oUnit As Unit = goUnit(x)
                If oUnit Is Nothing Then Continue For
                If oUnit.Owner.ObjectID <> oPlayer.ObjectID Then Continue For
                If (oUnit.CurrentStatus And elUnitStatus.eGuildAsset) <> 0 Then
                    With oUnit
                        If yResp.Length > 30000 Then
                            lOverPacket += 1
                        Else
                            Dim oParent As Epica_GUID
                            oParent = CType(.ParentObject, Epica_GUID)
                            ReDim Preserve yResp(yResp.GetUpperBound(0) + 40)       '48
                            System.BitConverter.GetBytes(.ObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                            System.BitConverter.GetBytes(.ObjectID).CopyTo(yResp, lPos) : lPos += 4

                            Select Case oParent.ObjTypeID
                                Case ObjectType.eSolarSystem, ObjectType.ePlanet
                                    'Environment (Used for treeview)
                                    System.BitConverter.GetBytes(oParent.ObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                                    System.BitConverter.GetBytes(oParent.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                                    'Location
                                    System.BitConverter.GetBytes(.LocX).CopyTo(yResp, lPos) : lPos += 4
                                    System.BitConverter.GetBytes(.LocZ).CopyTo(yResp, lPos) : lPos += 4
                                Case ObjectType.eGalaxy
                                    'Environment (Used for treeview)
                                    System.BitConverter.GetBytes(oParent.ObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                                    System.BitConverter.GetBytes(oParent.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                                    'Location
                                    System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                    System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                Case ObjectType.eFacility, ObjectType.eUnit
                                    Dim oDockedAt As Epica_Entity = CType(oParent, Epica_Entity)
                                    If oDockedAt Is Nothing = False Then
                                        'Grand Parent (Used when double clicking the unit to find the environment)
                                        Dim oGrandParent As Epica_GUID
                                        oGrandParent = CType(oDockedAt.ParentObject, Epica_GUID)
                                        If oGrandParent Is Nothing = False Then
                                            System.BitConverter.GetBytes(oGrandParent.ObjTypeID).CopyTo(yResp, lPos) : lPos += 2
                                            System.BitConverter.GetBytes(oGrandParent.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                                        Else
                                            System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 2
                                            System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                        End If
                                        'Location
                                        System.BitConverter.GetBytes(oDockedAt.LocX).CopyTo(yResp, lPos) : lPos += 4
                                        System.BitConverter.GetBytes(oDockedAt.LocZ).CopyTo(yResp, lPos) : lPos += 4
                                    Else 'I dont know where it's docked! 
                                        'Environment (Used for treeview)
                                        System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 2
                                        System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                        'Location
                                        System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                        System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                    End If
                                Case Else
                                    Continue For
                                    'Environment (Used for treeview)
                                    System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 2
                                    System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                    'Location
                                    System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                                    System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
                            End Select
                            .EntityName.CopyTo(yResp, lPos) : lPos += 20

                            'Dim iWarpointUpkeep As Int64 = 0
                            'Dim lCR As Int32 = 0
                            'lCR = .EntityDef.CombatRating
                            'iWarpointUpkeep = Math.Max(lCR \ 10000, 1)
                            'System.BitConverter.GetBytes(iWarpointUpkeep).CopyTo(yResp, lPos) : lPos += 8
                            lCnt += 1
                        End If
                    End With
                End If
            End If
        Next
        If lOverPacket > 0 Then
            ReDim Preserve yResp(yResp.GetUpperBound(0) + 52)
            System.BitConverter.GetBytes(-2I).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lOverPacket).CopyTo(yResp, lPos) : lPos += 4
            'Environment (Used for treeview)
            System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
            'Location
            System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(-1).CopyTo(yResp, lPos) : lPos += 4
            lCnt += 1
        End If
        If lCnt = 0 Then Exit Sub
        System.BitConverter.GetBytes(lCnt).CopyTo(yResp, 2)

        moClients(lIndex).SendData(yResp)
    End Sub

    Private Structure mUniversalAsset
        Dim iEnvirTypeID As Int16
        Dim lEnvirID As Int32
        Dim lParentStarID As Int32

        Dim iModelDefTypeID As Int32
        Dim lModelSubTypeID As Int32

        Dim sModelName As String
        Dim lUnitCount As Int32
    End Structure
    Private Sub HandleClientRequestUniversalAssets(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If lPlayerID <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
            If (mlAliasedRights(lIndex) And AliasingRights.eViewUnitsAndFacilities) = 0 Then Return
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        'One per hour at max.  If abused we can make once per day.
        If oPlayer.lLastUniversalAssetRequest = Int32.MinValue OrElse glCurrentCycle - oPlayer.lLastUniversalAssetRequest > 108000 Then
            oPlayer.lLastUniversalAssetRequest = glCurrentCycle
        Else : Return
        End If

        Dim lPos As Int32 = 6
        Dim lUB As Int32 = -1
        Dim lOverPacket As Int32 = 0
        Dim yResp(5) As Byte
        Dim oUniversalAsset(-1) As mUniversalAsset

        Dim iEnvirTypeID As Int16
        Dim lEnvirID As Int32
        Dim iModelDefTypeID As Int32
        Dim lModelSubTypeID As Int32
        Dim sModelName As String

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestUniversalAssets).CopyTo(yResp, 0)
        For x As Int32 = 0 To glUnitUB
            If glUnitIdx(x) <> -1 Then
                Dim oUnit As Unit = goUnit(x)
                If oUnit Is Nothing Then Continue For
                If oUnit.Owner.ObjectID <> oPlayer.ObjectID Then Continue For

                Dim oDef As Epica_Entity_Def
                oDef = CType(oUnit, Unit).EntityDef
                If oDef Is Nothing Then Continue For

                Dim oModelDef As ModelDef = goModelDefs.GetModelDef(oDef.ModelID)
                If oModelDef Is Nothing Then Continue For

                iModelDefTypeID = oModelDef.TypeID
                lModelSubTypeID = oModelDef.SubTypeID
                sModelName = BytesToString(oDef.DefName)

                With oUnit
                    Dim oParent As Epica_GUID
                    oParent = CType(.ParentObject, Epica_GUID)
                    Select Case oParent.ObjTypeID
                        Case ObjectType.eSolarSystem, ObjectType.ePlanet
                            iEnvirTypeID = oParent.ObjTypeID
                            lEnvirID = oParent.ObjectID
                        Case ObjectType.eGalaxy
                            iEnvirTypeID = oParent.ObjTypeID
                            lEnvirID = oParent.ObjectID
                        Case ObjectType.eFacility, ObjectType.eUnit
                            Dim oDockedAt As Epica_Entity = CType(oParent, Epica_Entity)
                            If oDockedAt Is Nothing = False Then
                                Dim oGrandParent As Epica_GUID
                                oGrandParent = CType(oDockedAt.ParentObject, Epica_GUID)
                                If oGrandParent Is Nothing = False Then
                                    iEnvirTypeID = oGrandParent.ObjTypeID
                                    lEnvirID = oGrandParent.ObjectID
                                Else
                                    Continue For
                                End If
                            Else 'I dont know where it's docked! 
                                Continue For
                            End If
                        Case Else
                            Continue For
                    End Select

                    Dim lIdx As Int32 = -1
                    For y As Int32 = 0 To lUB
                        If oUniversalAsset(y).iEnvirTypeID = iEnvirTypeID AndAlso oUniversalAsset(y).lEnvirID = lEnvirID AndAlso oUniversalAsset(y).iModelDefTypeID = iModelDefTypeID AndAlso oUniversalAsset(y).lModelSubTypeID = lModelSubTypeID Then
                            lIdx = y
                            Exit For
                        End If
                    Next
                    If lIdx = -1 Then
                        lUB += 1
                        lIdx = lUB
                        ReDim Preserve oUniversalAsset(lIdx)
                        With oUniversalAsset(lIdx)
                            .iEnvirTypeID = iEnvirTypeID
                            .lEnvirID = lEnvirID
                            .iModelDefTypeID = iModelDefTypeID
                            .lModelSubTypeID = lModelSubTypeID
                            .lUnitCount = 1
                            .sModelName = BytesToString(oDef.DefName)
                            .lParentStarID = -1
                            If oParent.ObjTypeID = ObjectType.ePlanet Then
                                Dim oPlanet As Planet = CType(oUnit.ParentObject, Planet)
                                If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then
                                    .lParentStarID = oPlanet.ParentSystem.ObjectID
                                End If
                            End If
                        End With
                    Else : oUniversalAsset(lIdx).lUnitCount += 1
                    End If
                End With
            End If
        Next

        If lUB < 0 Then Exit Sub

        ReDim Preserve yResp(oUniversalAsset.Length * 40 + 5)
        System.BitConverter.GetBytes(lUB).CopyTo(yResp, 2) : lPos = 6
        For x As Int32 = 0 To lUB
            System.BitConverter.GetBytes(oUniversalAsset(x).iEnvirTypeID).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(oUniversalAsset(x).lEnvirID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oUniversalAsset(x).lParentStarID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oUniversalAsset(x).iModelDefTypeID).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(oUniversalAsset(x).lModelSubTypeID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oUniversalAsset(x).lUnitCount).CopyTo(yResp, lPos) : lPos += 4
            System.Text.ASCIIEncoding.ASCII.GetBytes(oUniversalAsset(x).sModelName).CopyTo(yResp, lPos) : lPos += 20
        Next

        moClients(lIndex).SendData(yResp)
    End Sub
    Private Sub HandleClientSetEntityStatus(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 8)

        Dim oEntity As Epica_Entity
        Dim iTemp As Int16

        oEntity = CType(GetEpicaObject(lObjID, iTypeID), Epica_Entity)
        If oEntity Is Nothing Then Return

        If oEntity.Owner.ObjectID <> mlClientPlayer(lIndex) AndAlso (mlAliasedAs(lIndex) <> oEntity.Owner.ObjectID OrElse (mlAliasedRights(lIndex) And AliasingRights.eAlterAutoLaunchPower) = 0) Then
            Return
        End If

        'TODO: Ensure that oEntity is a facility before doing below...

        If lStatus = -elUnitStatus.eFacilityPowered Then
            'ok, turning off the facility
            If CType(oEntity, Facility).SetActive(False) = True Then
                iTemp = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
                If iTemp = ObjectType.ePlanet Then
                    CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
                ElseIf iTemp = ObjectType.eSolarSystem Then
                    CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
                End If
            End If

        ElseIf lStatus = elUnitStatus.eFacilityPowered Then
            If CType(oEntity, Facility).SetActive(True) = True Then
                iTemp = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
                If iTemp = ObjectType.ePlanet Then
                    CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
                ElseIf iTemp = ObjectType.eSolarSystem Then
                    CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
                End If
            End If
        End If

        'Dim yResp(11) As Byte
        'System.BitConverter.GetBytes(EpicaMessageCode.eSetEntityStatus).CopyTo(yResp, 0)
        'oEntity.GetGUIDAsString.CopyTo(yResp, 2)
        'If (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
        '    System.BitConverter.GetBytes(-elUnitStatus.eFacilityPowered).CopyTo(yResp, 8)
        'Else
        '    System.BitConverter.GetBytes(elUnitStatus.eFacilityPowered).CopyTo(yResp, 8)
        'End If
        'moClients(lIndex).SendData(yResp)
    End Sub
	Private Sub HandleClientVersion(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
		Dim lTheirVersion As Int32 = System.BitConverter.ToInt32(yData, 2)

		Dim yMsg(5) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eClientVersion).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(gl_CLIENT_VERSION).CopyTo(yMsg, 2)

		moClients(lSocketIndex).SendData(yMsg)

		If lTheirVersion <> gl_CLIENT_VERSION Then
			moClients(lSocketIndex).Disconnect()
			moClients(lSocketIndex) = Nothing
		End If
	End Sub
	Private Sub HandleCreateFleet(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lIdx As Int32 = -1

		Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eCreateBattleGroups) = 0 Then Return
		End If

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		Dim lPlayersBG As Int32 = 0
		Dim bBGNumUsed(oPlayer.oSpecials.yMaxBattlegroups) As Boolean
		For X As Int32 = 0 To glUnitGroupUB
			If glUnitGroupIdx(X) <> -1 AndAlso goUnitGroup(X).oOwner.ObjectID = oPlayer.ObjectID Then
				lPlayersBG += 1

				Dim lVal As Int32 = CInt(Val(BytesToString(goUnitGroup(X).UnitGroupName)))
				If lVal > 0 AndAlso lVal <= bBGNumUsed.GetUpperBound(0) Then bBGNumUsed(lVal) = True
			End If
		Next X

		If lPlayersBG >= oPlayer.oSpecials.yMaxBattlegroups Then
			'TODO: Indicate to the player why it failed
			Return
		ElseIf oPlayer.oSpecials.yMaxBattleGroupUnits <> 255 AndAlso lCnt > oPlayer.oSpecials.yMaxBattleGroupUnits Then
			'TODO: Indicate to the player why it failed
			Return
		End If

		For X As Int32 = 0 To glUnitGroupUB
			If glUnitGroupIdx(X) = -1 Then
				lIdx = X
				glUnitGroupIdx(X) = 0		'to not permit other handlecreatefleets to overwrite it
				Exit For
			End If
		Next X

		If lIdx = -1 Then
			ReDim Preserve glUnitGroupIdx(glUnitGroupUB + 1)
			ReDim Preserve goUnitGroup(glUnitGroupUB + 1)
			glUnitGroupUB += 1
			lIdx = glUnitGroupUB
		End If

		'Now... 
		goUnitGroup(lIdx) = New UnitGroup()
		With goUnitGroup(lIdx)
			.ObjectID = -1
			.ObjTypeID = ObjectType.eUnitGroup
			.oOwner = oPlayer
			.oParentObject = GetEpicaObject(lParentID, iParentTypeID)
			.UnitUB = -1
			.UnitGroupName = StringToBytes("New Unit Group")
		End With

		If goUnitGroup(lIdx).SaveObject = False Then
			glUnitGroupIdx(lIdx) = -1
			'TODO: Relay the failure back to the client
			Return
		Else : glUnitGroupIdx(lIdx) = goUnitGroup(lIdx).ObjectID
		End If

		Dim lLandUnitCnt As Int32 = 0
		Dim lSpaceUnitCnt As Int32 = 0

		For X As Int32 = 0 To lCnt - 1
			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			'Now... 
			For lUnitIdx As Int32 = 0 To glUnitUB
				If glUnitIdx(lUnitIdx) = lID Then
					Dim oUnit As Unit = goUnit(lUnitIdx)
					If oUnit Is Nothing = False Then
                        goUnitGroup(lIdx).AddUnit(lUnitIdx, True)
                        If (oUnit.EntityDef.yChassisType And (ChassisType.eSpaceBased Or ChassisType.eNavalBased)) <> 0 Then
                            lSpaceUnitCnt += 1
                        Else : lLandUnitCnt += 1
                        End If
					End If

					Exit For
				End If
			Next lUnitIdx
		Next X

		'Now, determine whether this is an army or not
		Dim sName As String = "New Battlegroup"
		Dim lBGNum As Int32 = 0
		For X As Int32 = 1 To bBGNumUsed.GetUpperBound(0)
			If bBGNumUsed(X) = False Then
				lBGNum = X
				Exit For
			End If
		Next X
		If lBGNum <> 0 Then sName = (lBGNum).ToString
		Dim sFinalchar As String = sName.Substring(sName.Length - 1, 1)

		Select Case sFinalchar
			Case "1"
				sName &= "st"
			Case "2"
				sName &= "nd"
			Case "3"
				sName &= "rd"
			Case Else
				sName &= "th"
		End Select
		If lLandUnitCnt > lSpaceUnitCnt Then
			sName &= " Army"
		Else : sName &= " Fleet"
		End If
		goUnitGroup(lIdx).UnitGroupName = StringToBytes(sName)

		moClients(lIndex).SendData(GetAddObjectMessage(goUnitGroup(lIdx), GlobalMessageCode.eAddObjectCommand))
	End Sub
	Private Sub HandleCreateGuild(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer.oFormingGuild Is Nothing Then Return

		Dim lPos As Int32 = 2
		Dim lIcon As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim sName As String = GetStringFromBytes(yData, lPos, 50).Trim : lPos += 50

		Dim sCompareName As String = sName.Trim.ToUpper

		Dim yValue As Byte = 0
		For X As Int32 = 0 To glGuildUB
			If glGuildIdx(X) <> -1 Then
				Dim oGuild As Guild = goGuild(X)
				If oGuild Is Nothing = False Then
					If oGuild.sSearchableName = sCompareName Then
						yValue = 1
						Exit For
					End If
				End If
			End If
		Next X
		If yValue = 0 Then
			For X As Int32 = 0 To glPlayerUB
				If glPlayerIdx(X) > 0 Then
					Dim oTmpPlayer As Player = goPlayer(X)
					If oTmpPlayer Is Nothing = False Then
						If oTmpPlayer.sPlayerName = sCompareName Then
							yValue = 1
							Exit For
						End If
					End If
				End If
			Next X
		End If

		If yValue = 1 Then
			Dim yMsg(52) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eCheckGuildName).CopyTo(yMsg, 0)
			Array.Copy(yData, 2, yMsg, 2, 50)
			yMsg(52) = yValue
			moClients(lIndex).SendData(yData)
		Else
            Dim lCnt As Int32 = 0
            Dim sUnaccepted As String = ""
			With oPlayer.oFormingGuild
				If .oMembers Is Nothing = False Then
					For X As Int32 = 0 To .oMembers.GetUpperBound(0)
						If .oMembers(X) Is Nothing = False Then
							If .oMembers(X).lMemberID = lPlayerID OrElse (.oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                                lCnt += 1
                                If sUnaccepted <> "" Then sUnaccepted &= ", "
                                sUnaccepted &= .oMembers(X).oMember.sPlayerNameProper
							End If
						End If
					Next X
				End If
			End With
			If lCnt < 4 Then
                LogEvent(LogEventType.PossibleCheat, "Player attempting to create guild without 4 signups: " & mlClientPlayer(lIndex))
                moClients(lIndex).SendData(CreateChatMsg(-1, "The following players have not accepted: " & sUnaccepted, ChatMessageType.eAllianceMessage))
			Else
				Dim oGuild As New Guild()
				With oGuild
					.dtFormed = Now.ToUniversalTime
					.dtLastGuildRuleChange = Now.ToUniversalTime
					.InMyDomain = True
					.iRecruitFlags = 0
					.lBaseGuildFlags = oPlayer.oFormingGuild.lBaseGuildFlags
					.lIcon = oPlayer.oFormingGuild.lIcon
					.ObjectID = -1
					.ObjTypeID = ObjectType.eGuild
					.sSearchableName = sCompareName
					.yGuildName = oPlayer.oFormingGuild.yGuildName
					.yGuildTaxBaseDay = oPlayer.oFormingGuild.yGuildTaxBaseDay
					.yGuildTaxBaseMonth = oPlayer.oFormingGuild.yGuildTaxBaseMonth
					.yGuildTaxRateInterval = oPlayer.oFormingGuild.yGuildTaxRateInterval
					.yState = eyGuildState.Formed
					.yVoteWeightType = oPlayer.oFormingGuild.yVoteWeightType
				End With
				Dim lIdx As Int32 = AddGuild(oGuild)
				If lIdx = -1 Then Return

				Dim lFounderID As Int32 = -1
				Dim lBeforeSaveRankID() As Int32 = Nothing
				Dim lAfterSaveRankID() As Int32 = Nothing
				Dim lRankIDUB As Int32 = -1
				If oPlayer.oFormingGuild.oRanks Is Nothing = False Then
					For X As Int32 = 0 To oPlayer.oFormingGuild.oRanks.GetUpperBound(0)
						If oPlayer.oFormingGuild.oRanks(X) Is Nothing = False Then
							lRankIDUB += 1
							ReDim Preserve lBeforeSaveRankID(lRankIDUB)
							ReDim Preserve lAfterSaveRankID(lRankIDUB)
							lBeforeSaveRankID(lRankIDUB) = oPlayer.oFormingGuild.oRanks(X).lRankID
							oPlayer.oFormingGuild.oRanks(X).lRankID = -1
							oPlayer.oFormingGuild.oRanks(X).ParentGuild = oGuild
							If oPlayer.oFormingGuild.oRanks(X).SaveObject(oGuild.ObjectID) = False Then Return
							lAfterSaveRankID(lRankIDUB) = oPlayer.oFormingGuild.oRanks(X).lRankID
							oGuild.AddRank(oPlayer.oFormingGuild.oRanks(X))
						End If
					Next X
				Else
					Dim oRank As New GuildRank
					With oRank
						.lRankID = -1
                        .lRankPermissions = RankPermissions.AcceptApplicant Or RankPermissions.AcceptEvents Or RankPermissions.BuildGuildBase Or _
                          RankPermissions.ChangeMOTD Or RankPermissions.ChangeRankNames Or RankPermissions.ChangeRankPermissions Or _
                          RankPermissions.ChangeRankVotingWeight Or RankPermissions.ChangeRecruitment Or RankPermissions.CreateEvents Or _
                          RankPermissions.CreateRanks Or RankPermissions.DeleteEvents Or RankPermissions.DeleteRanks Or RankPermissions.DemoteMember Or _
                          RankPermissions.InviteMember Or RankPermissions.ModifyGuildRelation Or _
                          RankPermissions.PromoteMember Or RankPermissions.ProposeVotes Or RankPermissions.RejectMember Or _
                          RankPermissions.RemoveMember Or RankPermissions.ViewBankLog Or RankPermissions.ViewContentsHiSec Or _
                          RankPermissions.ViewContentsLowSec Or RankPermissions.ViewEventAttachments Or RankPermissions.ViewEvents Or _
                          RankPermissions.ViewGuildBase Or RankPermissions.ViewVotesHistory Or RankPermissions.ViewVotesInProgress Or _
                          RankPermissions.WithdrawHiSec Or RankPermissions.WithdrawLowSec
						.lVoteStrength = 1
						.ParentGuild = oGuild
						.TaxRateFlat = 0
						.TaxRatePercentage = 0
						.TaxRatePercType = eyGuildTaxPercType.CashFlow
						.yPosition = 0
						.yRankName = StringToBytes("Founder")
						If .SaveObject(oGuild.ObjectID) = False Then Return
					End With
					lFounderID = oRank.lRankID
					oGuild.AddRank(oRank)
				End If

				With oPlayer.oFormingGuild
					If .oMembers Is Nothing = False Then
						For X As Int32 = 0 To .oMembers.GetUpperBound(0)
							If .oMembers(X) Is Nothing = False Then
								Dim bFound As Boolean = False
								For Y As Int32 = 0 To lRankIDUB
									If lBeforeSaveRankID(Y) = .oMembers(X).lCreateRankID Then
										bFound = True
										.oMembers(X).lCreateRankID = lAfterSaveRankID(Y)
										Exit For
									End If
								Next Y
								If bFound = False Then .oMembers(X).lCreateRankID = lFounderID

								Dim oTmpPlayer As Player = .oMembers(X).oMember
								If oTmpPlayer Is Nothing = False Then
									oTmpPlayer.lGuildID = oGuild.ObjectID
									oTmpPlayer.lGuildRankID = .oMembers(X).lCreateRankID
								End If
								oGuild.AddMember(.oMembers(X))
							End If
						Next X
					End If
				End With

				'Dim oMember As New GuildMember
				'With oMember
				'	.lMemberID = lPlayerID
				'	.yMemberState = GuildMemberState.Invited Or GuildMemberState.Approved
				'End With
				'oGuild.AddMember(oMember)
				'oMember = Nothing

				'oPlayer.lGuildID = oGuild.ObjectID
				'oPlayer.lGuildRankID = lFounderID

				'For X As Int32 = 0 To oPlayer.lAcceptedInviteUB
				'	If oPlayer.lAcceptedInvites(X) <> -1 Then
				'		oMember = New GuildMember
				'		With oMember
				'			.lMemberID = oPlayer.lAcceptedInvites(X)
				'			.yMemberState = GuildMemberState.Invited Or GuildMemberState.Approved
				'		End With
				'		oGuild.AddMember(oMember)
				'		Dim oTmpPlayer As Player = GetEpicaPlayer(oPlayer.lAcceptedInvites(X))
				'		If oTmpPlayer Is Nothing = False Then
				'			oTmpPlayer.lGuildID = oGuild.ObjectID
				'			oTmpPlayer.lGuildRankID = lFounderID
				'		End If
				'	End If
				'Next X

				'now, call save again
				oGuild.SaveObject()
				'now, go back through and send the add guild objects
				Dim yMsg() As Byte = GetAddObjectMessage(oGuild, GlobalMessageCode.eAddObjectCommand)
				oGuild.SendMsgToGuildMembers(yMsg)

				'oPlayer.SendPlayerMessage(yMsg, False, 0)
				'For X As Int32 = 0 To oPlayer.lAcceptedInviteUB
				'	If oPlayer.lAcceptedInvites(X) <> -1 Then
				'		Dim oTmpPlayer As Player = GetEpicaPlayer(oPlayer.lAcceptedInvites(X))
				'		If oTmpPlayer Is Nothing = False Then
				'			oTmpPlayer.SendPlayerMessage(yMsg, False, 0)
				'		End If
				'	End If
				'Next X

				'Finally, do all players joined
				If oGuild.oMembers Is Nothing = False Then
					For X As Int32 = 0 To oGuild.oMembers.GetUpperBound(0)
						oGuild.MemberJoined(oGuild.oMembers(X).lMemberID, True)
					Next X
				End If
				'oGuild.MemberJoined(lPlayerID)
				'For X As Int32 = 0 To oPlayer.lAcceptedInviteUB
				'	If oPlayer.lAcceptedInvites(X) <> -1 Then
				'		oGuild.MemberJoined(oPlayer.lAcceptedInvites(X))
				'	End If
				'Next X
			End If
		End If
	End Sub
	Private Sub HandleDeathDeposit(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lAmt As Int32 = System.BitConverter.ToInt32(yData, 2)

		If mlAliasedAs(lIndex) > 0 Then Return

		'Assume failure first...
		System.BitConverter.GetBytes(CInt(-1)).CopyTo(yData, 2)

		If lAmt > 0 Then
			Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
			If oPlayer Is Nothing = False Then

                Dim lMaxValue As Int32 = Int32.MaxValue
                Dim blMaxAllowed As Int64 = oPlayer.blMaxPopulation - oPlayer.blDBPopulation
                If blMaxAllowed < 0 Then blMaxAllowed = 0
                If blMaxAllowed < 20000000 Then lMaxValue = CInt(blMaxAllowed * 100)

                If CLng(CLng(oPlayer.DeathBudgetBalance) + CLng(lAmt)) > lMaxValue Then
                    lAmt = lMaxValue - oPlayer.DeathBudgetBalance
                End If

                If lAmt <= oPlayer.blCredits AndAlso oPlayer.PlayerIsDead = False Then
                    'Ok, it works
                    oPlayer.blCredits -= lAmt
                    oPlayer.DeathBudgetBalance += lAmt
                    oPlayer.DataChanged()

                    LogEvent(LogEventType.ExtensiveLogging, "DeathBudget: " & mlClientPlayer(lIndex) & " adds " & lAmt.ToString)
                End If

                If oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                    oPlayer.DeathBudgetBalance = 100000000
                    oPlayer.DataChanged()
                End If

                System.BitConverter.GetBytes(oPlayer.DeathBudgetBalance).CopyTo(yData, 2)
            Else : Return
            End If
        End If

		moClients(lIndex).SendData(yData)
    End Sub
    Private Sub HandleDecommissionTransport(ByVal yData() As Byte, ByVal lIndex As Int32)
        If mlAliasedAs(lIndex) <> -1 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then Return

        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing Then Return

        Dim lTransID As Int32 = System.BitConverter.ToInt32(yData, 2)
        If oPlayer.RemoveTransport(lTransID) = True Then
            moClients(lIndex).SendData(yData)
        End If
    End Sub
	Private Sub HandleDeinfiltrateAgent(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oAgent As Agent = GetEpicaAgent(lAgentID)
		If oAgent Is Nothing = False Then
			'Now, verify we have rights
			Dim lPlayerID As Int32 = oAgent.oOwner.ObjectID

			If lPlayerID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = lPlayerID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) <> 0) Then
				Dim lResult As Int32 = oAgent.AttemptDeinfiltration()
				Dim yResp(9) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eDeinfiltrateAgent).CopyTo(yResp, 0)
				System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yResp, 2)
				System.BitConverter.GetBytes(lResult).CopyTo(yResp, 6)
				oAgent.oOwner.SendPlayerMessage(yResp, False, AliasingRights.eViewAgents)
			Else
				LogEvent(LogEventType.PossibleCheat, "DeinfiltrateAgent Player Owner Mismatch: " & mlClientPlayer(lIndex))
			End If
		End If

	End Sub
	Private Sub HandleDeleteDesign(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iTechTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

		Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
		If oPlayer Is Nothing = False Then
			Dim oTech As Epica_Tech = oPlayer.GetTech(lTechID, iTechTypeID)
			If oTech Is Nothing = False AndAlso oTech.Owner.ObjectID = oPlayer.ObjectID Then
				'Ok, we got it, now, check if it can be deleted
				If oTech.CanBeDeleted() = True Then
					oPlayer.RemoveTech(lTechID, iTechTypeID)
					If oTech.DeleteMe() = True Then
						moClients(lIndex).SendData(yData)
					End If
				End If
			End If
		End If
	End Sub
	Private Sub HandleDeleteEmailItem(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eAlterEmail) = 0 Then Return
		End If
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

        If oPlayer Is Nothing = False Then
            If oPlayer.InMyDomain = False Then
                SendPassThruMsg(yData, oPlayer.ObjectID, oPlayer.ObjectID, oPlayer.ObjTypeID)
                Return
            End If

            Dim yResp() As Byte = DoDeleteEmailItem(oPlayer, yData, 0)
            If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
        End If

	End Sub
	Private Sub HandleDeleteFleet(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, 2)
		If lFleetID = -1 Then Return

		For X As Int32 = 0 To glUnitGroupUB
			If glUnitGroupIdx(X) = lFleetID Then
                'If mlClientPlayer(lIndex) = goUnitGroup(X).oOwner.ObjectID Then
                If goUnitGroup(X).CanAlterComposition = True Then
                    goUnitGroup(X).DeleteMe()
                    goUnitGroup(X) = Nothing
                    glUnitGroupIdx(X) = -1
                Else : moClients(lIndex).SendData(GetAddObjectMessage(goUnitGroup(X), GlobalMessageCode.eAddObjectCommand))
                End If
                'Else
                '	LogEvent(LogEventType.PossibleCheat, "HandleDeleteFleet Owner Mismatch: " & mlClientPlayer(lIndex))
                'End If
				Exit For
			End If
		Next X
	End Sub
    Private Sub HandleDeleteTradeHistoryItem(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If mlClientPlayer(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        If oPlayer.InMyDomain = False Then
            SendPassThruMsg(yData, oPlayer.ObjectID, oPlayer.ObjectID, oPlayer.ObjTypeID)
            Return
        End If

        DoDeleteTradeHistoryItem(oPlayer, yData, 0)
        'oPlayer.DeleteTradeHistoryItem(lTransDate, lOtherPlayer, yResult, yEventType)
    End Sub
    Private Sub HandleEditRouteTemplate(ByVal yData() As Byte, ByVal lIndex As Int32)
        'Ok, get the player
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eMoveUnits Or AliasingRights.eViewUnitsAndFacilities)) <> (AliasingRights.eMoveUnits Or AliasingRights.eViewUnitsAndFacilities) Then
                Return
            Else
                lPlayerID = mlAliasedAs(lIndex)
            End If
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim lPos As Int32 = 2
        Dim yAction As Byte = yData(lPos) : lPos += 1
        Dim lTemplateID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If yAction = 0 Then
            'delete
            oPlayer.RemoveTemplate(lTemplateID)
        Else
            'add or update
            Dim oTemplate As RouteTemplate = Nothing
            If yAction = 255 Then
                oTemplate = New RouteTemplate
                oTemplate.lTemplateID = -1
            Else
                oTemplate = oPlayer.GetRouteTemplate(lTemplateID)
            End If

            With oTemplate
                ReDim .TemplateName(19)
                Array.Copy(yData, lPos, .TemplateName, 0, 20)
                lPos += 20

                .lPlayerID = oPlayer.ObjectID

                .lItemUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
                ReDim .uItems(.lItemUB)

                For X As Int32 = 0 To .lItemUB
                    Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim lLoadItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iLoadItemTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim yFlag As Byte = yData(lPos) : lPos += 1

                    With .uItems(X)
                        .oDest = CType(GetEpicaObject(lDestID, iDestTypeID), Epica_GUID)
                        .lLocX = lLocX
                        .lLocZ = lLocZ
                        .SetLoadItem(lLoadItemID, iLoadItemTypeID, oPlayer, yFlag)
                    End With
                Next X

                .SaveObject()
            End With

            If yAction = 255 Then
                oPlayer.AddTemplate(oTemplate)
            End If
        End If

        Dim yResp() As Byte = oPlayer.HandleGetRouteTemplateList()
        If yResp Is Nothing = False Then oPlayer.SendPlayerMessage(yResp, False, AliasingRights.eMoveUnits)

    End Sub
	Private Sub HandleEmailSettings(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		'Ok, first, ensure that the player id matches
		If lPlayerID <> mlClientPlayer(lIndex) Then Return

		'Ok, get our player object
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return
        If oPlayer.InMyDomain = False Then
            SendPassThruMsg(yData, lPlayerID, lPlayerID, ObjectType.ePlayer)
            Return
        End If

        Dim yResp() As Byte = DoEmailSettings(oPlayer, yData, 0)
        If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
    End Sub
    Private Sub HandleExceptionReport(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        Dim lPos As Int32 = 2
        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lLen > 25000 Then Return
        Dim sData As String = GetStringFromBytes(yData, lPos, lLen) : lPos += 20
        Dim lDTStamp As Int32 = CInt(Now.ToString("MMddhhmmss"))

        Dim oComm As OleDb.OleDbCommand = Nothing
        Try
            sData = MakeDBStr(sData)

            oComm = New OleDb.OleDbCommand("INSERT INTO tblExceptions (PlayerID, DTStamp, ExceptionReport) VALUES (" & lPlayerID & ", " & lDTStamp & ", '" & sData & "')", goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Throw New Exception("No records effected in execute!")
            End If
        Catch
            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            sPath &= "Crash_" & lDTStamp & "_" & lPlayerID & ".txt"
            Dim oFS As New IO.FileStream(sPath, IO.FileMode.Append)
            Dim oWrite As New IO.StreamWriter(oFS)
            oWrite.AutoFlush = True

            oWrite.Write(sData)
            oWrite.Close()
            oWrite.Dispose()
            oFS.Close()
            oFS.Dispose()
        End Try
        If oComm Is Nothing = False Then oComm.Dispose()
        oComm = Nothing
    End Sub
    Private Sub HandleForcefulDismantle(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim lPos As Int32 = 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For X As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2


            Dim oEntity As Epica_Entity = Nothing
            If iObjTypeID = ObjectType.eUnit Then
                oEntity = GetEpicaUnit(lObjID)
            ElseIf iObjTypeID = ObjectType.eFacility Then
                oEntity = GetEpicaFacility(lObjID)
            End If
            If oEntity Is Nothing Then Continue For
            If oEntity.Owner.ObjectID <> lPlayerID Then Continue For

            'Force dismantle does not guarantee caches
            'With oEntity
            '    Dim iTemp As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID

            '    Dim oDef As Epica_Entity_Def = Nothing
            '    If .ObjTypeID = ObjectType.eUnit Then
            '        oDef = CType(oEntity, Unit).EntityDef
            '    Else
            '        oDef = CType(oEntity, Facility).EntityDef
            '    End If
            '    If oDef Is Nothing Then Continue For

            '    If oDef.oPrototype Is Nothing = False AndAlso oDef.oPrototype.Owner Is Nothing = False AndAlso oDef.oPrototype.Owner.ObjectID <> 0 Then
            '        Dim oSocket As NetSock = Nothing
            '        If iTemp = ObjectType.ePlanet Then
            '            oSocket = CType(.ParentObject, Planet).oDomain.DomainSocket
            '        ElseIf iTemp = ObjectType.eSolarSystem Then
            '            oSocket = CType(.ParentObject, SolarSystem).oDomain.DomainSocket
            '        End If
            '        If oSocket Is Nothing = False Then
            '            Dim lParentID As Int32 = CType(.ParentObject, Epica_GUID).ObjectID

            '            Dim yCacheType As Byte = 0
            '            If (oDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
            '                yCacheType = yCacheType Or MineralCacheType.eGround
            '            ElseIf (oDef.yChassisType And ChassisType.eAtmospheric) <> 0 OrElse (oDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
            '                yCacheType = yCacheType Or MineralCacheType.eFlying
            '            ElseIf (oDef.yChassisType And ChassisType.eNavalBased) <> 0 Then
            '                yCacheType = yCacheType Or MineralCacheType.eNaval
            '            End If

            '            oDef.oPrototype.CreateDismantledCaches(oSocket, .CurrentStatus, .LocX, .LocZ, lParentID, iTemp, lArmorHP(0), lArmorHP(1), lArmorHP(2), lArmorHP(3), yCacheType)
            '        End If
            '    End If

            'End With

            DestroyEntity(oEntity, True, -1, True, "ForcefulDismantle")
        Next X

    End Sub
	Private Sub HandleGetAgentStatus(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) = 0 Then Return
		End If

		Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim oAgent As Agent = GetEpicaAgent(lAgentID)
		If oAgent Is Nothing Then Return
		If lPlayerID = oAgent.oOwner.ObjectID Then
			moClients(lIndex).SendData(oAgent.GetAgentStatusMsg())
		Else : LogEvent(LogEventType.PossibleCheat, "HandleSetAgentStatus Player Owner Mismatch: " & mlClientPlayer(lIndex))
		End If
	End Sub
	Private Sub HandleGetArchivedItems(ByRef yData() As Byte, ByVal lIndex As Int32)

		Dim lID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then lID = mlAliasedAs(lIndex)
		Dim bAliased As Boolean = GetEpicaPlayer(mlClientPlayer(lIndex)).lAliasingPlayerID > 0

		Dim oPlayer As Player = GetEpicaPlayer(lID)
		If oPlayer Is Nothing = False Then

			Dim yCache(200000) As Byte
			Dim yFinal() As Byte
			Dim lPos As Int32 = 0
			Dim lSingleMsgLen As Int32
			Dim yTemp() As Byte

			'Now, send our  Facility Defs
			If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eAddProduction) <> 0 Then
				For Y As Int32 = 0 To glFacilityDefUB
					If glFacilityDefIdx(Y) <> -1 AndAlso (goFacilityDef(Y).OwnerID = lID OrElse goFacilityDef(Y).OwnerID = 0) Then
                        If goFacilityDef(Y).oPrototype Is Nothing = False AndAlso goFacilityDef(Y).oPrototype.yArchived <> 1 Then Continue For
						yTemp = GetAddObjectMessage(goFacilityDef(Y), GlobalMessageCode.eGetArchivedItems)
						lSingleMsgLen = yTemp.Length
						'Ok, before we continue, check if we need to increase our cache
						If lPos + lSingleMsgLen + 2 > yCache.Length Then
							ReDim Preserve yCache(yCache.Length + 200000)
						End If
						System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
						lPos += 2
						yTemp.CopyTo(yCache, lPos)
						lPos += lSingleMsgLen

						If goFacilityDef(Y).ProductionRequirements Is Nothing = False Then
							yTemp = GetAddObjectMessage(goFacilityDef(Y).ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
							lSingleMsgLen = yTemp.Length
							'Ok, before we continue, check if we need to increase our cache
							If lPos + lSingleMsgLen + 2 > yCache.Length Then
								ReDim Preserve yCache(yCache.Length + 200000)
							End If
							System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
							lPos += 2
							yTemp.CopyTo(yCache, lPos)
							lPos += lSingleMsgLen
						End If
					End If
				Next Y
			End If
			If lPos <> 0 Then
				ReDim yFinal(lPos - 1)
				Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
				moClients(lIndex).SendLenAppendedData(yFinal)
			End If
			lPos = 0

			If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
				For Y As Int32 = 0 To oPlayer.mlAlloyUB
                    If oPlayer.mlAlloyIdx(Y) <> -1 AndAlso oPlayer.moAlloy(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moAlloy(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moAlloy(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moAlloy(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0

				For Y As Int32 = 0 To oPlayer.mlArmorUB
                    If oPlayer.mlArmorIdx(Y) <> -1 AndAlso oPlayer.moArmor(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moArmor(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moArmor(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moArmor(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0

				For Y As Int32 = 0 To oPlayer.mlEngineUB
                    If oPlayer.mlEngineIdx(Y) <> -1 AndAlso oPlayer.moEngine(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moEngine(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moEngine(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moEngine(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0

				For Y As Int32 = 0 To oPlayer.mlHullUB
                    If oPlayer.mlHullIdx(Y) <> -1 AndAlso oPlayer.moHull(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moHull(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moHull(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moHull(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0

				For Y As Int32 = 0 To oPlayer.mlPrototypeUB
                    If oPlayer.mlPrototypeIdx(Y) <> -1 AndAlso oPlayer.moPrototype(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moPrototype(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moPrototype(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moPrototype(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0

				For Y As Int32 = 0 To oPlayer.mlRadarUB
                    If oPlayer.mlRadarIdx(Y) <> -1 AndAlso oPlayer.moRadar(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moRadar(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moRadar(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moRadar(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0

				For Y As Int32 = 0 To oPlayer.mlShieldUB
                    If oPlayer.mlShieldIdx(Y) <> -1 AndAlso oPlayer.moShield(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moShield(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moShield(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moShield(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0

				For Y As Int32 = 0 To oPlayer.mlWeaponUB
                    If oPlayer.mlWeaponIdx(Y) <> -1 AndAlso oPlayer.moWeapon(Y).yArchived = 1 Then
                        yTemp = GetAddObjectMessage(oPlayer.moWeapon(Y), GlobalMessageCode.eGetArchivedItems)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moWeapon(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                        yTemp = GetAddObjectMessage(oPlayer.moWeapon(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
				If lPos <> 0 Then
					ReDim yFinal(lPos - 1)
					Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
					moClients(lIndex).SendLenAppendedData(yFinal)
				End If
				lPos = 0
			End If

			'and player intel
			If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) <> 0 Then
				If oPlayer.moItemIntel Is Nothing = False Then
					For Y As Int32 = 0 To Math.Min(oPlayer.mlItemIntelUB, oPlayer.moItemIntel.GetUpperBound(0))
                        If oPlayer.moItemIntel(Y) Is Nothing = False AndAlso oPlayer.moItemIntel(Y).yArchived = 1 Then
                            yTemp = GetAddObjectMessage(oPlayer.moItemIntel(Y), GlobalMessageCode.eGetArchivedItems)
                            If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                lSingleMsgLen = yTemp.Length
                                'Ok, before we continue, check if we need to increase our cache
                                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                    ReDim Preserve yCache(yCache.Length + 200000)
                                End If
                                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                lPos += 2
                                yTemp.CopyTo(yCache, lPos)
                                lPos += lSingleMsgLen
                            End If
                        End If
					Next Y
				End If

				'Send the player tech knowledges
				For Y As Int32 = 0 To oPlayer.mlPlayerTechKnowledgeUB
                    If oPlayer.myPlayerTechKnowledgeUsed(Y) <> 0 AndAlso oPlayer.moPlayerTechKnowledge(Y).yArchived = 1 Then
                        yTemp = oPlayer.moPlayerTechKnowledge(Y).GetAddMsg()
                        If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
				Next Y
			End If
			If lPos <> 0 Then
				ReDim yFinal(lPos - 1)
				Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
				moClients(lIndex).SendLenAppendedData(yFinal)
			End If
			lPos = 0

			If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) <> 0 Then
				For Y As Int32 = 0 To glPlayerMissionUB
					If glPlayerMissionIdx(Y) <> -1 AndAlso goPlayerMission(Y).oPlayer.ObjectID = oPlayer.ObjectID Then
                        If goPlayerMission(Y).yArchived <> 1 Then Continue For
						yTemp = goPlayerMission(Y).GetAddObjectMessage
						If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
							lSingleMsgLen = yTemp.Length
							'Ok, before we continue, check if we need to increase our cache
							If lPos + lSingleMsgLen + 2 > yCache.Length Then
								ReDim Preserve yCache(yCache.Length + 200000)
							End If
							System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
							lPos += 2
							yTemp.CopyTo(yCache, lPos)
							lPos += lSingleMsgLen
						End If
					End If
				Next Y
			End If
			If lPos <> 0 Then
				ReDim yFinal(lPos - 1)
				Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
				moClients(lIndex).SendLenAppendedData(yFinal)
			End If
			lPos = 0

			'Send our unit defs
			If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eAddProduction) <> 0 Then
				For Y As Int32 = 0 To glUnitDefUB
					If glUnitDefIdx(Y) <> -1 AndAlso (goUnitDef(Y).OwnerID = lID OrElse goUnitDef(Y).OwnerID = 0) Then
                        If goUnitDef(Y).oPrototype Is Nothing OrElse goUnitDef(Y).oPrototype.yArchived <> 1 Then Continue For
						yTemp = GetAddObjectMessage(goUnitDef(Y), GlobalMessageCode.eGetArchivedItems)
						If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
							lSingleMsgLen = yTemp.Length
							'Ok, before we continue, check if we need to increase our cache
							If lPos + lSingleMsgLen + 2 > yCache.Length Then
								ReDim Preserve yCache(yCache.Length + 200000)
							End If
							System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
							lPos += 2
							yTemp.CopyTo(yCache, lPos)
							lPos += lSingleMsgLen
						End If
						If goUnitDef(Y).ProductionRequirements Is Nothing = False Then
							yTemp = GetAddObjectMessage(goUnitDef(Y).ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
							If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
								lSingleMsgLen = yTemp.Length
								'Ok, before we continue, check if we need to increase our cache
								If lPos + lSingleMsgLen + 2 > yCache.Length Then
									ReDim Preserve yCache(yCache.Length + 200000)
								End If
								System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
								lPos += 2
								yTemp.CopyTo(yCache, lPos)
								lPos += lSingleMsgLen
							End If
						End If
					End If
				Next Y
			End If
			If lPos <> 0 Then
				ReDim yFinal(lPos - 1)
				Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
				moClients(lIndex).SendLenAppendedData(yFinal)
			End If
			lPos = 0

			'Now, send the initial request back to the client so that it can send us a new request (if needed)
			' Our Player Has REquested Details flag has been set above in the last message sent to the client
			' so the anti-hack mechanism should be in place already
			'oSocket.SendData(yData)
			'lSingleMsgLen = yData.Length
			''Ok, before we continue, check if we need to increase our cache
			'If lPos + lSingleMsgLen + 2 > yCache.Length Then
			'	ReDim Preserve yCache(yCache.Length + 200000)
			'End If
			'System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
			'lPos += 2
			'yData.CopyTo(yCache, lPos)
			'lPos += lSingleMsgLen

			'Now, send it all...
			'If lPos <> 0 Then
			'	ReDim yFinal(lPos - 1)
			'	Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
			'	moClients(lIndex).SendLenAppendedData(yFinal)
			'End If

		End If
	End Sub
	Private Sub HandleGetAvailableResources(ByVal lIndex As Int32, ByVal yData() As Byte)
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eViewTreasury) = 0 Then Return
		End If
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		If oPlayer.lLastAvailableResourcesRequest = Int32.MinValue OrElse glCurrentCycle - oPlayer.lLastAvailableResourcesRequest > ml_AVAILABLE_RESOURCES_REQUEST_TIME Then
			oPlayer.lLastAvailableResourcesRequest = glCurrentCycle
		Else : Return
		End If

		'Get the colony object
		Dim oColony As Colony = Nothing
		If iObjTypeID = ObjectType.eColony Then
			oColony = GetEpicaColony(lObjID)
		ElseIf iObjTypeID = ObjectType.ePlanet Then
			Dim lIdx As Int32 = oPlayer.GetColonyFromParent(lObjID, iObjTypeID)
			If lIdx <> -1 Then oColony = goColony(lIdx)
		Else
			If iObjTypeID <> ObjectType.eUnit AndAlso iObjTypeID <> ObjectType.eFacility Then
				LogEvent(LogEventType.PossibleCheat, "HandleGetAvailableResources Unexpected TypeID: " & lObjID & ", " & iObjTypeID & ". PlayerID: " & mlClientPlayer(lIndex))
				Return
			End If

			Dim oEntity As Epica_Entity = CType(GetEpicaObject(lObjID, iObjTypeID), Epica_Entity)
			Dim lIdx As Int32 = -1
			If oEntity Is Nothing = False Then
				With CType(oEntity.ParentObject, Epica_GUID)
					If .ObjTypeID = ObjectType.eSolarSystem Then
						'space station, so the colony's Parent is the entity
						lIdx = oEntity.Owner.GetColonyFromParent(oEntity.ObjectID, oEntity.ObjTypeID)
					Else
						'Planet, so the colony's Parent is the entity's parent
						lIdx = oEntity.Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
					End If
				End With
			End If
			If lIdx <> -1 Then oColony = goColony(lIdx)
		End If

		If oColony Is Nothing Then
			LogEvent(LogEventType.PossibleCheat, "HandleGetAvailableResources returned no colony! PlayerID: " & mlClientPlayer(lIndex))
			Return
		End If

		Dim X As Int32

        'Dim oCargo As Epica_GUID
		Dim lPos As Int32

		'ok... here we go...
		With oColony
			Dim lObjIDs() As Int32 = Nothing
			Dim iObjTypeIDs() As Int16 = Nothing
			Dim lQtyAvails() As Int32 = Nothing
            'Dim iTmpTypeID As Int16
			Dim lListUB As Int32 = -1

			Dim blCargoCapTotal As Int64 = 0
			Dim blCargoCapAvail As Int64 = 0

			For X = 0 To .mlMineralCacheUB
				If .mlMineralCacheIdx(X) > -1 AndAlso .mlMineralCacheID(X) = glMineralCacheIdx(.mlMineralCacheIdx(X)) Then
                    If goMineralCache(.mlMineralCacheIdx(X)).Quantity < 1 Then Continue For

                    'find it in our array
					Dim bFound As Boolean = False
					For Y As Int32 = 0 To lListUB
						If lObjIDs(Y) = .mlMineralCacheMineralID(X) AndAlso iObjTypeIDs(Y) = ObjectType.eMineral Then
							lQtyAvails(Y) += goMineralCache(.mlMineralCacheIdx(X)).Quantity
							bFound = True
							Exit For
						End If
					Next Y

					If bFound = False Then
						lListUB += 1
						ReDim Preserve lObjIDs(lListUB)
						ReDim Preserve iObjTypeIDs(lListUB)
						ReDim Preserve lQtyAvails(lListUB)
						lObjIDs(lListUB) = .mlMineralCacheMineralID(X)
						iObjTypeIDs(lListUB) = ObjectType.eMineral
						lQtyAvails(lListUB) = goMineralCache(.mlMineralCacheIdx(X)).Quantity
					End If
				End If
            Next X
            For X = 0 To .mlComponentCacheUB
                If .mlComponentCacheIdx(X) > -1 AndAlso .mlComponentCacheID(X) = glComponentCacheIdx(.mlComponentCacheIdx(X)) Then
                    If goComponentCache(.mlComponentCacheIdx(X)).Quantity < 1 Then Continue For
                    'find it in our array
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To lListUB
                        If lObjIDs(Y) = .mlComponentCacheCompID(X) AndAlso iObjTypeIDs(Y) = goComponentCache(.mlComponentCacheIdx(X)).ComponentTypeID Then
                            lQtyAvails(Y) += goComponentCache(.mlComponentCacheIdx(X)).Quantity
                            bFound = True
                            Exit For
                        End If
                    Next Y

                    If bFound = False Then
                        lListUB += 1
                        ReDim Preserve lObjIDs(lListUB)
                        ReDim Preserve iObjTypeIDs(lListUB)
                        ReDim Preserve lQtyAvails(lListUB)
                        lObjIDs(lListUB) = .mlComponentCacheCompID(X)
                        iObjTypeIDs(lListUB) = goComponentCache(.mlComponentCacheIdx(X)).ComponentTypeID
                        lQtyAvails(lListUB) = goComponentCache(.mlComponentCacheIdx(X)).Quantity
                    End If
                End If
            Next X

            'For X = 0 To .ChildrenUB
            '	If .lChildrenIdx(X) <> -1 Then
            '                 If .oChildren(X).Active AndAlso ((.oChildren(X).CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0) AndAlso Colony.ProductionTypeSharesColonyCargo(.oChildren(X).yProductionType) = True Then

            '                     blCargoCapAvail += .oChildren(X).Cargo_Cap
            '                     blCargoCapTotal += .oChildren(X).EntityDef.Cargo_Cap

            '                     For lCIdx As Int32 = 0 To .oChildren(X).lCargoUB
            '                         If .oChildren(X).lCargoIdx(lCIdx) <> -1 Then
            '                             'Now, check if the content is a Tech or mineral cache
            '                             iTmpTypeID = .oChildren(X).oCargoContents(lCIdx).ObjTypeID
            '                             oCargo = .oChildren(X).oCargoContents(lCIdx)

            '                             Dim lID As Int32 = -1
            '                             Dim iTypeID As Int16 = -1
            '                             Dim lQty As Int32 = 0
            '                             Select Case iTmpTypeID
            '                                 Case ObjectType.eMineralCache
            '                                     With CType(oCargo, MineralCache)
            '                                         lID = .oMineral.ObjectID
            '                                         iTypeID = ObjectType.eMineral
            '                                         lQty = .Quantity
            '                                     End With
            '                                 Case ObjectType.eComponentCache
            '                                     With CType(oCargo, ComponentCache)
            '                                         lID = .ComponentID
            '                                         iTypeID = .ComponentTypeID
            '                                         lQty = .Quantity
            '                                     End With
            '                             End Select

            '                             If lQty <> 0 Then
            '                                 'find it in our array
            '                                 For Y As Int32 = 0 To lListUB
            '                                     If lObjIDs(Y) = lID AndAlso iObjTypeIDs(Y) = iTypeID Then
            '                                         lQtyAvails(Y) += lQty
            '                                         lQty = 0
            '                                         Exit For
            '                                     End If
            '                                 Next Y

            '                                 If lQty <> 0 Then
            '                                     lListUB += 1
            '                                     ReDim Preserve lObjIDs(lListUB)
            '                                     ReDim Preserve iObjTypeIDs(lListUB)
            '                                     ReDim Preserve lQtyAvails(lListUB)
            '                                     lObjIDs(lListUB) = lID
            '                                     iObjTypeIDs(lListUB) = iTypeID
            '                                     lQtyAvails(lListUB) = lQty
            '                                 End If
            '                             End If
            '                         End If
            '                     Next
            '                 End If
            '	End If
            'Next X

            blCargoCapTotal = .TotalCargoCapAvailable
            blCargoCapAvail = .Cargo_Cap

			'Now, put our message together...
			Dim yResp(43 + ((lListUB + 1) * 10)) As Byte

			System.BitConverter.GetBytes(GlobalMessageCode.eGetAvailableResources).CopyTo(yResp, 0)
			CType(.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yResp, 2)
            lPos = 8

            'need to reduce the colony's cargo used from the blCargoCapAvail
            'blCargoCapAvail -= .TotalCargoCapUsed

			System.BitConverter.GetBytes(.ObjectID).CopyTo(yResp, lPos) : lPos += 4
			System.BitConverter.GetBytes(.Population).CopyTo(yResp, lPos) : lPos += 4
			System.BitConverter.GetBytes(.ColonyEnlisted).CopyTo(yResp, lPos) : lPos += 4
			System.BitConverter.GetBytes(.ColonyOfficers).CopyTo(yResp, lPos) : lPos += 4
			System.BitConverter.GetBytes(blCargoCapTotal).CopyTo(yResp, lPos) : lPos += 8
			System.BitConverter.GetBytes(blCargoCapAvail).CopyTo(yResp, lPos) : lPos += 8

			'Now, how many objects we sending?
			System.BitConverter.GetBytes(lListUB + 1I).CopyTo(yResp, lPos) : lPos += 4

			For X = 0 To lListUB
				System.BitConverter.GetBytes(lObjIDs(X)).CopyTo(yResp, lPos) : lPos += 4
				System.BitConverter.GetBytes(iObjTypeIDs(X)).CopyTo(yResp, lPos) : lPos += 2
				System.BitConverter.GetBytes(lQtyAvails(X)).CopyTo(yResp, lPos) : lPos += 4
			Next X
			moClients(lIndex).SendData(yResp)

		End With
	End Sub
	Private Sub HandleGetColonyChildList(ByRef yData() As Byte, ByVal lClientIndex As Int32)
		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 8)

		Dim oColony As Colony = Nothing
		Dim yResp() As Byte = Nothing
		Dim lPos As Int32

		'TODO: ok get the colony where the owner id = lplayerid and the parent = the envir

		If iEnvirTypeID = ObjectType.eFacility Then
			Dim oTmpFac As Facility = GetEpicaFacility(lEnvirID)
			If oTmpFac Is Nothing = False Then
				oColony = oTmpFac.ParentColony
			End If
		End If

		If oColony Is Nothing Then Return

		Dim lCnt As Int32 = 0

		For X As Int32 = 0 To oColony.ChildrenUB
			If oColony.lChildrenIdx(X) <> -1 Then lCnt += 1
		Next X

		'Now, do our deal, check if the owner is the related player ID of the socket
		lPos = 16
		If mlClientPlayer(lClientIndex) = oColony.Owner.ObjectID OrElse (mlAliasedAs(lClientIndex) = oColony.Owner.ObjectID AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eViewColonyStats) <> 0) Then
			'Getting MY list, so I should already know the Entity Defs
			Dim lHangarCap As Int32 = 0
			Dim lCargoCap As Int32 = 0

			If iEnvirTypeID = ObjectType.eFacility Then
				'Space station, we are going to send Cargo and Hangar Cap too as they change
				ReDim yResp(23 + (16 * lCnt))
			Else
				ReDim yResp(15 + (16 * lCnt))
			End If

			yData.CopyTo(yResp, 0)

			System.BitConverter.GetBytes(lCnt).CopyTo(yResp, 12)
			For X As Int32 = 0 To oColony.ChildrenUB
				If oColony.lChildrenIdx(X) <> -1 Then
					With oColony.oChildren(X)
						.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
						.EntityDef.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
						System.BitConverter.GetBytes(.CurrentStatus).CopyTo(yResp, lPos) : lPos += 4

						If iEnvirTypeID = ObjectType.eFacility Then
							lHangarCap += .EntityDef.Hangar_Cap
							lCargoCap += .EntityDef.Cargo_Cap
						End If
					End With
				End If
			Next X

			If iEnvirTypeID = ObjectType.eFacility Then
				'ok, add our HANGAR Cap
				System.BitConverter.GetBytes(lHangarCap).CopyTo(yResp, lPos) : lPos += 4
				'and then our CARGO CAP
				System.BitConverter.GetBytes(lCargoCap).CopyTo(yResp, lPos) : lPos += 4
			End If
		Else
			'TODO: Define what happens here...
		End If

		If yResp Is Nothing = False Then moClients(lClientIndex).SendData(yResp)
	End Sub
	Private Sub HandleGetColonyDetails(ByVal yData() As Byte, ByVal lIndex As Int32)
		Dim X As Int32
		Dim oColony As Colony = Nothing
		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

		Dim yResp() As Byte

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eViewColonyStats) = 0 Then Return
		End If

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing = False Then
			X = oPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)
			If X <> -1 Then oColony = goColony(X)
		End If

		If oColony Is Nothing = False Then
			ReDim yResp(84)

			System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyDetails).CopyTo(yResp, 0)

			With oColony
				.GetGUIDAsString.CopyTo(yResp, 2)
				System.BitConverter.GetBytes(.Population).CopyTo(yResp, 8)
				System.BitConverter.GetBytes(.NumberOfJobs).CopyTo(yResp, 12)
				System.BitConverter.GetBytes(.ColonyEnlisted).CopyTo(yResp, 16)
				System.BitConverter.GetBytes(.ColonyOfficers).CopyTo(yResp, 20)
				System.BitConverter.GetBytes(.PowerGeneration).CopyTo(yResp, 24)
				System.BitConverter.GetBytes(.PowerConsumption).CopyTo(yResp, 28)
				.ColonyName.CopyTo(yResp, 32)
				yResp(52) = .UnemploymentRate
				System.BitConverter.GetBytes(.PoweredHousing).CopyTo(yResp, 53)
				System.BitConverter.GetBytes(.UnpoweredHousing).CopyTo(yResp, 57)
				'System.BitConverter.GetBytes(CShort(.MoraleMultiplier * 100)).CopyTo(yResp, 61)
				'System.BitConverter.GetBytes(.ColonyGrowthRate).CopyTo(yResp, 63)
				System.BitConverter.GetBytes(CInt(.MoraleMultiplier * 100)).CopyTo(yResp, 61)
				System.BitConverter.GetBytes(CShort(.ColonyGrowthRate)).CopyTo(yResp, 65)
				yResp(67) = .TaxRate
				System.BitConverter.GetBytes(.TotalPowerNeeded).CopyTo(yResp, 68)
				yResp(72) = .GetColonyIntelligence()
				System.BitConverter.GetBytes(.GetResearchJobCount).CopyTo(yResp, 73)

				Dim lValue As Int32 = 0
				If .Owner.PlayerIsAtWar = True Then
					If .Owner.lWarSentiment < 0 Then
						lValue = .Owner.lWarSentiment
					End If
				Else
					If .Owner.lWarSentiment > 0 Then
						lValue = .Owner.lWarSentiment
					End If
				End If
				System.BitConverter.GetBytes(lValue).CopyTo(yResp, 77)
				System.BitConverter.GetBytes(.iControlledGrowth).CopyTo(yResp, 81)
				System.BitConverter.GetBytes(.iControlledMorale).CopyTo(yResp, 83)
			End With

			moClients(lIndex).SendData(yResp)
		End If

	End Sub
	Private Sub HandleGetColonyList(ByRef yData() As Byte, ByVal lIndex As Int32)
		'MsgCode (2), PlayerID (4)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
		If lPlayerID < 1 Then Return
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		Dim lPos As Int32

		If lPlayerID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = lPlayerID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewColonyStats) <> 0) Then
			Dim lCnt As Int32 = 0
			For X As Int32 = 0 To oPlayer.mlColonyUB
				If oPlayer.mlColonyIdx(X) <> -1 Then
					'Ok, mlColonyIdx(X) is a Index of the global Colony array
					If glColonyIdx(oPlayer.mlColonyIdx(X)) = oPlayer.mlColonyID(X) Then
						lCnt += 1
					End If
				End If
			Next X

			If lCnt = 0 Then Return

			Dim yMsg(9 + (lCnt * 30)) As Byte
			lPos = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyList).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

			For X As Int32 = 0 To oPlayer.mlColonyUB
				If oPlayer.mlColonyIdx(X) <> -1 Then
					'Ok, mlColonyIdx(X) is a Index of the global Colony array
					If glColonyIdx(oPlayer.mlColonyIdx(X)) = oPlayer.mlColonyID(X) Then
						Dim oColony As Colony = goColony(oPlayer.mlColonyIdx(X))
						If oColony Is Nothing = False Then
							With oColony
								System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
								.ColonyName.CopyTo(yMsg, lPos) : lPos += 20

								Select Case CType(.ParentObject, Epica_GUID).ObjTypeID
									Case ObjectType.eFacility
										'Ok, must be a space station, so we get the parent of the parent
										CType(CType(.ParentObject, Facility).ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos)
									Case Else
										CType(.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos)
								End Select
								lPos += 6
							End With
						End If
					End If
				End If
			Next X

			moClients(lIndex).SendData(yMsg)
		Else
            LogEvent(LogEventType.PossibleCheat, "HandleGetColonyList PlayerID Mismatch: " & mlClientPlayer(lIndex))
		End If

	End Sub
	Private Sub HandleGetColonyResearchList(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2	'for msgcode

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eAddResearch) = 0 Then Return
		End If

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing = False Then
			Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

			Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)
			If lColonyIdx > -1 AndAlso glColonyIdx(lColonyIdx) > -1 Then
				Dim oColony As Colony = goColony(lColonyIdx)
				If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = oPlayer.ObjectID Then
					Dim yResp() As Byte = oColony.HandleGetColonyResearchList()
					If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
				End If
			End If
		End If

    End Sub
    Private Sub HandleGetCPPenaltyList(ByVal yData() As Byte, ByVal lIndex As Int32)
        'Ok, get the player to whom we reference
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If lPlayerID <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
            If (mlAliasedRights(lIndex) And AliasingRights.eViewUnitsAndFacilities) = 0 Then Return
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            oPlayer.SendCPPenaltyList()
        End If
    End Sub
    Private Sub HandleGetDAValues(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim ySubTypeID As Byte = yData(lPos) : lPos += 1
        Dim yHullTypeID As Byte = yData(lPos) : lPos += 1
        Dim lMin1ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMin2ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMin3ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMin4ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMin5ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMin6ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lMin1DA As Int32 = 0
        Dim lMin2DA As Int32 = 0
        Dim lMin3DA As Int32 = 0
        Dim lMin4DA As Int32 = 0
        Dim lMin5DA As Int32 = 0
        Dim lMin6DA As Int32 = 0

        Select Case iTypeID
            Case ObjectType.eEngineTech
                Dim oTech As New EngineTechComputer()
                With oTech
                    .lHullTypeID = yHullTypeID
                    lMin1DA = .GetDADiff(lMin1ID, 0)
                    lMin2DA = .GetDADiff(lMin2ID, 1)
                    lMin3DA = .GetDADiff(lMin3ID, 2)
                    lMin4DA = .GetDADiff(lMin4ID, 3)
                    lMin5DA = .GetDADiff(lMin5ID, 4)
                    lMin6DA = .GetDADiff(lMin6ID, 5)
                End With
                oTech = Nothing
            Case ObjectType.eRadarTech
                Dim oTech As New RadarTechComputer()
                With oTech
                    .lHullTypeID = yHullTypeID
                    lMin1DA = .GetDADiff(lMin1ID, 0)
                    lMin2DA = .GetDADiff(lMin2ID, 1)
                    lMin3DA = .GetDADiff(lMin3ID, 2)
                    lMin4DA = .GetDADiff(lMin4ID, 3)
                    lMin5DA = .GetDADiff(lMin5ID, 4)
                    lMin6DA = .GetDADiff(lMin6ID, 5)
                End With
                oTech = Nothing
            Case ObjectType.eShieldTech
                Dim oTech As New ShieldTechComputer()
                With oTech
                    .lHullTypeID = yHullTypeID
                    lMin1DA = .GetDADiff(lMin1ID, 0)
                    lMin2DA = .GetDADiff(lMin2ID, 1)
                    lMin3DA = .GetDADiff(lMin3ID, 2)
                    lMin4DA = .GetDADiff(lMin4ID, 3)
                    lMin5DA = .GetDADiff(lMin5ID, 4)
                    lMin6DA = .GetDADiff(lMin6ID, 5)
                End With
                oTech = Nothing
            Case ObjectType.eWeaponTech
                Select Case ySubTypeID
                    Case WeaponClassType.eBomb
                        Dim yPayload As Byte = yData(lPos) : lPos += 1
                        Dim oTech As New BombTechComputer()
                        With oTech
                            .lHullTypeID = yHullTypeID
                            .yPayloadType = yPayload
                            lMin1DA = .GetDADiff(lMin1ID, 0)
                            lMin2DA = .GetDADiff(lMin2ID, 1)
                            lMin3DA = .GetDADiff(lMin3ID, 2)
                            lMin4DA = .GetDADiff(lMin4ID, 3)
                            lMin5DA = .GetDADiff(lMin5ID, 4)
                            lMin6DA = .GetDADiff(lMin6ID, 5)
                        End With
                        oTech = Nothing
                    Case WeaponClassType.eEnergyBeam
                        Dim oTech As New SolidTechComputer()
                        With oTech
                            .lHullTypeID = yHullTypeID
                            lMin1DA = .GetDADiff(lMin1ID, 0)
                            lMin2DA = .GetDADiff(lMin2ID, 1)
                            lMin3DA = .GetDADiff(lMin3ID, 2)
                            lMin4DA = .GetDADiff(lMin4ID, 3)
                            lMin5DA = .GetDADiff(lMin5ID, 4)
                            lMin6DA = .GetDADiff(lMin6ID, 5)
                        End With
                        oTech = Nothing
                    Case WeaponClassType.eEnergyPulse
                        Dim oTech As New PulseTechComputer()
                        With oTech
                            .lHullTypeID = yHullTypeID
                            lMin1DA = .GetDADiff(lMin1ID, 0)
                            lMin2DA = .GetDADiff(lMin2ID, 1)
                            lMin3DA = .GetDADiff(lMin3ID, 2)
                            lMin4DA = .GetDADiff(lMin4ID, 3)
                            lMin5DA = .GetDADiff(lMin5ID, 4)
                            lMin6DA = .GetDADiff(lMin6ID, 5)
                        End With
                        oTech = Nothing
                    Case WeaponClassType.eMissile
                        Dim yPayload As Byte = yData(lPos) : lPos += 1
                        Dim oTech As New MissileTechComputer()
                        With oTech
                            .lHullTypeID = yHullTypeID
                            .yPayloadType = yPayload
                            lMin1DA = .GetDADiff(lMin1ID, 0)
                            lMin2DA = .GetDADiff(lMin2ID, 1)
                            lMin3DA = .GetDADiff(lMin3ID, 2)
                            lMin4DA = .GetDADiff(lMin4ID, 3)
                            lMin5DA = .GetDADiff(lMin5ID, 4)
                            lMin6DA = .GetDADiff(lMin6ID, 5)
                        End With
                        oTech = Nothing
                    Case WeaponClassType.eProjectile
                        Dim yPayload As Byte = yData(lPos) : lPos += 1
                        Dim yProj As Byte = yData(lPos) : lPos += 1
                        Dim oTech As New ProjectileTechComputer()
                        With oTech
                            .lHullTypeID = yHullTypeID
                            .yPayloadType = yPayload
                            .yProjectionType = yProj
                            lMin1DA = .GetDADiff(lMin1ID, 0)
                            lMin2DA = .GetDADiff(lMin2ID, 1)
                            lMin3DA = .GetDADiff(lMin3ID, 2)
                            lMin4DA = .GetDADiff(lMin4ID, 3)
                            lMin5DA = .GetDADiff(lMin5ID, 4)
                            lMin6DA = .GetDADiff(lMin6ID, 5)
                        End With
                        oTech = Nothing
                End Select
        End Select

        Dim oRand As New Random()
        Dim ySeedCoeff As Byte = CByte(oRand.Next(0, 256))
        oRand = Nothing

        Dim lSeedVal As Int32 = (CInt(iTypeID) + 1) * (CInt(ySubTypeID) + 1) * (CInt(yHullTypeID) + 1) * (CInt(ySeedCoeff) + 1)

        oRand = New Random(lSeedVal)
        Dim lMaxVal As Int32 = Int32.MaxValue \ 2
        lMin1DA += oRand.Next(0, lMaxVal)
        lMin2DA += oRand.Next(0, lMaxVal)
        lMin3DA += oRand.Next(0, lMaxVal)
        lMin4DA += oRand.Next(0, lMaxVal)
        lMin5DA += oRand.Next(0, lMaxVal)
        lMin6DA += oRand.Next(0, lMaxVal)
        oRand = Nothing

        Dim yMsg(30) As Byte
        lPos = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetDAValues).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
        yMsg(lPos) = ySubTypeID : lPos += 1
        yMsg(lPos) = yHullTypeID : lPos += 1
        yMsg(lPos) = ySeedCoeff : lPos += 1
        System.BitConverter.GetBytes(lMin1DA).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin2DA).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin3DA).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin4DA).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin5DA).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin6DA).CopyTo(yMsg, lPos) : lPos += 4

        moClients(lIndex).SendData(yMsg)
    End Sub
	Private Sub HandleGetEntityDetails(ByRef yData() As Byte, ByVal lIndex As Int32)
		'Ok, the player wants to know about something...
		Dim lPos As Int32 = 2
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		'TODO: Determine what else needs to go here...
		Select Case iObjTypeID
			Case ObjectType.eMineralCache
				Dim oCache As MineralCache = GetEpicaMineralCache(lObjID)
				If oCache Is Nothing = False Then
					Dim yResp(28) As Byte

					With oCache
						lPos = 0
						System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yResp, lPos) : lPos += 2
						.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
						System.BitConverter.GetBytes(.LocX).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(.LocZ).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(.oMineral.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                        yResp(lPos) = .CacheTypeID : lPos += 1
						System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(.Concentration).CopyTo(yResp, lPos) : lPos += 4
					End With

					If mb_MONITOR_MSG_ACTIVITY = True Then Me.moMonitor.AddOutMsg(GlobalMessageCode.eGetEntityDetails, MsgMonitor.eMM_AppType.ClientConnection, yResp.Length, mlClientPlayer(lIndex))
					moClients(lIndex).SendData(yResp)
				End If
		End Select

	End Sub
	Private Sub HandleGetEntityName(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4


        For X As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim oObj As Object = GetEpicaObject(lObjID, iObjTypeID)
            If oObj Is Nothing Then Continue For

            If iObjTypeID = ObjectType.eGuild Then
                Dim yResp(57) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                GetEpicaObjectName(iObjTypeID, oObj).CopyTo(yResp, 8)
                moClients(lSocketIndex).SendData(yResp)
            Else
                Dim yResp(27) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                GetEpicaObjectName(iObjTypeID, oObj).CopyTo(yResp, 8)
                moClients(lSocketIndex).SendData(yResp)
            End If

        Next X

        '      Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        'Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        'Dim oObj As Object = GetEpicaObject(lObjID, iObjTypeID)

        'If oObj Is Nothing Then Return

        'Dim yResp(27) As Byte

        'yData.CopyTo(yResp, 0)

        'GetEpicaObjectName(iObjTypeID, oObj).CopyTo(yResp, 8)
        'moClients(lSocketIndex).SendData(yResp)
	End Sub
	Private Sub HandleGetEnvirConstructionObjects(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lOwnerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lOwnerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eViewUnitsAndFacilities) = 0 Then Return
		End If

		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

		'TODO: here, we would determine if the player has any special techs that would alter the results...

		GetEnvirConstObjects(lEnvirID, iEnvirTypeID, lOwnerID, moClients(lIndex))

		Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
		If oPlayer Is Nothing = False Then
			If oPlayer.sDXDiag = "" AndAlso oPlayer.bDXDiagRequested = False Then
				oPlayer.bDXDiagRequested = True
				Dim yMsg(1) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eRequestDXDiag).CopyTo(yMsg, 0)
				moClients(lIndex).SendData(yMsg)
			End If
		End If

	End Sub
	Private Sub HandleGetGuildAssets(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

		Dim yResp() As Byte = oPlayer.oGuild.GetGuildAssetsMsg()
		If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
    End Sub
    Private Sub HandleGetGuildBillboards(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iTypeID <> ObjectType.ePlanet Then Return

        Dim oPlanet As Planet = GetEpicaPlanet(lPlanetID)
        If oPlanet Is Nothing Then Return

        moClients(lIndex).SendData(oPlanet.GetBillBoardsResponse())
    End Sub
    Private Sub HandleGetGuildInvites(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To oPlayer.lInvitedToJoinUB
            If oPlayer.lInvitedToJoins(X) > 0 Then
                lCnt += 1
            End If
        Next X

        Dim yMsg(9 + (70 * lCnt)) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildInvites).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        For X As Int32 = 0 To oPlayer.lInvitedToJoinUB
            If oPlayer.lInvitedToJoins(X) > 0 Then
                Dim oGuild As Guild = GetEpicaGuild(oPlayer.lInvitedToJoins(X))
                If oGuild Is Nothing = False Then
                    With oGuild
                        System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        .yGuildName.CopyTo(yMsg, lPos) : lPos += 50
                        Dim lVal As Int32 = 0
                        If .dtFormed <> Date.MinValue Then lVal = GetDateAsNumber(.dtFormed)
                        System.BitConverter.GetBytes(lVal).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.iRecruitFlags).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(.lBaseGuildFlags).CopyTo(yMsg, lPos) : lPos += 4
                        yMsg(lPos) = .yVoteWeightType : lPos += 1
                        yMsg(lPos) = .yGuildTaxRateInterval : lPos += 1
                        System.BitConverter.GetBytes(.lIcon).CopyTo(yMsg, lPos) : lPos += 4
                    End With
                End If
            End If
        Next X

        moClients(lIndex).SendData(yMsg)



    End Sub
    Private Sub HandleGetImposedAgentEffects(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) <> -1 Then
            If (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) = 0 Then
                LogEvent(LogEventType.PossibleCheat, "HandleGetImposedAgentEffects, player has no rights: " & mlClientPlayer(lIndex))
                Return
            Else : lPlayerID = mlAliasedAs(lIndex)
            End If
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim yResp() As Byte = oPlayer.GetImposedAgentListMsg()
            If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
        End If
    End Sub
	Private Sub HandleGetIntelSellOrderDetail(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lTP_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		
		Dim oPlayer As Player = Nothing
		If lTP_ID < 1 Then
			oPlayer = GetEpicaPlayer(Math.Abs(lTP_ID))
		Else
			Dim oFac As Facility = GetEpicaFacility(lTP_ID)
			If oFac Is Nothing = False Then
				oPlayer = oFac.Owner
			End If
		End If

        If oPlayer Is Nothing = False Then

            If oPlayer.InMyDomain = False Then
                SendPassThruMsg(yData, mlClientPlayer(lIndex), oPlayer.ObjectID, oPlayer.ObjTypeID)
                Return
            End If

            Dim yResp() As Byte = DoGetIntelSellOrderDetail(oPlayer, yData, 0)
            If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)

        End If
	End Sub
	Private Sub HandleGetItemIntelDetail(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iItemTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lOtherPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) <> -1 Then
			If (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) = 0 Then
				LogEvent(LogEventType.PossibleCheat, "HandleGetItemIntelDetail, player has no rights: " & mlClientPlayer(lIndex))
				Return
			Else : lPlayerID = mlAliasedAs(lIndex)
			End If
		End If

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing = False Then
			Dim oPII As PlayerItemIntel = oPlayer.GetPlayerItemIntel(lItemID, iItemTypeID, lOtherPlayerID)
			If oPII Is Nothing = False Then
				moClients(lIndex).SendData(oPII.GetDetails)
			Else
				LogEvent(LogEventType.PossibleCheat, "HandleGetItemIntelDetail, player does not item intel: " & mlClientPlayer(lIndex))
				Return
			End If
		End If
	End Sub
	Private Sub HandleGetMyVoteValue(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		Dim lPos As Int32 = 2
		Dim yType As Byte = yData(lPos) : lPos += 1

		If yType = 0 Then
			'guild
			If oPlayer.oGuild Is Nothing Then Return
			Dim oVote As GuildVote = oPlayer.oGuild.GetVote(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
			If oVote Is Nothing = False Then
				Dim yResp(7) As Byte
				lPos = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eGetMyVoteValue).CopyTo(yResp, lPos) : lPos += 2
				yResp(lPos) = 0 : lPos += 1
				System.BitConverter.GetBytes(oVote.VoteID).CopyTo(yResp, lPos) : lPos += 4
				yResp(lPos) = oVote.GetPlayerVote(lPlayerID) : lPos += 1
				moClients(lIndex).SendData(yResp)
			End If
		Else
			'senate
		End If

	End Sub
	Private Sub HandleGetNonOwnerDetails(ByRef yData() As Byte, ByVal lIndex As Int32)
		'MscCode (2), GUID (6)
		Dim lObjectID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

		'TODO: Add more here as needed
		Select Case Math.Abs(iObjTypeID)
			Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eRadarTech, ObjectType.ePrototype, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
				'Dim yMsg() As Byte = Nothing
				'If iObjTypeID < 0 Then
				'    For X As Int32 = 0 To glComponentCacheUB
				'        If glComponentCacheIdx(X) = lObjectID Then
				'            Dim oTech As Epica_Tech = Nothing
				'            oTech = goComponentCache(X).GetComponent()
				'            If oTech Is Nothing Then Return
				'            yMsg = oTech.GetNonOwnerMsg()
				'            System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 2)
				'            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 6)
				'            Exit For
				'        End If
				'    Next X
				'Else
				Dim oTech As Epica_Tech = Nothing
                oTech = QuickLookupTechnology(lObjectID, Math.Abs(iObjTypeID))
                If oTech Is Nothing Then
                    Dim oTmpCache As ComponentCache = GetEpicaComponentCache(lObjectID)
                    If oTmpCache Is Nothing Then Return
                    oTech = oTmpCache.GetComponent
                    If oTech Is Nothing Then Return
                    Try
                        Dim yTech() As Byte = oTech.GetNonOwnerMsg
                        System.BitConverter.GetBytes(lObjectID).CopyTo(yTech, 2)
                        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yTech, 6)
                        moClients(lIndex).SendData(yTech)
                    Catch
                    End Try
                    Return
                End If
                If oTech Is Nothing Then Return
                Dim yResp() As Byte = oTech.GetNonOwnerMsg()
                If yResp Is Nothing = False Then
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                    moClients(lIndex).SendData(yResp)
                End If

				'    yMsg = oTech.GetNonOwnerMsg()
				'End If
				'If yMsg Is Nothing = False Then moClients(lIndex).SendData(yMsg)
			Case ObjectType.eAgent
				Dim oAgent As Agent = GetEpicaAgent(lObjectID)
				If oAgent Is Nothing Then Return
				Dim yDef() As Byte = oAgent.GetNonOwnMsg()
				Dim yMsg(yDef.GetUpperBound(0) + 2) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yMsg, 0)
				yDef.CopyTo(yMsg, 2)
				moClients(lIndex).SendData(yMsg)
			Case ObjectType.eUnit
				Dim oUnit As Unit = GetEpicaUnit(lObjectID)
				If oUnit Is Nothing Then Return

				Dim yDef() As Byte = oUnit.EntityDef.GetObjAsString
				Dim yMsg(yDef.GetUpperBound(0) + 2) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yMsg, 0)
				yDef.CopyTo(yMsg, 2)
				System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 2)
				System.BitConverter.GetBytes(oUnit.ObjTypeID).CopyTo(yMsg, 6)
				moClients(lIndex).SendData(yMsg)
			Case ObjectType.eFacility
				Dim oFac As Facility = GetEpicaFacility(lObjectID)
				If oFac Is Nothing Then Return

				Dim yDef() As Byte = oFac.EntityDef.GetObjAsString
				Dim yMsg(yDef.GetUpperBound(0) + 2) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yMsg, 0)
				yDef.CopyTo(yMsg, 2)
				System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 2)
				System.BitConverter.GetBytes(oFac.ObjTypeID).CopyTo(yMsg, 6)
				moClients(lIndex).SendData(yMsg)
			Case ObjectType.eMineralCache
				Dim oCache As MineralCache = GetEpicaMineralCache(lObjectID)
				If oCache Is Nothing = False Then
					If oCache.oMineral Is Nothing = False Then
						Dim yMin() As Byte = oCache.oMineral.GetNonOwnerMsg()
						Dim yMsg(7 + yMin.Length) As Byte

						System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yMsg, 0)
						System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 2)
						System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 6)
						yMin.CopyTo(yMsg, 8)
                        moClients(lIndex).SendData(yMsg)
                    Else
                        Dim oMineral As Mineral = GetEpicaMineral(lObjectID)
                        If oMineral Is Nothing = False Then
                            Dim yMin() As Byte = oMineral.GetNonOwnerMsg()
                            Dim yMsg(7 + yMin.Length) As Byte

                            System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 2)
                            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 6)
                            yMin.CopyTo(yMsg, 8)
                            moClients(lIndex).SendData(yMsg)
                        End If
                    End If
				End If
			Case ObjectType.eMineral
				Dim oMineral As Mineral = GetEpicaMineral(lObjectID)
				If oMineral Is Nothing = False Then
					Dim yMin() As Byte = oMineral.GetNonOwnerMsg()
					Dim yMsg(5 + yMin.Length) As Byte

					System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 2)
					System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 6)
					yMin.CopyTo(yMsg, 8)
					moClients(lIndex).SendData(yMsg)
				End If
			Case ObjectType.eComponentCache
				For X As Int32 = 0 To glComponentCacheUB
					If glComponentCacheIdx(X) = lObjectID Then
						Try
							Dim yTech() As Byte = goComponentCache(X).GetComponent().GetNonOwnerMsg
							Dim yMsg(7 + yTech.Length) As Byte
							System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yMsg, 0)
							System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 2)
							System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 6)
							Array.Copy(yTech, 0, yMsg, 8, yTech.Length)
							moClients(lIndex).SendData(yMsg)
						Catch
						End Try
						Exit For
					End If
				Next X
		End Select
	End Sub
	Private Sub HandleGetOrderSpecifics(ByRef yData() As Byte, ByVal lIndex As Int32)
		If goGTC Is Nothing = False Then
			Dim yMsg() As Byte = goGTC.HandleGetOrderSpecifics(yData)
			If yMsg Is Nothing = False Then
				moClients(lIndex).SendData(yMsg)
			End If
		End If
	End Sub
	Private Sub HandleGetPlayerScores(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

		If lPlayerID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = lPlayerID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewDiplomacy) <> 0) Then
			Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
			If oPlayer Is Nothing = False Then
				For X As Int32 = 0 To oPlayer.mlPlayerIntelUB
					If oPlayer.mlPlayerIntelIdx(X) <> -1 Then
						moClients(lIndex).SendData(GetAddObjectMessage(oPlayer.moPlayerIntel(X), GlobalMessageCode.eAddObjectCommand))
					End If
				Next X

				'Now, send my stuff
				Try
					Dim lCnt As Int32 = 0
					For X As Int32 = 0 To oPlayer.mlColonyUB
						If oPlayer.mlColonyIdx(X) <> -1 AndAlso glColonyIdx(oPlayer.mlColonyIdx(X)) = oPlayer.mlColonyID(X) AndAlso CType(goColony(oPlayer.mlColonyIdx(X)).ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then lCnt += 1
					Next X

                    With oPlayer
                        'If .lLastGlobalRequestTurnIn = 0 OrElse glCurrentCycle - .lLastGlobalRequestTurnIn > 9000 Then
                        'goMsgSys.SendRequestGlobalPlayerScores(.ObjectID, 64, .ObjectID)
                        'If .lLastGlobalRequestTurnIn = 0 Then
                        .lLGTechScore = .TechnologyScore
                        .lLGDiplomacyScore = .DiplomacyScore
                        .lLGWealthScore = .WealthScore
                        .lLGProductionScore = .ProductionScore
                        .lLGPopulationScore = .PopulationScore
                        .lLGMilitaryScore = .lMilitaryScore
                        .lLGTotalScore = .TotalScore
                        'End If
                        'End If
                    End With

                    Dim yMsg(31 + (lCnt * 5)) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eGetPlayerScores).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(oPlayer.lLGTechScore).CopyTo(yMsg, 2)
                    System.BitConverter.GetBytes(oPlayer.lLGDiplomacyScore).CopyTo(yMsg, 6)
                    System.BitConverter.GetBytes(CInt(oPlayer.lLGMilitaryScore \ 50)).CopyTo(yMsg, 10)      'was 100
                    System.BitConverter.GetBytes(oPlayer.lLGPopulationScore).CopyTo(yMsg, 14)
                    System.BitConverter.GetBytes(oPlayer.lLGProductionScore).CopyTo(yMsg, 18)
                    System.BitConverter.GetBytes(oPlayer.lLGWealthScore).CopyTo(yMsg, 22)
                    System.BitConverter.GetBytes(oPlayer.lLGTotalScore).CopyTo(yMsg, 26)
                    System.BitConverter.GetBytes(CShort(lCnt)).CopyTo(yMsg, 30)
                    Dim lPos As Int32 = 32
                    For X As Int32 = 0 To oPlayer.mlColonyUB
                        If oPlayer.mlColonyIdx(X) <> -1 AndAlso glColonyIdx(oPlayer.mlColonyIdx(X)) = oPlayer.mlColonyID(X) AndAlso CType(goColony(oPlayer.mlColonyIdx(X)).ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                            System.BitConverter.GetBytes(CType(goColony(oPlayer.mlColonyIdx(X)).ParentObject, Epica_GUID).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                            yMsg(lPos) = goColony(oPlayer.mlColonyIdx(X)).GovScore : lPos += 1
                        End If
                    Next X

                    moClients(lIndex).SendData(yMsg)
                Catch
                    'do nothing
                End Try
            End If
        End If
	End Sub
	Private Sub HandleGetPMUpdate(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2	'for msgcode
		Dim lPM_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) = 0 Then Return
		End If

		Dim oPM As PlayerMission = GetEpicaPlayerMission(lPM_ID)
		If oPM Is Nothing Then Return
		If oPM.oPlayer.ObjectID = lPlayerID Then
			'now, return our msg
			moClients(lIndex).SendData(oPM.GetPMUpdateMsg())
		Else : LogEvent(LogEventType.PossibleCheat, "HandleGetPMUpdate: Player mismatch. Player: " & mlClientPlayer(lIndex))
		End If
	End Sub
	Private Sub HandleGetRouteList(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2

		Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oUnit As Unit = GetEpicaUnit(lID)
		If oUnit Is Nothing Then Return

		If oUnit.Owner.ObjectID = mlClientPlayer(lIndex) OrElse (oUnit.Owner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewUnitsAndFacilities) <> 0) Then

            Dim yResp(10 + ((oUnit.lRouteUB + 1) * 27)) As Byte
			lPos = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eGetRouteList).CopyTo(yResp, lPos) : lPos += 2
			System.BitConverter.GetBytes(oUnit.ObjectID).CopyTo(yResp, lPos) : lPos += 4
			If oUnit.bRoutePaused = True Then yResp(lPos) = 1 Else yResp(lPos) = 0
			lPos += 1
			System.BitConverter.GetBytes(oUnit.lRouteUB + 1).CopyTo(yResp, lPos) : lPos += 4
			For X As Int32 = 0 To oUnit.lRouteUB
				With oUnit.uRoute(X)
					If .oDest Is Nothing = False Then
						.oDest.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
						If (.oDest.ObjTypeID = ObjectType.eUnit OrElse .oDest.ObjTypeID = ObjectType.eFacility) AndAlso CType(.oDest, Epica_Entity).ParentObject Is Nothing = False Then
							CType(CType(.oDest, Epica_Entity).ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
						Else
							System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
							System.BitConverter.GetBytes(-1S).CopyTo(yResp, lPos) : lPos += 2
						End If
					Else
						System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(-1S).CopyTo(yResp, lPos) : lPos += 2
						System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(-1S).CopyTo(yResp, lPos) : lPos += 2
					End If
					If .oLoadItem Is Nothing = False Then
						.oLoadItem.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
					Else
                        System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(-1S).CopyTo(yResp, lPos) : lPos += 2
					End If
					System.BitConverter.GetBytes(.lLocX).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lLocZ).CopyTo(yResp, lPos) : lPos += 4

                    yResp(lPos) = .yExtraFlags : lPos += 1
				End With
			Next X

			moClients(lIndex).SendData(yResp)
		Else : LogEvent(LogEventType.PossibleCheat, "HandleGetRouteList: owner mismatch. PlayerID: " & mlClientPlayer(lIndex))
		End If
    End Sub
    Private Sub HandleGetRouteTemplates(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eMoveUnits Or AliasingRights.eViewUnitsAndFacilities)) <> (AliasingRights.eMoveUnits Or AliasingRights.eViewUnitsAndFacilities) Then
                Return
            Else
                lPlayerID = mlAliasedAs(lIndex)
            End If
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim yResp() As Byte = oPlayer.HandleGetRouteTemplateList
        If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
    End Sub
	Private Sub HandleGetSkillList(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, 2)

		Dim oAgent As Agent = GetEpicaAgent(lAgentID)
		If oAgent Is Nothing = False AndAlso oAgent.oOwner Is Nothing = False Then
			If oAgent.oOwner.ObjectID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = oAgent.oOwner.ObjectID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) <> 0) Then

				Dim yResp(9 + ((oAgent.SkillUB + 1) * 5)) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yResp, lPos) : lPos += 2

				With oAgent
					System.BitConverter.GetBytes(.ObjectID).CopyTo(yResp, lPos) : lPos += 4
					System.BitConverter.GetBytes(.SkillUB + 1).CopyTo(yResp, lPos) : lPos += 4

					For X As Int32 = 0 To .SkillUB
						System.BitConverter.GetBytes(.Skills(X).ObjectID).CopyTo(yResp, lPos) : lPos += 4
						yResp(lPos) = .SkillProf(X) : lPos += 1
					Next X
				End With

				moClients(lIndex).SendData(yResp)
			End If
		End If
    End Sub
    Private Sub HandleGetSpecialTechGuaranteeList(ByVal yData() As Byte, ByVal lIndex As Int32)

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.GuaranteedSpecialTechID > 0 Then Return 'oPlayer.yPlayerPhase <> eyPlayerPhase.eSecondPhase Then Return

        Dim yMsg(gyGuaranteedList.GetUpperBound(0)) As Byte

        gyGuaranteedList.CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(oPlayer.GuaranteedSpecialTechID).CopyTo(yMsg, 6)

        moClients(lIndex).SendData(yMsg)
    End Sub
	Private Sub HandleGetTradeDeliveries(ByRef yData() As Byte, ByVal lIndex As Int32)
		If goGTC Is Nothing Then Return
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
		If mlClientPlayer(lIndex) <> lPlayerID AndAlso (mlAliasedAs(lIndex) <> lPlayerID OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewTrades) = 0) Then
			LogEvent(LogEventType.PossibleCheat, "HandleGetTradeDeliveries, PlayerID requested does not equal socket's player ID. SocketPlayerID: " & mlClientPlayer(lIndex))
		Else
			Dim yResp() As Byte = goGTC.HandleGetTradeDeliveries(lPlayerID)
			If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
            goGTC.SendPlayerDirectTrades(lPlayerID, moClients(lIndex))
		End If
	End Sub
	Private Sub HandleGetTradeHistory(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

		If mlClientPlayer(lIndex) = lPlayerID OrElse (mlAliasedAs(lIndex) = lPlayerID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewTrades) <> 0) Then
			Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing = False Then

                If oPlayer.InMyDomain = False Then
                    SendPassThruMsg(yData, lPlayerID, lPlayerID, ObjectType.ePlayer)
                    Return
                End If

                moClients(lIndex).SendData(oPlayer.GetTradeHistoryMsg())
            End If
        Else
            LogEvent(LogEventType.PossibleCheat, "HandleGetTradeHistory: Player ID Mismatch: " & mlClientPlayer(lIndex))
        End If
	End Sub
	Private Sub HandleGetTradePostList(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eViewTrades) = 0 Then Return
		End If

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		Dim lColonyIdx() As Int32 = Nothing
		Dim lChildIdx() As Int32 = Nothing
		Dim lUB As Int32 = -1

		For X As Int32 = 0 To oPlayer.mlColonyUB
			If oPlayer.mlColonyIdx(X) <> -1 AndAlso glColonyIdx(oPlayer.mlColonyIdx(X)) = oPlayer.mlColonyID(X) Then
				With goColony(oPlayer.mlColonyIdx(X))
					For Y As Int32 = 0 To .ChildrenUB
						If .lChildrenIdx(Y) <> -1 AndAlso .oChildren(Y).yProductionType = ProductionType.eTradePost Then
							lUB += 1
							ReDim Preserve lColonyIdx(lUB)
							ReDim Preserve lChildIdx(lUB)
							lColonyIdx(lUB) = oPlayer.mlColonyIdx(X)
                            lChildIdx(lUB) = Y
                            Exit For
						End If
					Next Y
				End With
			End If
		Next X

		'Now, create our list...
        Dim yResp(3 + ((lUB + 1) * 44)) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostList).CopyTo(yResp, lPos) : lPos += 2
		System.BitConverter.GetBytes(CShort(lUB + 1)).CopyTo(yResp, lPos) : lPos += 2
		For X As Int32 = 0 To lUB
			With goColony(lColonyIdx(X))
				System.BitConverter.GetBytes(.oChildren(lChildIdx(X)).ObjectID).CopyTo(yResp, lPos) : lPos += 4

                '.ColonyName.CopyTo(yResp, lPos) : lPos += 20
                Dim oTmpGuid As Epica_GUID = CType(.oChildren(lChildIdx(X)).ParentObject, Epica_GUID)
                If oTmpGuid Is Nothing = False Then
                    If oTmpGuid.ObjTypeID = ObjectType.ePlanet Then
                        CType(.oChildren(lChildIdx(X)).ParentObject, Planet).PlanetName.CopyTo(yResp, lPos) : lPos += 20
                    ElseIf oTmpGuid.ObjTypeID = ObjectType.eSolarSystem Then
                        CType(.oChildren(lChildIdx(X)).ParentObject, SolarSystem).SystemName.CopyTo(yResp, lPos) : lPos += 20
                    ElseIf oTmpGuid.ObjTypeID = ObjectType.eFacility Then
                        oTmpGuid = CType(CType(.oChildren(lChildIdx(X)).ParentObject, Facility).ParentObject, Epica_GUID)
                        If oTmpGuid.ObjTypeID = ObjectType.eSolarSystem Then
                            CType(oTmpGuid, SolarSystem).SystemName.CopyTo(yResp, lPos) : lPos += 20
                        Else
                            StringToBytes("Unknown").CopyTo(yResp, lPos) : lPos += 20
                        End If
                    Else
                        StringToBytes("Unknown").CopyTo(yResp, lPos) : lPos += 20
                    End If
                End If
                .oChildren(lChildIdx(X)).EntityName.CopyTo(yResp, lPos) : lPos += 20
			End With
		Next X

		moClients(lIndex).SendData(yResp)
    End Sub
    Private Structure uTradeMineralCache
        Public MineralCacheID As Int32
        Public MineralID As Int32
        Public Quantity As Int32
    End Structure
    Private Structure uTradeComponentCache
        Public ComponentCacheID As Int32
        Public ComponentTypeID As Int16
        Public ComponentID As Int32
        Public ComponentOwnerID As Int32
        Public Quantity As Int32
    End Structure
    Private Sub HandleGetTradePostTradeables(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lTradePostID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim oTP As Facility = GetEpicaFacility(lTradePostID)

        If oTP Is Nothing = False Then
            Dim oColony As Colony = oTP.ParentColony
            If oColony Is Nothing = True Then Return
            If oTP.Owner.ObjectID = mlClientPlayer(lIndex) OrElse (oTP.Owner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewTrades) <> 0) Then
                Dim lComponents As Int32 = 0
                Dim lMinerals As Int32 = 0
                Dim lUnits As Int32 = 0
                Dim lIntels As Int32 = goGTC.GetNumberOfIntelItemsBeingSold(lTradePostID)

                Dim uMinerals(-1) As uTradeMineralCache
                Dim uComponents(-1) As uTradeComponentCache

                For X As Int32 = 0 To oColony.mlMineralCacheUB
                    If oColony.mlMineralCacheIdx(X) > -1 Then
                        If glMineralCacheIdx(oColony.mlMineralCacheIdx(X)) = oColony.mlMineralCacheID(X) Then
                            Dim oCache As MineralCache = goMineralCache(oColony.mlMineralCacheIdx(X))
                            If oCache Is Nothing = False And oCache.Quantity > 0 Then
                                lMinerals += 1
                                ReDim Preserve uMinerals(lMinerals - 1)
                                With uMinerals(lMinerals - 1)
                                    .MineralCacheID = oCache.ObjectID
                                    .MineralID = oCache.oMineral.ObjectID
                                    .Quantity = oCache.Quantity
                                End With
                            End If
                        End If
                    End If
                Next X
                For X As Int32 = 0 To oColony.mlComponentCacheUB
                    If oColony.mlComponentCacheIdx(X) > -1 Then
                        If glComponentCacheIdx(oColony.mlComponentCacheIdx(X)) = oColony.mlComponentCacheID(X) Then
                            Dim oCache As ComponentCache = goComponentCache(oColony.mlComponentCacheIdx(X))
                            If oCache Is Nothing = False AndAlso oCache.Quantity > 0 Then
                                Dim bFound As Boolean = False
                                For Z As Int32 = 0 To lComponents - 1
                                    If uComponents(Z).ComponentID = oCache.ComponentID AndAlso uComponents(Z).ComponentTypeID = oCache.ComponentTypeID AndAlso uComponents(Z).ComponentOwnerID = oCache.ComponentOwnerID Then
                                        bFound = True
                                        uComponents(Z).Quantity += oCache.Quantity
                                        Exit For
                                    End If
                                Next Z
                                If bFound = False Then
                                    lComponents += 1
                                    ReDim Preserve uComponents(lComponents - 1)
                                    With uComponents(lComponents - 1)
                                        .ComponentCacheID = oCache.ObjectID
                                        .ComponentTypeID = oCache.ComponentTypeID
                                        .ComponentID = oCache.ComponentID
                                        .ComponentOwnerID = oCache.ComponentOwnerID
                                        .Quantity = oCache.Quantity
                                    End With
                                End If
                            End If
                        End If
                    End If
                Next X

                'components and minerals
                For X As Int32 = 0 To oColony.ChildrenUB
                    If oColony.lChildrenIdx(X) > -1 Then
                        Dim oFac As Facility = oColony.oChildren(X)
                        If oFac Is Nothing = False AndAlso oFac.Active = True Then
                            'If (oFac.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                            '    For Y As Int32 = 0 To oFac.lCargoUB
                            '        If oFac.lCargoIdx(Y) <> -1 Then
                            '            If oFac.oCargoContents(Y).ObjTypeID = ObjectType.eComponentCache Then
                            '                Dim oCache As ComponentCache = CType(oFac.oCargoContents(Y), ComponentCache)
                            '                If oCache Is Nothing Then Continue For
                            '                Dim bFound As Boolean = False
                            '                For Z As Int32 = 0 To lComponents - 1
                            '                    If uComponents(Z).ComponentID = oCache.ComponentID AndAlso uComponents(Z).ComponentTypeID = oCache.ComponentTypeID AndAlso uComponents(Z).ComponentOwnerID = oCache.ComponentOwnerID Then
                            '                        bFound = True
                            '                        uComponents(Z).Quantity += oCache.Quantity
                            '                        Exit For
                            '                    End If
                            '                Next Z
                            '                If bFound = False Then
                            '                    lComponents += 1
                            '                    ReDim Preserve uComponents(lComponents - 1)
                            '                    With uComponents(lComponents - 1)
                            '                        .ComponentCacheID = oCache.ObjectID
                            '                        .ComponentTypeID = oCache.ComponentTypeID
                            '                        .ComponentID = oCache.ComponentID
                            '                        .ComponentOwnerID = oCache.ComponentOwnerID
                            '                        .Quantity = oCache.Quantity
                            '                    End With
                            '                End If
                            '            End If
                            '        End If
                            '    Next Y
                            'End If
                            If (oFac.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                                For Y As Int32 = 0 To oFac.lHangarUB
                                    If oFac.lHangarIdx(Y) <> -1 Then
                                        If oFac.oHangarContents(Y).ObjTypeID = ObjectType.eUnit Then lUnits += 1
                                    End If
                                Next Y
                            End If
                        End If
                    End If
                Next X

                Dim yResp(19 + (lComponents * 30) + (lMinerals * 26) + (lUnits * 19) + (lIntels * 16)) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostTradeables).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(lTradePostID).CopyTo(yResp, lPos) : lPos += 4
                yResp(lPos) = CByte(oTP.lTradePostSellSlotsUsed) : lPos += 1
                yResp(lPos) = CByte(oTP.lTradePostBuySlotsUsed) : lPos += 1

                With oTP.ParentColony
                    System.BitConverter.GetBytes(.Population).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.ColonyEnlisted).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.ColonyOfficers).CopyTo(yResp, lPos) : lPos += 4
                End With

                For X As Int32 = 0 To lMinerals - 1
                    With uMinerals(X)
                        System.BitConverter.GetBytes(.MineralCacheID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(ObjectType.eMineralCache).CopyTo(yResp, lPos) : lPos += 2
                        System.BitConverter.GetBytes(.MineralID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4
                        Dim oSO As SellOrder = goGTC.GetSellOrderItem(oTP.ObjectID, .MineralCacheID, ObjectType.eMineralCache)
                        If oSO Is Nothing = False Then
                            System.BitConverter.GetBytes(CInt(oSO.blQuantity)).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(oSO.blPrice).CopyTo(yResp, lPos) : lPos += 8
                        Else
                            System.BitConverter.GetBytes(0I).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(0L).CopyTo(yResp, lPos) : lPos += 8
                        End If
                    End With 
                Next X

                For X As Int32 = 0 To lComponents - 1
                    With uComponents(X)
                        System.BitConverter.GetBytes(.ComponentCacheID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(CShort(.ComponentTypeID * -1S)).CopyTo(yResp, lPos) : lPos += 2
                        System.BitConverter.GetBytes(.ComponentID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.ComponentOwnerID).CopyTo(yResp, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4

                        Dim oSO As SellOrder = goGTC.GetSellOrderItem(oTP.ObjectID, .ComponentID, CShort(.ComponentTypeID * -1S))
                        If oSO Is Nothing = False Then
                            System.BitConverter.GetBytes(CInt(oSO.blQuantity)).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(oSO.blPrice).CopyTo(yResp, lPos) : lPos += 8
                        Else
                            System.BitConverter.GetBytes(0I).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(0L).CopyTo(yResp, lPos) : lPos += 8
                        End If
                    End With
                Next X

                'units 
                For lFac As Int32 = 0 To oColony.ChildrenUB
                    If oColony.lChildrenIdx(lFac) > -1 Then
                        Dim oFac As Facility = oColony.oChildren(lFac)
                        If oFac Is Nothing = False AndAlso oFac.Active = True Then
                            If (oFac.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 AndAlso oFac.yProductionType = ProductionType.eTradePost Then
                                For Y As Int32 = 0 To oFac.lHangarUB
                                    If oFac.lHangarIdx(Y) <> -1 Then
                                        If oFac.oHangarContents(Y).ObjTypeID = ObjectType.eUnit Then
                                            With CType(oFac.oHangarContents(Y), Unit)
                                                .GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6

                                                Dim oSO As SellOrder = goGTC.GetSellOrderItem(oTP.ObjectID, .ObjectID, .ObjTypeID)
                                                If oSO Is Nothing = False Then
                                                    System.BitConverter.GetBytes(CInt(oSO.blQuantity)).CopyTo(yResp, lPos) : lPos += 4
                                                    System.BitConverter.GetBytes(oSO.blPrice).CopyTo(yResp, lPos) : lPos += 8
                                                Else
                                                    System.BitConverter.GetBytes(0I).CopyTo(yResp, lPos) : lPos += 4
                                                    System.BitConverter.GetBytes(0L).CopyTo(yResp, lPos) : lPos += 8
                                                End If

                                                yResp(lPos) = .EntityDef.yChassisType : lPos += 1
                                            End With
                                        End If
                                    End If
                                Next Y
                                Exit For
                            End If
                        End If
                    End If
                Next lFac 

                goGTC.FillWithIntelItemsBeingsold(yResp, lPos, lTradePostID)

                moClients(lIndex).SendData(yResp)

            Else
                LogEvent(LogEventType.PossibleCheat, "GetTradePostTradeables Owners do not match")
            End If
        End If
    End Sub
	Private Sub HandleGuildBankTrans(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return
        If oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return
        If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount AndAlso oPlayer.AccountStatus <> AccountStatusType.eMondelisActive AndAlso gb_IN_OPEN_BETA = False Then Return

        If oPlayer.lLastAvailableResourcesRequest = Int32.MinValue OrElse glCurrentCycle - oPlayer.lLastAvailableResourcesRequest > 30 Then
            oPlayer.lLastAvailableResourcesRequest = glCurrentCycle
        Else : Return
        End If

		Dim lPos As Int32 = 2 'for msgcode
		Dim yTransType As Byte = yData(lPos) : lPos += 1
		Dim bHighSec As Boolean = (yTransType And 128) <> 0
		If bHighSec = True Then yTransType = CByte(yTransType Xor 128)

		Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim blQty As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

		'check our permissions
		If yTransType = 2 Then
			'withdrawing... check our rights
			If bHighSec = True Then
				If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ViewContentsHiSec) = False OrElse _
				 oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.WithdrawHiSec) = False Then
					'cheater
					LogEvent(LogEventType.PossibleCheat, "Player withdrawing from bank without permission: " & mlClientPlayer(lIndex))
					Return
				End If
			Else
				If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ViewContentsLowSec) = False OrElse _
				  oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.WithdrawLowSec) = False Then
					'cheater
					LogEvent(LogEventType.PossibleCheat, "Player withdrawing from bank without permission: " & mlClientPlayer(lIndex))
					Return
				End If
			End If
		End If

        Dim oGuild As Guild = oPlayer.oGuild
        If oGuild Is Nothing Then Return
        If yTransType = 0 Then
            'deposit
            Dim blActual As Int64 = Math.Min(blQty, oPlayer.blCredits)
            If blActual < 1 Then Return
            oPlayer.blCredits -= blActual
            oGuild.blTreasury += blActual
            'LogEvent(LogEventType.ExtensiveLogging, "GuildDeposit: " & mlClientPlayer(lIndex) & " deposited " & blActual.ToString)
            Dim oEvent As GuildTransLog = GuildTransLog.CreateGuildTransLog(oGuild.ObjectID, oPlayer.ObjectID, blActual, oGuild.blTreasury, 0)
            If oEvent Is Nothing = False Then oGuild.AddBankLogItem(oEvent)
        Else
            'withdraw
            Dim blActual As Int64 = Math.Min(blQty, oGuild.blTreasury)
            If blActual < 1 Then Return
            oGuild.blTreasury -= blActual
            oPlayer.blCredits += blActual
            'LogEvent(LogEventType.ExtensiveLogging, "GuildWithdraw: " & mlClientPlayer(lIndex) & " withdrew " & blActual.ToString)
            Dim oEvent As GuildTransLog = GuildTransLog.CreateGuildTransLog(oGuild.ObjectID, oPlayer.ObjectID, blActual, oGuild.blTreasury, 1)
            If oEvent Is Nothing = False Then oGuild.AddBankLogItem(oEvent)
        End If
        oGuild.SendGuildTreasuryUpdate()

	End Sub
	Private Sub HandleGuildCreationAcceptance(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

		Dim lPos As Int32 = 2	'for msgcode
		Dim lGuildFormer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lTestPlayer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If lTestPlayer <> lPlayerID Then
			LogEvent(LogEventType.PossibleCheat, "HandleGuildCreationAcceptance player mismatch: " & lPlayerID)
			Return
		End If
		Dim yValue As Byte = yData(lPos) : lPos += 1

		Dim oFormer As Player = GetEpicaPlayer(lGuildFormer)
		If oFormer Is Nothing = False Then
			'ok, find member object
			If oFormer.oFormingGuild Is Nothing = False Then
				With oFormer.oFormingGuild
					If .oMembers Is Nothing = False Then
						For X As Int32 = 0 To .oMembers.GetUpperBound(0)
							If .oMembers(X) Is Nothing = False AndAlso .oMembers(X).lMemberID = lPlayerID Then
								If yValue = 0 Then
									If (.oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
										.oMembers(X).yMemberState = .oMembers(X).yMemberState Xor GuildMemberState.Approved
									End If
								ElseIf yValue = 1 Then
									.oMembers(X).yMemberState = .oMembers(X).yMemberState Or GuildMemberState.Approved
								ElseIf yValue = 255 Then
									.SendMsgToGuildMembers(yData)
									.oMembers(X) = Nothing
									Return
								End If
								Exit For
							End If
						Next X
						.SendMsgToGuildMembers(yData)
					End If
				End With
			End If
		End If

	End Sub
	Private Sub HandleGuildCreationUpdate(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

		'only the player creating the guild sends this message
		If oPlayer.oFormingGuild Is Nothing Then oPlayer.oFormingGuild = New Guild
		With oPlayer.oFormingGuild
			Dim lPos As Int32 = 2	'for msgcode

			Dim lCreator As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If lCreator <> lPlayerID Then
				LogEvent(LogEventType.PossibleCheat, "Creator does not match player in guild creation: " & lPlayerID)
				oPlayer.oFormingGuild = Nothing
				Return
			End If

			.lBaseGuildFlags = CType(System.BitConverter.ToInt32(yData, lPos), elGuildFlags) : lPos += 4
			.lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			ReDim .yGuildName(49)
			Array.Copy(yData, lPos, .yGuildName, 0, 50) : lPos += 50
			.yGuildTaxBaseMonth = yData(lPos) : lPos += 1
			.yGuildTaxBaseDay = yData(lPos) : lPos += 1
			.yGuildTaxRateInterval = CType(yData(lPos), eyGuildInterval) : lPos += 1
			.yState = CType(yData(lPos), eyGuildState) : lPos += 1
			.yVoteWeightType = CType(yData(lPos), eyVoteWeightType) : lPos += 1

			Dim lMemberCnt As Int32 = yData(lPos) : lPos += 1
			ReDim .oMembers(lMemberCnt - 1)
			.lMemberUB = lMemberCnt - 1
			For X As Int32 = 0 To lMemberCnt - 1
				.oMembers(X) = New GuildMember

				.oMembers(X).lMemberID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.oMembers(X).lCreateRankID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.oMembers(X).yMemberState = CType(yData(lPos), GuildMemberState) : lPos += 1
				If .oMembers(X).lMemberID <> lPlayerID Then
					If (.oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
						.oMembers(X).yMemberState = .oMembers(X).yMemberState Xor GuildMemberState.Approved
					End If
				End If
			Next X

			Dim lRankCnt As Int32 = yData(lPos) : lPos += 1
			ReDim .oRanks(lRankCnt - 1)
			.lRankUB = lRankCnt - 1
			For X As Int32 = 0 To lRankCnt - 1
				.oRanks(X) = New GuildRank()
				.oRanks(X).lRankID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.oRanks(X).lRankPermissions = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.oRanks(X).lVoteStrength = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				ReDim .oRanks(X).yRankName(19)
				Array.Copy(yData, lPos, .oRanks(X).yRankName, 0, 20) : lPos += 20
				.oRanks(X).TaxRateFlat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.oRanks(X).TaxRatePercentage = yData(lPos) : lPos += 1
				.oRanks(X).TaxRatePercType = CType(yData(lPos), eyGuildTaxPercType) : lPos += 1
				.oRanks(X).yPosition = yData(lPos) : lPos += 1
			Next X

			If .oMembers Is Nothing = False Then
				For X As Int32 = 0 To .oMembers.GetUpperBound(0)
					If .oMembers(X) Is Nothing = False AndAlso .oMembers(X).lMemberID <> oPlayer.ObjectID Then
						If (.oMembers(X).yMemberState And GuildMemberState.AcceptedGuildFormInvite) <> 0 Then
							If .oMembers(X).oMember Is Nothing = False Then
								.oMembers(X).oMember.SendPlayerMessage(yData, False, 0)
							End If
						End If
					End If
				Next X
			End If
		End With
	End Sub
	Private Sub HandleGuildMemberStatus(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim lPos As Int32 = 2	 'for msgcode
		Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lMemberID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lStatusUpdate As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oActingPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oActingPlayer Is Nothing Then Return

		Select Case lStatusUpdate
			Case -1
				'demote the specified member
				If oActingPlayer.oGuild Is Nothing = False Then
					If oActingPlayer.oGuild.RankHasPermission(oActingPlayer.lGuildRankID, RankPermissions.DemoteMember) = False Then
						LogEvent(LogEventType.PossibleCheat, "Player attempting to demote member without permission: " & lPlayerID)
						Return
					End If
					Dim oMyRank As GuildRank = oActingPlayer.oGuild.GetRank(oActingPlayer.lGuildRankID)
					If oMyRank Is Nothing Then Return
					Dim oTarget As Player = GetEpicaPlayer(lMemberID)
					If oTarget Is Nothing = False AndAlso oTarget.oGuild Is Nothing = False AndAlso oTarget.oGuild.ObjectID = oActingPlayer.oGuild.ObjectID Then

						Dim oCurrRank As GuildRank = oActingPlayer.oGuild.GetRank(oTarget.lGuildRankID)
						If oCurrRank Is Nothing Then Return
						If oCurrRank.yPosition < oMyRank.yPosition Then
							LogEvent(LogEventType.PossibleCheat, "Player attempting to demote player higher ranked: " & lPlayerID)
							Return
						End If

						Dim oNextRank As GuildRank = oActingPlayer.oGuild.GetNextRankPosition(1, oTarget.lGuildRankID)
						If oNextRank Is Nothing = False Then
							oTarget.lGuildRankID = oNextRank.lRankID
							oActingPlayer.oGuild.SendMsgToGuildMembers(yData)
						End If
					Else : Return
					End If
				End If
			Case -2
				'promote the specified member
				If oActingPlayer.oGuild Is Nothing = False Then
					If oActingPlayer.oGuild.RankHasPermission(oActingPlayer.lGuildRankID, RankPermissions.PromoteMember) = False Then
						LogEvent(LogEventType.PossibleCheat, "Player attempting to promote member without permission: " & lPlayerID)
						Return
					End If
					Dim oMyRank As GuildRank = oActingPlayer.oGuild.GetRank(oActingPlayer.lGuildRankID)
					If oMyRank Is Nothing Then Return
					Dim oTarget As Player = GetEpicaPlayer(lMemberID)
					If oTarget Is Nothing = False AndAlso oTarget.oGuild Is Nothing = False AndAlso oTarget.oGuild.ObjectID = oActingPlayer.oGuild.ObjectID Then

						Dim oCurrRank As GuildRank = oActingPlayer.oGuild.GetRank(oTarget.lGuildRankID)
						If oCurrRank Is Nothing Then Return
						If oCurrRank.yPosition < oMyRank.yPosition Then
							LogEvent(LogEventType.PossibleCheat, "Player attempting to promote player higher ranked: " & lPlayerID)
							Return
						End If

						Dim oNextRank As GuildRank = oActingPlayer.oGuild.GetNextRankPosition(-1, oTarget.lGuildRankID)
						If oNextRank Is Nothing = False Then
							oTarget.lGuildRankID = oNextRank.lRankID
							oActingPlayer.oGuild.SendMsgToGuildMembers(yData)
						End If
					Else : Return
					End If
				End If
			Case Int32.MinValue
				'remove the specified player or reject the player if they are applying
				If oActingPlayer.oGuild Is Nothing = False Then

					'Ok, determine what we are doing... if the target is not a member of the guild, we want the "Reject" member
					' otherwise, we want to "Remove" the member
					Dim oMember As GuildMember = oActingPlayer.oGuild.GetMember(lMemberID)
					If oMember Is Nothing = False Then
						If (oMember.yMemberState And GuildMemberState.Approved) <> 0 Then
							'ok, member is already a member...
							Dim oGuild As Guild = oActingPlayer.oGuild
							If oGuild.RankHasPermission(oActingPlayer.lGuildRankID, RankPermissions.RemoveMember) = False AndAlso oActingPlayer.ObjectID <> lMemberID Then
								LogEvent(LogEventType.PossibleCheat, "Player attempting to remove member from guild without permission: " & mlClientPlayer(lIndex))
								Return
							End If

							'ok, we have the rights... do it
							If (oMember.oMember.iInternalEmailSettings And eEmailSettings.eGuildMembershipNotices) <> 0 Then
								Dim oSB As New System.Text.StringBuilder()
								oSB.AppendLine("You have been removed from " & BytesToString(oGuild.yGuildName) & " by " & oActingPlayer.sPlayerNameProper & ".")
                                Dim oPC As PlayerComm = oMember.oMember.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Guild Removal", oActingPlayer.ObjectID, GetDateAsNumber(Now), False, oMember.oMember.sPlayerNameProper, Nothing)
								If oPC Is Nothing = False Then
                                    oMember.oMember.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
								End If
								oPC = Nothing
							End If

                            oMember.oMember.SendPlayerMessage(yData, False, 0)
                            oMember.oMember.dtLastGuildMembership = Now

							oGuild.RemoveMember(lMemberID)

							'Now, remove the member
							oMember.oMember.lGuildID = -1
							oMember.oMember.lGuildRankID = -1
							oGuild.SendMsgToGuildMembers(yData)
 
						Else
							'ok, member is applied or invited...
							Dim oGuild As Guild = oActingPlayer.oGuild

							If oGuild.RankHasPermission(oActingPlayer.lGuildRankID, RankPermissions.RejectMember) = False AndAlso oActingPlayer.ObjectID <> oMember.lMemberID Then
								LogEvent(LogEventType.PossibleCheat, "Player attempting to reject member without permission: " & mlClientPlayer(lIndex))
								Return
							End If

							If (oMember.yMemberState And GuildMemberState.Applied) <> 0 Then
                                Dim oPC As PlayerComm = oMember.oMember.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                                 "Your application to join " & BytesToString(oActingPlayer.oGuild.yGuildName) & " has been rejected.", "Application Rejected", oActingPlayer.ObjectID, GetDateAsNumber(Now), False, oMember.oMember.sPlayerNameProper, Nothing)
								If oPC Is Nothing = False Then
									oMember.oMember.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
								End If
								oPC = Nothing

                                oGuild.RemoveMember(lMemberID)
                                If oMember.oMember.lGuildID = oGuild.ObjectID Then
                                    oMember.oMember.lGuildID = -1
                                    oMember.oMember.lGuildRankID = -1
                                End If
                                oGuild.SendMsgToGuildMembers(yData)
 

							ElseIf (oMember.yMemberState And GuildMemberState.Invited) <> 0 Then
								oGuild.RemoveMember(lMemberID)
								For X As Int32 = 0 To oMember.oMember.lInvitedToJoinUB
									If oMember.oMember.lInvitedToJoins(X) = oGuild.ObjectID Then
										oMember.oMember.lInvitedToJoins(X) = -1
									End If
                                Next X
                                If oMember.oMember.lGuildID = oGuild.ObjectID Then
                                    oMember.oMember.lGuildID = -1
                                    oMember.oMember.lGuildRankID = -1
                                End If
								oGuild.SendMsgToGuildMembers(yData)
							End If
						End If
						oMember.oMember.SendPlayerMessage(yData, False, 0)
						
					End If
				End If
			Case GuildMemberState.Approved
				Dim oGuild As Guild = GetEpicaGuild(lGuildID)
                If oGuild Is Nothing Then Return

                Dim oTmpPlayer As Player = GetEpicaPlayer(lMemberID)
                If oTmpPlayer Is Nothing = False AndAlso (oTmpPlayer.dtLastGuildMembership <> Date.MinValue AndAlso oTmpPlayer.dtLastGuildMembership.AddDays(3) > Now) Then
                    Dim tsTemp As TimeSpan = oTmpPlayer.dtLastGuildMembership.AddDays(3).Subtract(Now)
                    Dim sMsg As String = "You cannot attempt to join a guild for another "
                    If tsTemp.Days > 0 Then sMsg &= tsTemp.Days.ToString & " days "
                    If tsTemp.Hours > 0 Then sMsg &= tsTemp.Hours.ToString & " hours "
                    If tsTemp.Minutes > 0 Then sMsg &= tsTemp.Minutes.ToString & " minutes "
                    Dim yResp() As Byte = CreateChatMsg(oTmpPlayer.ObjectID, sMsg, ChatMessageType.eSysAdminMessage)
                    oTmpPlayer.SendPlayerMessage(yResp, False, 0)
                    Return
                End If

				If oActingPlayer.ObjectID = lMemberID Then
					'if acting is the target and I am invited, 
					Dim oMember As GuildMember = oGuild.GetMember(lMemberID)
					If oMember Is Nothing Then Return
					If (oMember.yMemberState And GuildMemberState.Invited) <> 0 Then
						'check if I am already in a guild... if not, I approve the invite and join the guild
						If oMember.oMember.oGuild Is Nothing = False Then
							Return
						Else
							oMember.yMemberState = oMember.yMemberState Or GuildMemberState.Approved
							oMember.oMember.lGuildID = oGuild.ObjectID
                            oGuild.MemberJoined(lMemberID, False)
                            oMember.oMember.dtLastGuildMembership = Now
						End If
					End If
				Else
					Dim oMember As GuildMember = oGuild.GetMember(lMemberID)
					'if target is applied then acting must be within the guild and must have rights to approve apps... then approve app
					If oActingPlayer.lGuildID = lGuildID Then
						If (oMember.yMemberState And GuildMemberState.Applied) <> 0 Then
							If oGuild.RankHasPermission(oActingPlayer.lGuildRankID, RankPermissions.AcceptApplicant) = False Then
								LogEvent(LogEventType.PossibleCheat, "Player accepting applicant without permission: " & mlClientPlayer(lIndex))
								Return
							End If

							If oMember.oMember.oGuild Is Nothing = False Then
								Return
							End If

							oMember.yMemberState = oMember.yMemberState Or GuildMemberState.Approved
							oMember.oMember.lGuildID = oGuild.ObjectID
                            oGuild.MemberJoined(lMemberID, False)
                            oMember.oMember.dtLastGuildMembership = Now
						End If
					Else
						LogEvent(LogEventType.PossibleCheat, "Acting Player not in appropriate guild: " & mlClientPlayer(lIndex))
						Return
					End If
				End If
			Case GuildMemberState.Applied
				'target is applying for membership
				Dim oGuild As Guild = GetEpicaGuild(lGuildID)
				If oGuild Is Nothing = False Then
					If (oGuild.iRecruitFlags And eiRecruitmentFlags.GuildRecruiting) = 0 Then
						LogEvent(LogEventType.PossibleCheat, "Guild not recruiting but player applied: " & mlClientPlayer(lIndex))
						Return
					End If
					If oActingPlayer.oGuild Is Nothing = False Then Return

                    Dim oTmpPlayer As Player = GetEpicaPlayer(lMemberID)
                    If oTmpPlayer Is Nothing = False AndAlso (oTmpPlayer.dtLastGuildMembership <> Date.MinValue AndAlso oTmpPlayer.dtLastGuildMembership.AddDays(3) > Now) Then
                        Dim tsTemp As TimeSpan = oTmpPlayer.dtLastGuildMembership.AddDays(3).Subtract(Now)
                        Dim sMsg As String = "You cannot attempt to join a guild for another "
                        If tsTemp.Days > 0 Then sMsg &= tsTemp.Days.ToString & " days "
                        If tsTemp.Hours > 0 Then sMsg &= tsTemp.Hours.ToString & " hours "
                        If tsTemp.Minutes > 0 Then sMsg &= tsTemp.Minutes.ToString & " minutes "
                        Dim yResp() As Byte = CreateChatMsg(oTmpPlayer.ObjectID, sMsg, ChatMessageType.eSysAdminMessage)
                        oTmpPlayer.SendPlayerMessage(yResp, False, 0)
                        Return
                    End If

					Dim oMember As GuildMember = oGuild.GetMember(lMemberID)
                    If oMember Is Nothing = False Then
                        If oMember.oMember Is Nothing = False Then
                            If (oMember.yMemberState And (GuildMemberState.Applied Or GuildMemberState.Approved)) = 0 Then
                                oMember.yMemberState = oMember.yMemberState Or GuildMemberState.Applied
                                oGuild.PlayerApplied(lMemberID)
                            End If
                        End If
                    Else
                        oMember = New GuildMember()
                        oMember.lMemberID = lMemberID
                        oMember.yMemberState = GuildMemberState.Applied
                        oGuild.AddMember(oMember)
                        oGuild.PlayerApplied(lMemberID)
                    End If
				End If
			Case GuildMemberState.Rejected
				'if i am target and i am invited, i reject the invitation
				If lMemberID = oActingPlayer.ObjectID Then
					Dim oGuild As Guild = GetEpicaGuild(lGuildID)
					If oGuild Is Nothing Then Return
					Dim oMember As GuildMember = oGuild.GetMember(lMemberID)
					If oMember Is Nothing = False Then
						oMember.yMemberState = oMember.yMemberState Or GuildMemberState.Rejected
						oGuild.SendMsgToGuildMembers(yData)
					End If
				End If
		End Select


	End Sub
	Private Sub HandleGuildRequestDetails(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2	 'for msgcode

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return
		If oPlayer.oGuild Is Nothing Then Return

		Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim yRequestType As Byte = yData(lPos) : lPos += 1
		Dim iTypeID As Int16 = 0
		If yRequestType = eyGuildRequestDetailsType.GuildRel Then
			iTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		End If
		Dim yResp() As Byte = oPlayer.oGuild.HandleRequestDetails(CType(yRequestType, eyGuildRequestDetailsType), lItemID, iTypeID, oPlayer.lGuildRankID, oPlayer.ObjectID)
		If yResp Is Nothing = False Then
			moClients(lIndex).SendData(yResp)
		End If
	End Sub
	Private Sub HandleInviteFormGuild(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2

		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim yType As Byte = yData(lPos) : lPos += 1

		If yType = 1 Then
			'inviting...
			Dim sPlayerName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			sPlayerName = sPlayerName.ToUpper.Trim

			Dim oInviter As Player = GetEpicaPlayer(lPlayerID)
			If oInviter Is Nothing Then Return
			Dim oPlayer As Player = Nothing
			For X As Int32 = 0 To glPlayerUB
				If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).sPlayerName = sPlayerName Then
					oPlayer = goPlayer(X)
					Exit For
				End If
			Next X

			If oPlayer Is Nothing Then
				moClients(lIndex).SendData(CreateChatMsg(-1, "Could not find a player by that name.", ChatMessageType.eAllianceMessage))
				Dim yForward(22) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
				yForward(lPos) = 255 : lPos += 1
				oPlayer.PlayerName.CopyTo(yForward, lPos) : lPos += 20
				oInviter.SendPlayerMessage(yForward, False, 0)
			ElseIf oPlayer.oGuild Is Nothing = False Then
				moClients(lIndex).SendData(CreateChatMsg(-1, "That player is already in a guild.", ChatMessageType.eAllianceMessage))
				Dim yForward(22) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
				yForward(lPos) = 255 : lPos += 1
				oPlayer.PlayerName.CopyTo(yForward, lPos) : lPos += 20
				oInviter.SendPlayerMessage(yForward, False, 0)
			ElseIf oPlayer.ObjectID = lPlayerID Then
				moClients(lIndex).SendData(CreateChatMsg(-1, "Must invite someone other than yourself.", ChatMessageType.eAllianceMessage))
				Dim yForward(22) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
				yForward(lPos) = 255 : lPos += 1
				oPlayer.PlayerName.CopyTo(yForward, lPos) : lPos += 20
                oInviter.SendPlayerMessage(yForward, False, 0)
                'ElseIf oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
                '    moClients(lIndex).SendData(CreateChatMsg(-1, "That player is not yet out of the tutorial.", ChatMessageType.eAllianceMessage))
                '    Dim yForward(22) As Byte
                '    System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
                '    yForward(lPos) = 255 : lPos += 1
                '    oPlayer.PlayerName.CopyTo(yForward, lPos) : lPos += 20
                '    oInviter.SendPlayerMessage(yForward, False, 0)
            ElseIf oPlayer.dtLastGuildMembership <> Date.MinValue AndAlso oPlayer.dtLastGuildMembership.AddDays(3) > Now Then
                moClients(lIndex).SendData(CreateChatMsg(-1, "That player recently left a guild and cannot join another.", ChatMessageType.eAllianceMessage))
                Dim yForward(22) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
                yForward(lPos) = 255 : lPos += 1
                oPlayer.PlayerName.CopyTo(yForward, lPos) : lPos += 20
                oInviter.SendPlayerMessage(yForward, False, 0)
            Else
                Dim yForward(26) As Byte
                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
                yForward(lPos) = 1 : lPos += 1
                oInviter.PlayerName.CopyTo(yForward, lPos) : lPos += 20
                System.BitConverter.GetBytes(oInviter.ObjectID).CopyTo(yForward, lPos) : lPos += 4

                oPlayer.SendPlayerMessage(yForward, False, 0)
			End If
		ElseIf yType = 255 Then
			'rejecting
			Dim lInviter As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim oInviter As Player = GetEpicaPlayer(lInviter)
			If oInviter Is Nothing Then Return
			Dim yForward(22) As Byte
			lPos = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
			yForward(lPos) = 255 : lPos += 1
			Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
			oPlayer.PlayerName.CopyTo(yForward, lPos) : lPos += 20
			oInviter.SendPlayerMessage(yForward, False, 0)
		ElseIf yType = 2 Then
			'accepting
			Dim lInviter As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim oInviter As Player = GetEpicaPlayer(lInviter)
			If oInviter Is Nothing Then Return
			If oInviter.oFormingGuild Is Nothing Then Return

			If oInviter.oFormingGuild.oMembers Is Nothing Then ReDim oInviter.oFormingGuild.oMembers(-1)
			With oInviter.oFormingGuild
				Dim oLowestRank As GuildRank = .GetLowestRank()
				Dim bFound As Boolean = False
				For X As Int32 = 0 To .oMembers.GetUpperBound(0)
					If .oMembers(X) Is Nothing = False AndAlso .oMembers(X).lMemberID = lPlayerID Then
						.oMembers(X).yMemberState = .oMembers(X).yMemberState Or GuildMemberState.AcceptedGuildFormInvite
						If oLowestRank Is Nothing = False Then .oMembers(X).lCreateRankID = oLowestRank.lRankID
						bFound = True
						Exit For
					End If
				Next X
				If bFound = False Then
					ReDim Preserve .oMembers(.oMembers.GetUpperBound(0) + 1)
					.oMembers(.oMembers.GetUpperBound(0)) = New GuildMember
					Dim oMember As GuildMember = .oMembers(.oMembers.GetUpperBound(0))
					If oLowestRank Is Nothing = False Then oMember.lCreateRankID = oLowestRank.lRankID
					oMember.lMemberID = lPlayerID
					oMember.yMemberState = GuildMemberState.Invited Or GuildMemberState.AcceptedGuildFormInvite
				End If
			End With

			Dim yForward(26) As Byte
			lPos = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yForward, lPos) : lPos += 2
			yForward(lPos) = 2 : lPos += 1
			Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
			oPlayer.PlayerName.CopyTo(yForward, lPos) : lPos += 20
			System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yForward, lPos) : lPos += 4
			oInviter.SendPlayerMessage(yForward, False, 0)
		ElseIf yType = 128 Then
			'ok, inviter is kicking/cancelling the invited player
			Dim sPlayerName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			sPlayerName = sPlayerName.ToUpper.Trim

			Dim oInviter As Player = GetEpicaPlayer(lPlayerID)
			If oInviter Is Nothing Then Return
			If oInviter.oFormingGuild Is Nothing Then Return
			Dim oPlayer As Player = Nothing
			For X As Int32 = 0 To glPlayerUB
				If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).sPlayerName = sPlayerName Then
					oPlayer = goPlayer(X)
					Exit For
				End If
			Next X
			If oPlayer Is Nothing Then Return

			Dim yForward(10) As Byte
			lPos = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eGuildCreationAcceptance).CopyTo(yForward, lPos) : lPos += 2
			System.BitConverter.GetBytes(lPlayerID).CopyTo(yForward, lPos) : lPos += 4
			System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yForward, lPos) : lPos += 4
			yForward(lPos) = 255 : lPos += 1
			oInviter.oFormingGuild.SendMsgToGuildMembers(yForward)
			oInviter.oFormingGuild.RemoveMember(oPlayer.ObjectID)
			oPlayer.SendPlayerMessage(yData, False, 0)
		End If

	End Sub
	Private Sub HandleInvitePlayerToGuild(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim sPlayerName As String = GetStringFromBytes(yData, 2, 20)
		sPlayerName = sPlayerName.ToUpper.Trim

		Dim oInviter As Player = GetEpicaPlayer(lPlayerID)
		If oInviter Is Nothing Then Return
		If oInviter.oGuild Is Nothing Then Return
		If oInviter.oGuild.RankHasPermission(oInviter.lGuildRankID, RankPermissions.InviteMember) = False Then
			LogEvent(LogEventType.PossibleCheat, "Player inviting others without permission: " & mlClientPlayer(lIndex))
			Return
		End If

		'ok, 
		Dim oPlayer As Player = Nothing
		For X As Int32 = 0 To glPlayerUB
			If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).sPlayerName = sPlayerName Then
				oPlayer = goPlayer(X)
				Exit For
			End If
		Next X
		If oPlayer Is Nothing Then
			moClients(lIndex).SendData(CreateChatMsg(-1, "Could not find a player by that name.", ChatMessageType.eAllianceMessage))
		ElseIf oPlayer.oGuild Is Nothing = False Then
            moClients(lIndex).SendData(CreateChatMsg(-1, "That player is already in a guild.", ChatMessageType.eAllianceMessage))
        ElseIf oPlayer.dtLastGuildMembership <> Date.MinValue AndAlso oPlayer.dtLastGuildMembership.AddDays(3) > Now Then
            moClients(lIndex).SendData(CreateChatMsg(-1, "That player recently left a guild and cannot be invited to join a new one.", ChatMessageType.eAllianceMessage))
        Else
            'ok, associate the guild to the player in an invitation and add the player member object and then send the mssage to the player
            oPlayer.AddInvitedToJoin(oInviter.oGuild.ObjectID)

            If (oPlayer.iInternalEmailSettings And eEmailSettings.eGuildMembershipNotices) <> 0 Then
                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oInviter.sPlayerNameProper & " has invited you to join " & BytesToString(oInviter.oGuild.yGuildName) & ". To view your invitations, open the guild window.", "Guild Invitation", oInviter.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then
                    oPlayer.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
                oPC = Nothing
            End If

            Dim oMember As GuildMember = oInviter.oGuild.GetMember(oPlayer.ObjectID)
            If oMember Is Nothing = False Then
                oMember.yMemberState = oMember.yMemberState Or GuildMemberState.Invited
            Else
                oMember = New GuildMember()
                oMember.lMemberID = oPlayer.ObjectID
                oMember.yMemberState = GuildMemberState.Invited
                oInviter.oGuild.AddMember(oMember)
            End If

            Dim yForward() As Byte
            Dim yTmp() As Byte = oMember.GetObjectSmallString()
            If yTmp Is Nothing = False Then
                ReDim yForward(1 + yTmp.Length)
                System.BitConverter.GetBytes(GlobalMessageCode.eInvitePlayerToGuild).CopyTo(yForward, 0)
                yTmp.CopyTo(yForward, 2)
                oInviter.oGuild.SendMsgToGuildMembers(yForward)
            End If

		End If
	End Sub
	Private Sub HandleLoginRequest(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
		Dim X As Int32
		Dim Y As Int32
		Dim yUserName(19) As Byte
		Dim yPassword(19) As Byte
		Dim lID As Int32
		Dim ySame As Byte			'0 = no match, 1 = username match but password mismatch, 2 = perfect match
		Dim bAlias As Boolean = False
		Dim yAliasUserName(19) As Byte
		Dim yAliasPassword(19) As Byte
		Dim lAliasID As Int32
		Dim lRights As Int32 = 0

		Try

            Dim lEnvirID As Int32 = -1
            Dim iEnvirTypeID As Int16 = -1

            Dim oEnc As New StrEncDec

            Dim lPos As Int32 = 2   'for msgcode
            Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
            yUserName = oEnc.Decrypt(yUserName)

            Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 21
            yPassword = oEnc.Decrypt(yPassword)

            If yData(lPos) <> 0 Then bAlias = True
            lPos += 1
            Array.Copy(yData, lPos, yAliasUserName, 0, 20) : lPos += 20
            yAliasUserName = oEnc.Decrypt(yAliasUserName)
            Array.Copy(yData, lPos, yAliasPassword, 0, 20) : lPos += 21
            yAliasPassword = oEnc.Decrypt(yAliasPassword)

            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    With goPlayer(X)
                        ySame = 2

                        For Y = 0 To 19
                            If yUserName(Y) <> .PlayerUserName(Y) Then
                                ySame = 0
                                Exit For
                            End If
                        Next Y

                        If ySame = 2 Then
                            'Ok, check the password
                            For Y = 0 To 19
                                If yPassword(Y) <> .PlayerPassword(Y) Then
                                    ySame = 1
                                    Exit For
                                End If
                            Next Y

                            If ySame = 2 Then

                                LogEvent(LogEventType.Informational, "Login Account: " & BytesToString(.PlayerName))

                                If .lPlayerIcon = 0 Then
                                    lID = LoginResponseCodes.eAccountSetup
                                Else
                                    'Ok, found the player... check the account status
                                    Select Case .AccountStatus
                                        Case AccountStatusType.eActiveAccount, AccountStatusType.eOpenBetaAccount, AccountStatusType.eTrialAccount, AccountStatusType.eMondelisActive

                                            If .AccountStatus = AccountStatusType.eOpenBetaAccount AndAlso gb_IN_OPEN_BETA = False Then
                                                lID = LoginResponseCodes.eAccountInactive
                                            Else

                                                lID = glPlayerIdx(X)
                                                lAliasID = glPlayerIdx(X)

                                                'this is causing issues
                                                If goPlayer(X).oSocket Is Nothing = False Then
                                                    goPlayer(X).lAlreadyInUseCnt += 1
                                                    If goPlayer(X).lAlreadyInUseCnt > 10 Then
                                                        goPlayer(X).oSocket = Nothing
                                                        goPlayer(X).lAlreadyInUseCnt = 0
                                                        LogEvent(LogEventType.PossibleCheat, "AlreadyInUseCnt Exceeded 10 for player: " & goPlayer(X).ObjectID)
                                                    Else
                                                        lID = LoginResponseCodes.eAccountInUse
                                                        Exit For
                                                    End If
                                                End If

                                                If bAlias = True Then
                                                    ''Ok, now, do a lookup for the aliased player
                                                    Dim bFound As Boolean = False
                                                    'For Y = 0 To .lAllowanceUB
                                                    '    If .lAllowanceIdx(Y) <> -1 Then
                                                    '        bFound = True
                                                    '        For Z As Int32 = 0 To 19
                                                    '            If .uAllowanceLogin(Y).yUserName(Z) <> yAliasUserName(Z) OrElse .uAllowanceLogin(Y).yPassword(Z) <> yAliasPassword(Z) Then
                                                    '                bFound = False
                                                    '                Exit For
                                                    '            End If
                                                    '        Next Z

                                                    '        If bFound = True Then
                                                    '            'Ok, found it
                                                    '            lAliasID = .lAllowanceIdx(Y)
                                                    '            lRights = .uAllowanceLogin(Y).lRights
                                                    '            .lAliasingPlayerID = .oAllowances(Y).ObjectID
                                                    '            lEnvirID = .oAllowances(Y).lLastViewedEnvir
                                                    '            iEnvirTypeID = .oAllowances(Y).iLastViewedEnvirType

                                                    '            LogEvent(LogEventType.Informational, "Player " & goPlayer(X).sPlayerNameProper & " logging in as alias of " & .oAllowances(Y).sPlayerNameProper)
                                                    '            Exit For
                                                    '        End If
                                                    '    End If
                                                    'Next Y
                                                    'If bFound = False Then
                                                    '    'return bad result
                                                    '    ySame = 0
                                                    '    Exit For
                                                    'End If

                                                    If .ObjectID = 1 OrElse .ObjectID = 2 OrElse .ObjectID = 6 OrElse .ObjectID = 7 Then
                                                        For Y = 0 To glPlayerUB
                                                            If glPlayerIdx(Y) > -1 Then
                                                                Dim oTmp As Player = goPlayer(Y)
                                                                If oTmp Is Nothing = False Then
                                                                    bFound = True
                                                                    For Z As Int32 = 0 To 19
                                                                        If oTmp.PlayerUserName(Z) <> yAliasUserName(Z) Then
                                                                            bFound = False
                                                                            Exit For
                                                                        End If
                                                                    Next Z

                                                                    If bFound = True Then
                                                                        'Ok, found it
                                                                        lAliasID = oTmp.ObjectID
                                                                        lRights = AliasingRights.eAddProduction Or AliasingRights.eAddResearch Or AliasingRights.eAlterAgents Or AliasingRights.eAlterAutoLaunchPower Or AliasingRights.eAlterColonyStats Or _
                                                                          AliasingRights.eAlterDiplomacy Or AliasingRights.eAlterEmail Or AliasingRights.eAlterTrades Or AliasingRights.eCancelProduction Or AliasingRights.eCancelResearch Or _
                                                                          AliasingRights.eChangeBehavior Or AliasingRights.eChangeEnvironment Or AliasingRights.eCreateBattleGroups Or AliasingRights.eCreateDesigns Or AliasingRights.eDockUndockUnits Or _
                                                                          AliasingRights.eModifyBattleGroups Or AliasingRights.eMoveUnits Or AliasingRights.eTransferCargo Or AliasingRights.eViewAgents Or AliasingRights.eViewBattleGroups Or _
                                                                          AliasingRights.eViewBudget Or AliasingRights.eViewColonyStats Or AliasingRights.eViewDiplomacy Or AliasingRights.eViewEmail Or AliasingRights.eViewMining Or _
                                                                          AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns Or AliasingRights.eViewTrades Or AliasingRights.eViewTreasury Or AliasingRights.eViewUnitsAndFacilities
                                                                        .lAliasingPlayerID = oTmp.ObjectID
                                                                        lEnvirID = oTmp.lLastViewedEnvir
                                                                        iEnvirTypeID = oTmp.iLastViewedEnvirType

                                                                        LogEvent(LogEventType.Informational, "Player " & goPlayer(X).sPlayerNameProper & " logging in as alias of " & oTmp.sPlayerNameProper)
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        Next Y
                                                    Else
                                                        For Y = 0 To .lAllowanceUB
                                                            If .lAllowanceIdx(Y) <> -1 Then
                                                                bFound = True
                                                                For Z As Int32 = 0 To 19
                                                                    If .uAllowanceLogin(Y).yUserName(Z) <> yAliasUserName(Z) OrElse .uAllowanceLogin(Y).yPassword(Z) <> yAliasPassword(Z) Then
                                                                        bFound = False
                                                                        Exit For
                                                                    End If
                                                                Next Z

                                                                If bFound = True Then
                                                                    'Ok, found it
                                                                    lAliasID = .lAllowanceIdx(Y)
                                                                    lRights = .uAllowanceLogin(Y).lRights
                                                                    .lAliasingPlayerID = .oAllowances(Y).ObjectID
                                                                    lEnvirID = .oAllowances(Y).lLastViewedEnvir
                                                                    iEnvirTypeID = .oAllowances(Y).iLastViewedEnvirType

                                                                    If .oAllowances(Y).AccountStatus <> AccountStatusType.eActiveAccount AndAlso .oAllowances(Y).AccountStatus <> AccountStatusType.eMondelisActive Then
                                                                        lID = LoginResponseCodes.eAccountInactive
                                                                    ElseIf (.oAllowances(Y).ySpawnSystemSetting And (eySpawnSettings.eSpawnAccept Or eySpawnSettings.eSpawnRefuse)) = 0 Then
                                                                        lID = LoginResponseCodes.eAccountInactive
                                                                    End If

                                                                    LogEvent(LogEventType.Informational, "Player " & goPlayer(X).sPlayerNameProper & " logging in as alias of " & .oAllowances(Y).sPlayerNameProper)
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next Y
                                                    End If
                                                Else
                                                    lEnvirID = .lLastViewedEnvir
                                                    iEnvirTypeID = .iLastViewedEnvirType
                                                End If

                                                'Do our associations here...
                                                mlClientPlayer(lSocketIndex) = glPlayerIdx(X)
                                                moClients(lSocketIndex).lSpecificID = glPlayerIdx(X)

                                                Dim bCurtainCancelled As Boolean = False
                                                If lAliasID <> glPlayerIdx(X) Then
                                                    mlAliasedAs(lSocketIndex) = lAliasID
                                                    mlAliasedRights(lSocketIndex) = lRights

                                                    If lAliasID > 0 Then
                                                        Dim oAliasPlayer As Player = GetEpicaPlayer(lAliasID)
                                                        If oAliasPlayer Is Nothing = False Then
                                                            'oAliasPlayer.CancelIronCurtain()
                                                            If oAliasPlayer.lNext15MinuteReset < glCurrentCycle Then oAliasPlayer.lCurrent15MinutesRemaining = 27000
                                                            AddToQueue(glCurrentCycle + oAliasPlayer.lCurrent15MinutesRemaining, QueueItemType.eIronCurtainFall, oAliasPlayer.ObjectID, -1, -1, -1, 0, 0, 0, 0)
                                                            oAliasPlayer.lNext15MinuteReset = glCurrentCycle + gl_DELAY_FOUR_HOURS
                                                            bCurtainCancelled = True
                                                        End If
                                                    End If
                                                End If
                                                If bCurtainCancelled = False Then
                                                    If .lNext15MinuteReset < glCurrentCycle Then .lCurrent15MinutesRemaining = 27000
                                                    AddToQueue(glCurrentCycle + .lCurrent15MinutesRemaining, QueueItemType.eIronCurtainFall, .ObjectID, -1, -1, -1, 0, 0, 0, 0)
                                                    .lNext15MinuteReset = glCurrentCycle + gl_DELAY_FOUR_HOURS
                                                End If

                                                'NOTE: This is a hack!!!
                                                'HACK: This is a hack for cross primary
                                                'TODO: Fix this hack!
                                                .InMyDomain = True
                                                .lConnectedPrimaryID = glServerID
                                                'END OF HACK

                                                .bInPlayerRequestDetails = False
                                                .oSocket = moClients(lSocketIndex)

                                                'Set the last login time...
                                                .LastLogin = Now
                                                mlClientStatusFlags(lSocketIndex) = mlClientStatusFlags(lSocketIndex) Or ClientStatusFlags.ePlayerLoggedIn

                                                'Dim yPlayerLoginMsg() As Byte = CreateChatMsg(-1, BytesToString(goPlayer(X).PlayerName) & " has logged on.", ChatMessageType.eSysAdminMessage)
                                                'For lTmpIdx As Int32 = 0 To glPlayerUB
                                                '	If lTmpIdx <> X AndAlso glPlayerIdx(lTmpIdx) <> -1 Then
                                                '		goPlayer(lTmpIdx).SendPlayerMessage(yPlayerLoginMsg, False, 0)
                                                '	End If
                                                'Next lTmpIdx

                                                .FinalizeEnvirAlerts()

                                                If lID > 0 Then
                                                    'Check our spawn flags
                                                    If .yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso .AccountStatus = AccountStatusType.eActiveAccount Then
                                                        'Ok, full player with an active account, check the spawn flag
                                                        If (.ySpawnSystemSetting And (eySpawnSettings.eSpawnAccept Or eySpawnSettings.eSpawnRefuse)) = 0 Then
                                                            'Ok, must offer the player a spawn option
                                                            .ySpawnSystemSetting = .ySpawnSystemSetting Or eySpawnSettings.eOfferedSpawn
                                                            .QueueMeToSave()

                                                            'Now, send the player the option to spawn
                                                            Dim yForward(5) As Byte
                                                            System.BitConverter.GetBytes(GlobalMessageCode.eRespawnSelection).CopyTo(yForward, 0)
                                                            System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yForward, 2)
                                                            moClients(lSocketIndex).SendData(yForward)

                                                            'and return out
                                                            Return
                                                        End If
                                                    End If

                                                    'tell our email/chat server that the player is online
                                                    Dim yEmail(16) As Byte
                                                    Dim lAID As Int32 = mlAliasedAs(lSocketIndex)
                                                    If lAID < 1 OrElse lAID = mlClientPlayer(lSocketIndex) Then lAID = -1
                                                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yEmail, 0)
                                                    System.BitConverter.GetBytes(mlClientPlayer(lSocketIndex)).CopyTo(yEmail, 2)
                                                    System.BitConverter.GetBytes(lAID).CopyTo(yEmail, 6)
                                                    yEmail(10) = 1
                                                    System.BitConverter.GetBytes(lEnvirID).CopyTo(yEmail, 11)
                                                    System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yEmail, 15)
                                                    SendToEmailSrvr(yEmail) 'moEmailSrvr.SendData(yEmail)

                                                    Dim yOp(9) As Byte
                                                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerConnectedPrimary).CopyTo(yOp, 0)
                                                    System.BitConverter.GetBytes(mlClientPlayer(lSocketIndex)).CopyTo(yOp, 2)
                                                    System.BitConverter.GetBytes(0I).CopyTo(yOp, 6)
                                                    moOperator.SendData(yOp)

                                                    If .GuaranteedSpecialTechID < 1 AndAlso bAlias = False AndAlso .yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso .AccountStatus = AccountStatusType.eActiveAccount Then
                                                        'ok, send the player a message
                                                        Dim yForward(1) As Byte
                                                        System.BitConverter.GetBytes(GlobalMessageCode.eSetSpecTechGuarantee).CopyTo(yForward, 0)
                                                        moClients(lSocketIndex).SendData(yForward)
                                                    End If
                                                End If

                                                'Now, before we can continue, check the EnvirID and TypeID, if they are both 0, then the player is new
                                                If .lLastViewedEnvir = 0 AndAlso .iLastViewedEnvirType = 0 Then
                                                    If .yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                                                        InitializePlayer(goPlayer(X), Int32.MinValue)

                                                        Dim yFrwrd(5) As Byte
                                                        System.BitConverter.GetBytes(GlobalMessageCode.eCreatePlanetInstance).CopyTo(yFrwrd, 0)
                                                        System.BitConverter.GetBytes(.ObjectID).CopyTo(yFrwrd, 2)
                                                        moDomains(0).SendData(yFrwrd)
                                                    Else
                                                        'New player... Begin an Asynchronous response...
                                                        If goPlayer(X).lStartedEnvirID < 1 Then goPlayer(X).lStartedEnvirID = -1
                                                        InitializePlayer(goPlayer(X), goPlayer(X).lStartedEnvirID)
                                                    End If
                                                    Return
                                                ElseIf .lLastViewedEnvir >= 500000000 AndAlso .iLastViewedEnvirType = ObjectType.ePlanet Then
                                                    If .yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                                                        Dim yFrwrd(5) As Byte
                                                        System.BitConverter.GetBytes(GlobalMessageCode.eCreatePlanetInstance).CopyTo(yFrwrd, 0)
                                                        System.BitConverter.GetBytes(.ObjectID).CopyTo(yFrwrd, 2)
                                                        moDomains(0).SendData(yFrwrd)
                                                        Return
                                                    Else
                                                        .lLastViewedEnvir = .lStartedEnvirID : .iLastViewedEnvirType = .iStartedEnvirTypeID
                                                    End If
                                                End If

                                                If .yPlayerPhase = eyPlayerPhase.eSecondPhase Then
                                                    'ok, send an update message of their time remaining
                                                    If .lTutorialStep > 313 Then
                                                        'If .PlayedTimeWhenTimerStarted = Int32.MinValue Then .PlayedTimeWhenTimerStarted = .TotalPlayTime

                                                        'Dim lNumberOfSeconds As Int32 = 14400 - (.TotalPlayTime - .PlayedTimeWhenTimerStarted)
                                                        'AddToQueue(glCurrentCycle + (lNumberOfSeconds * 30), QueueItemType.eTutorialFourHourTimerExpire, .ObjectID, .ObjTypeID, -1, -1, 0, 0, 0, 0)

                                                        'Dim yFrwrd(6) As Byte
                                                        'System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerTimer).CopyTo(yFrwrd, 0)
                                                        'yFrwrd(2) = 0
                                                        'System.BitConverter.GetBytes(lNumberOfSeconds).CopyTo(yFrwrd, 3)
                                                        'moClients(lSocketIndex).SendData(yFrwrd)
                                                    End If
                                                End If

                                                If .iLastViewedEnvirType <> 2 AndAlso .iLastViewedEnvirType <> 3 Then
                                                    .lLastViewedEnvir = .lStartedEnvirID
                                                    .iLastViewedEnvirType = .iStartedEnvirTypeID
                                                Else
                                                    Dim oObj As Epica_GUID = CType(GetEpicaObject(.lLastViewedEnvir, .iLastViewedEnvirType), Epica_GUID)
                                                    If oObj Is Nothing Then
                                                        .lLastViewedEnvir = .lStartedEnvirID
                                                        .iLastViewedEnvirType = .iStartedEnvirTypeID
                                                        lEnvirID = .lLastViewedEnvir
                                                        iEnvirTypeID = .iLastViewedEnvirType
                                                    End If
                                                End If

                                                If lEnvirID < 1 OrElse iEnvirTypeID < 1 Then
                                                    lEnvirID = .lLastViewedEnvir
                                                    iEnvirTypeID = .iLastViewedEnvirType
                                                End If

                                            End If

                                        Case AccountStatusType.eBannedAccount
                                            lID = LoginResponseCodes.eAccountBanned
                                        Case AccountStatusType.eInactiveAccount
                                            lID = LoginResponseCodes.eAccountInactive
                                        Case AccountStatusType.eMondelisInactive
                                            lID = LoginResponseCodes.eMondelisInactive
                                        Case AccountStatusType.eSuspendedAccount
                                            lID = LoginResponseCodes.eAccountSuspended
                                        Case Else
                                            lID = LoginResponseCodes.eUnknownFailure
                                    End Select
                                End If

                                Exit For
                            ElseIf ySame = 1 Then
                                Exit For
                            End If
                        End If
                    End With

                End If
            Next X

            'Ok, check our ySame...
            If ySame = 0 Then
                lID = LoginResponseCodes.eInvalidUserName
            ElseIf ySame = 1 Then
                lID = LoginResponseCodes.eInvalidPassword
            End If

            'Now, send our response...
            If lID < 1 Then
                lAliasID = lID
                lRights = 0
            End If
            Dim yResp(13) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lAliasID).CopyTo(yResp, 2)
            System.BitConverter.GetBytes(lID).CopyTo(yResp, 6)
            System.BitConverter.GetBytes(lRights).CopyTo(yResp, 10)
            moClients(lSocketIndex).SendData(yResp)



        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "HandleLoginRequest: " & ex.Message)
            Try
                moClients(lSocketIndex).Disconnect()
            Catch
            End Try
        End Try
	End Sub
	Private Sub HandleMarkEmailReadStatus(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 Then
			lPlayerID = mlAliasedAs(lIndex)
			If (mlAliasedRights(lIndex) And AliasingRights.eAlterEmail) = 0 Then Return
		End If
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If oPlayer.InMyDomain = False Then
                SendPassThruMsg(yData, lPlayerID, oPlayer.ObjectID, oPlayer.ObjTypeID)
                Return
            End If

            DoMarkEmailReadStatus(oPlayer, yData, 0)
        End If
    End Sub
    Private Sub HandleMineralBidMsg(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lFacID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim oFac As Facility = GetEpicaFacility(lFacID)
        If oFac Is Nothing = False Then
            If oFac.oMiningBid Is Nothing Then

                If oFac.lCacheIndex > -1 Then
                    If oFac.lCacheIndex <= glMineralCacheUB AndAlso glMineralCacheIdx(oFac.lCacheIndex) = oFac.lCacheID Then
                        oFac.oMiningBid = New MineBuyOrderManager()
                        With oFac.oMiningBid
                            .oParentFacility = oFac
                            .oMineralCache = goMineralCache(oFac.lCacheIndex)
                            .lMaxDaysSold = 0
                            .lDaysNotSold = 0
                            .lCurrentConseqDays = 0
                            .bSomethingSold = False
                        End With
                    End If
                End If

            End If
            If oFac.oMiningBid Is Nothing = False Then
                moClients(lIndex).SendData(oFac.oMiningBid.GetBidAsMsg(mlClientPlayer(lIndex)))
            End If
        End If

    End Sub
    Private Sub HandleMoveEmailToFolder(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eAlterEmail) = 0 Then Return
        End If
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

        If oPlayer Is Nothing = False Then
            If oPlayer.InMyDomain = False Then
                SendPassThruMsg(yData, lPlayerID, oPlayer.ObjectID, oPlayer.ObjTypeID)
                Return
            End If

            Dim yResp() As Byte = DoMoveEmailToFolder(oPlayer, yData, 0)
            If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
        End If

    End Sub
    Private Sub HandlePlaceGuildBillboardBid(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ChangeRecruitment) = False Then
            LogEvent(LogEventType.PossibleCheat, "Player changing recruitment without permission: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim lPos As Int32 = 2   'for msgcode
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lSlotID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lAmt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDuration As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlanet As Planet = GetEpicaPlanet(lPlanetID)

        lSlotID += 1        'offset for the zero based

        If oPlanet Is Nothing = False Then
            If lAmt < 1 OrElse lDuration < 1 Then
                oPlanet.RemoveBid(oPlayer.oGuild, lSlotID)
            Else
                oPlanet.AddBid(lAmt, lDuration, oPlayer.oGuild, lSlotID, True)
            End If
        End If
    End Sub
    Private Sub HandlePlayerActivityReport(ByVal yData() As Byte, ByVal lIndex As Int32)
        'If True = True Then Return

        Dim lPos As Int32 = 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFPS As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTotalProcTime As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lVirtualMemory As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lWorkingSet As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)

        Dim oSB As New System.Text.StringBuilder

        Try
            Dim lNow As Int32 = CInt(Now.ToString("MMddHHmmss"))
            oSB.Append("INSERT INTO tblPlayerActivityHeader (PlayerID, DTStamp, FPS, ProcTime, VMem, WrkSet) VALUES (")
            oSB.Append(lPlayerID & ", " & lNow.ToString & ", " & lFPS & ", " & lTotalProcTime & ", " & lVirtualMemory & ", " & lWorkingSet & ")")
            oSB.AppendLine()

            Try
                Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(oSB.ToString, goCN)
                oComm.ExecuteNonQuery()
            Catch
            End Try
            oSB.Length = 0

            For X As Int32 = 0 To lCnt - 1
                Dim lActivityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lVal1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lVal2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lTS As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4


                oSB.AppendLine("INSERT INTO tblPlayerActivity (PlayerID, DTStamp, ActivityID, Val1, Val2, TS) VALUES (" & lPlayerID & ", " & lNow & ", " & lActivityID & ", " & lVal1 & ", " & lVal2 & ", " & lTS & ")")
            Next X

            If oSB.Length > 0 Then
                Dim oComm As OleDb.OleDbCommand = Nothing
                Try
                    oComm = New OleDb.OleDbCommand(oSB.ToString, goCN)
                    oComm.ExecuteNonQuery()
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "SaveActivityReportEx: " & ex.Message)
                End Try
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "SaveActivityReport: " & ex.Message)
        End Try


    End Sub
	Private Sub HandlePlayerAliasConfig(ByRef yData() As Byte, ByVal lIndex As Int32)
		If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
			LogEvent(LogEventType.PossibleCheat, "Aliased Player attempting to configure an alias: " & mlClientPlayer(lIndex))
			Return
		End If

		'MsgCode 2, Type(1), AliasPlayer(20), AliasUN(20), AliasPW(20), Rights(4)
		Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
		If oPlayer Is Nothing Then Return
		Dim lPos As Int32 = 2	'for msgcode
		Dim yType As Byte = yData(lPos) : lPos += 1
		Dim yPlayer(19) As Byte
		Array.Copy(yData, lPos, yPlayer, 0, 20) : lPos += 20

		If yType = 0 Then		'remove
			For X As Int32 = 0 To oPlayer.lAliasUB
				If oPlayer.lAliasIdx(X) <> -1 AndAlso oPlayer.oAliases(X) Is Nothing = False Then
					Dim bFound As Boolean = True

					Dim yTempName(19) As Byte
					StringToBytes(BytesToString(oPlayer.oAliases(X).PlayerName).ToUpper).CopyTo(yTempName, 0)

					For Y As Int32 = 0 To 19
						'If oPlayer.oAliases(X).PlayerName(Y) <> yPlayer(Y) Then
						If yTempName(Y) <> yPlayer(Y) Then
							bFound = False
							Exit For
						End If
					Next Y
					If bFound = True Then
                        Dim oOther As Player = oPlayer.oAliases(X)

                        Dim sSQL As String = "DELETE FROM tblAlias WHERE PlayerID = " & oPlayer.ObjectID & " AND OtherPlayerID = " & oOther.ObjectID
                        Dim oComm As OleDb.OleDbCommand = Nothing
                        Try
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            oComm.ExecuteNonQuery()
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "Unable to delete Alias between " & oPlayer.ObjectID & " and " & oOther.ObjectID & ": " & ex.Message)
                        Finally
                            If oComm Is Nothing = False Then oComm.Dispose()
                            oComm = Nothing
                        End Try

						oPlayer.lAliasIdx(X) = -1
						oPlayer.uAliasLogin(X).lRights = 0
						ReDim oPlayer.uAliasLogin(X).yPassword(19)
						ReDim oPlayer.uAliasLogin(X).yUserName(19)
						oPlayer.oAliases(X) = Nothing

						For Y As Int32 = 0 To oOther.lAllowanceUB
							If oOther.lAllowanceIdx(Y) = oPlayer.ObjectID Then
								oOther.lAllowanceIdx(Y) = -1
								oOther.oAllowances(Y) = Nothing
								oOther.uAllowanceLogin(Y).lRights = 0
								ReDim oOther.uAllowanceLogin(Y).yPassword(19)
								ReDim oOther.uAllowanceLogin(Y).yUserName(19)
							End If
						Next Y

						'TODO: if oOther is online currently as the alias, drop them
                        Dim oPC As PlayerComm = oOther.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "You no longer have aliasing rights for " & oPlayer.sPlayerNameProper & ".", _
                          "Player Alias Removed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oOther.sPlayerNameProper, Nothing)
						If oPC Is Nothing = False Then
                            If oOther.lConnectedPrimaryID > -1 Then oOther.CrossPrimarySafeSendMsg(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand))
						End If

						moClients(lIndex).SendData(yData)

						'Now, set the first 4 bytes of PlayerName to the PlayerID and the second 4 to OtherPlayerID
						System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yData, 3)
						System.BitConverter.GetBytes(oOther.ObjectID).CopyTo(yData, 7)

						'BroadcastToDomains(yData)
						moOperator.SendData(yData)

						Exit For
					End If
				End If
			Next X
		Else
			'add/update - get our remaining items
			Dim yUserName(19) As Byte
			Dim yPassword(19) As Byte
			Dim lRights As Int32 = 0
			Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
			Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 20
			lRights = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			'see if it is an update
			Dim lIdx As Int32 = -1
			For X As Int32 = 0 To oPlayer.lAliasUB
				If oPlayer.lAliasIdx(X) <> -1 AndAlso oPlayer.oAliases(X) Is Nothing = False Then
					Dim bFound As Boolean = True
					Dim yTempName(19) As Byte
					StringToBytes(BytesToString(oPlayer.oAliases(X).PlayerName).ToUpper).CopyTo(yTempName, 0)
					For Y As Int32 = 0 To 19
						'If oPlayer.oAliases(X).PlayerName(Y) <> yPlayer(Y) Then
						If yTempName(Y) <> yPlayer(Y) Then
							bFound = False
							Exit For
						End If
					Next Y
					If bFound = True Then
						lIdx = X
						Exit For
					End If
				End If
			Next X

			If lIdx = -1 Then
				'Ok, need to add an alias... find the player
				Dim oOther As Player = Nothing
				For Y As Int32 = 0 To glPlayerUB
					If glPlayerIdx(Y) <> -1 AndAlso goPlayer(Y) Is Nothing = False Then
						Dim bFound As Boolean = True
						Dim yTempName(19) As Byte
						StringToBytes(BytesToString(goPlayer(Y).PlayerName).ToUpper).CopyTo(yTempName, 0)
						For Z As Int32 = 0 To 19
							If yTempName(Z) <> yPlayer(Z) Then
								bFound = False
								Exit For
							End If
						Next Z
						If bFound = True Then
							oOther = goPlayer(Y)
							Exit For
						End If
					End If
				Next Y
				If oOther Is Nothing Then
					yData(2) = 255		'indicate failure to find player
					moClients(lIndex).SendData(yData)
					Return
				End If

                If oOther.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
                    moClients(lIndex).SendData(CreateChatMsg(-1, "That player is still in the tutorial and cannot be assigned alias accounts.", ChatMessageType.eSysAdminMessage))
                    Return
                End If
                If oOther.AccountStatus <> AccountStatusType.eActiveAccount Then
                    moClients(lIndex).SendData(CreateChatMsg(-1, "That player is in Trial and cannot be assigned alias accounts.", ChatMessageType.eSysAdminMessage))
                    Return
                End If

				oPlayer.AddPlayerAlias(oOther, yUserName, yPassword, lRights)
				oOther.AddPlayerAliasAllowance(oPlayer, yUserName, yPassword, lRights)
				'Send msg to everyone
				yData(2) = 1		'definitely an Add
				moClients(lIndex).SendData(yData)
                If oOther.lConnectedPrimaryID > -1 Then oOther.CrossPrimarySafeSendMsg(yData)

				'Now, set the first 4 bytes of PlayerName to the PlayerID and the second 4 to OtherPlayerID
				System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yData, 3)
				System.BitConverter.GetBytes(oOther.ObjectID).CopyTo(yData, 7)
				'BroadcastToDomains(yData)
				moOperator.SendData(yData)

                Dim oPC As PlayerComm = oOther.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "You have been granted permission to access portions of " & oPlayer.sPlayerNameProper & "'s empire." & vbCrLf & vbCrLf & _
                   "To do so, you will need to log-in with an alias. Use your normal username and password and click the Alias Player checkbox." & vbCrLf & vbCrLf & _
                   "In the Alias Username and Alias Password fields enter the following credentials:" & vbCrLf & vbCrLf & _
                   "Alias Username: " & GetStringFromBytes(yUserName, 0, 20), "New Player Alias Data", oPlayer.ObjectID, GetDateAsNumber(Now), False, _
                   oOther.sPlayerNameProper, Nothing)
				If oPC Is Nothing = False Then
                    If oOther.lConnectedPrimaryID > -1 Then oOther.CrossPrimarySafeSendMsg(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand))
				End If
			Else
				'An update...
				Dim oOther As Player = oPlayer.oAliases(lIdx)
				If oOther Is Nothing = False Then
					oPlayer.AddPlayerAlias(oOther, yUserName, yPassword, lRights)
					oOther.AddPlayerAliasAllowance(oPlayer, yUserName, yPassword, lRights)

					yData(2) = 2	'indicates an update
					moClients(lIndex).SendData(yData)
					'Inform the other player
                    If oOther.lConnectedPrimaryID > -1 Then oOther.CrossPrimarySafeSendMsg(yData)

					'now, send response to domains 
					'Now, set the first 4 bytes of PlayerName to the PlayerID and the second 4 to OtherPlayerID
					System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yData, 3)
					System.BitConverter.GetBytes(oOther.ObjectID).CopyTo(yData, 7)
					'BroadcastToDomains(yData)
					moOperator.SendData(yData)

                    Dim oPC As PlayerComm = oOther.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "You have been granted permission to access portions of " & oPlayer.sPlayerNameProper & "'s empire." & vbCrLf & vbCrLf & _
                       "To do so, you will need to log-in with an alias. Use your normal username and password and click the Alias Player checkbox." & vbCrLf & vbCrLf & _
                       "In the Alias Username and Alias Password fields enter the following credentials:" & vbCrLf & vbCrLf & _
                       "Alias Username: " & GetStringFromBytes(yUserName, 0, 20), "New Player Alias Data", oPlayer.ObjectID, GetDateAsNumber(Now), False, _
                       oOther.sPlayerNameProper, Nothing)
					If oPC Is Nothing = False Then
                        If oOther.lConnectedPrimaryID > -1 Then oOther.CrossPrimarySafeSendMsg(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand))
					End If
				End If
			End If

		End If

	End Sub
	Private Sub HandlePlayerIsDead(ByRef yData() As Byte, ByVal lIndex As Int32)
		'Ok, player is telling us they are dead... get them
		Dim yRespawnFlag As Byte = yData(2)

		Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing = False AndAlso oPlayer.PlayerIsDead = True Then

            Dim yForward(7) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yForward, 0)
            System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yForward, 2)
            If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount Then yRespawnFlag = 255

            yForward(6) = yRespawnFlag
            'If oPlayer.bForcedRestart = True Then yForward(7) = 1 Else 
            yForward(7) = 0

            moOperator.SendData(yForward)

            ''Store our seconds played here in case we need it
            'Dim lSeconds As Int32 = oPlayer.TotalPlayTime
            'If oPlayer.oSocket Is Nothing = False Then
            '    lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
            'End If

            'Send this to disconnect the player
            Dim yResp(2) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yResp, 0) : lPos += 2
            yResp(lPos) = 2 : lPos += 1
            oPlayer.SendPlayerMessage(yResp, False, 0)

            ''where the player currently resides
            'If oPlayer.lStartedEnvirID < 500000000 Then
            '    Dim oTmpPlanet As Planet = GetEpicaPlanet(oPlayer.lStartedEnvirID)
            '    If oTmpPlanet Is Nothing = False OrElse oTmpPlanet.ParentSystem Is Nothing = False Then
            '        Dim yGNS(81) As Byte
            '        'Dim lPos As Int32 = 0
            '        lPos = 0
            '        System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
            '        yGNS(lPos) = NewsItemType.ePlayerDeath : lPos += 1
            '        oTmpPlanet.ParentSystem.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
            '        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
            '        System.BitConverter.GetBytes(oPlayer.blMaxPopulation).CopyTo(yGNS, lPos) : lPos += 8
            '        oTmpPlanet.PlanetName.CopyTo(yGNS, lPos) : lPos += 20

            '        oPlayer.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
            '        yGNS(lPos) = oPlayer.yGender : lPos += 1
            '        oPlayer.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

            '        SendToEmailSrvr(yGNS)
            '    End If
            '    oTmpPlanet = Nothing
            'End If

            'If oPlayer.DeathBudgetBalance <> -1 AndAlso oPlayer.bForcedRestart = False Then
            '    'Ok, we're good to go, the player wants to reset
            '    oPlayer.blCredits = oPlayer.DeathBudgetBalance
            '    If oPlayer.blCredits < 500000 Then oPlayer.blCredits += 500000
            '    oPlayer.DeathBudgetBalance = -1 'set to -1 as another state machine

            '    'Begin 24 hour timer on the death budget
            '    oPlayer.DeathBudgetEndTime = glCurrentCycle + 2592000       '24 hrs
            'Else
            '    oPlayer.DeathBudgetBalance = 0
            '    oPlayer.blCredits = 500000
            '    'Begin 24 hour timer on the death budget
            '    oPlayer.DeathBudgetEndTime = glCurrentCycle + 2592000       '24 hrs
            'End If

            'For X As Int32 = 0 To glUnitUB
            '    If glUnitIdx(X) <> -1 Then
            '        Dim oUnit As Unit = goUnit(X)
            '        If oUnit Is Nothing Then Continue For
            '        If oUnit.Owner.ObjectID = oPlayer.ObjectID Then
            '            Dim iTemp As Int16 = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
            '            If iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
            '                Dim yMsg(8) As Byte
            '                System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yMsg, 0)
            '                oUnit.GetGUIDAsString.CopyTo(yMsg, 2)
            '                yMsg(8) = 0

            '                If iTemp = ObjectType.ePlanet Then
            '                    CType(oUnit.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
            '                Else : CType(oUnit.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
            '                End If
            '            End If
            '            Try
            '                oUnit.DeleteEntity(X)
            '            Catch
            '            End Try
            '            glUnitIdx(X) = -1
            '        End If
            '    End If
            'Next X

            'For X As Int32 = 0 To glFacilityUB
            '    If glFacilityIdx(X) <> -1 Then
            '        Dim oFac As Facility = goFacility(X)
            '        If oFac Is Nothing = False AndAlso oFac.Owner.ObjectID = oPlayer.ObjectID Then
            '            If oFac.ParentObject Is Nothing = False Then
            '                Dim iTemp As Int16 = CType(oFac.ParentObject, Epica_GUID).ObjTypeID
            '                If iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
            '                    Dim yMsg(8) As Byte
            '                    System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yMsg, 0)
            '                    oFac.GetGUIDAsString.CopyTo(yMsg, 2)
            '                    yMsg(8) = 0

            '                    If iTemp = ObjectType.ePlanet Then
            '                        CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
            '                    Else : CType(oFac.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
            '                    End If
            '                End If
            '                Try
            '                    oFac.DeleteEntity(X)
            '                Catch
            '                End Try
            '            End If
            '            glFacilityIdx(X) = -1
            '        End If
            '    End If
            'Next X

            'oPlayer.blMaxPopulation = 0

            ''remove agents belonging to me
            'goAgentEngine.CancelPlayerAgency(oPlayer.ObjectID)

            'goGTC.CancelPlayersOrders(oPlayer.ObjectID)

            'Dim yRelRemove(9) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eRemovePlayerRel).CopyTo(yRelRemove, 0)
            'System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yRelRemove, 2)
            'System.BitConverter.GetBytes(-1I).CopyTo(yRelRemove, 6)
            'moOperator.SendData(yRelRemove)

            'oPlayer.lWarSentiment = 0
            'oPlayer.BadWarDecCPIncrease = 0
            'oPlayer.BadWarDecCPIncreaseEndCycle = 0
            'oPlayer.BadWarDecMoralePenalty = 0
            'oPlayer.BadWarDecMoralePenaltyEndCycle = 0

            'Dim yNewMsg(11) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yNewMsg, 0)
            'System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yNewMsg, 2)
            'System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty).CopyTo(yNewMsg, 6)
            'System.BitConverter.GetBytes(oPlayer.BadWarDecCPIncrease).CopyTo(yNewMsg, 8)
            'goMsgSys.BroadcastToDomains(yNewMsg)

            ''Remove our factions
            'For X As Int32 = 0 To 4
            '    If oPlayer.lSlotID(X) > 0 Then
            '        'ok, need to remove me from the faction of the player
            '        Dim oSlotPlayer As Player = GetEpicaPlayer(oPlayer.lSlotID(X))
            '        If oSlotPlayer Is Nothing = False Then
            '            For Y As Int32 = 0 To 2
            '                If oSlotPlayer.lFactionID(Y) = oPlayer.ObjectID Then
            '                    oSlotPlayer.lFactionID(Y) = -1
            '                    oSlotPlayer.ReverifySlots()
            '                    Exit For
            '                End If
            '            Next Y
            '        End If
            '    End If
            '    oPlayer.lSlotID(X) = -1
            '    oPlayer.ySlotState(X) = 0
            'Next X
            'For X As Int32 = 0 To 2
            '    If oPlayer.lFactionID(X) > 0 Then
            '        Dim oFactionPlayer As Player = GetEpicaPlayer(oPlayer.lFactionID(X))
            '        If oFactionPlayer Is Nothing = False Then
            '            For Y As Int32 = 0 To 4
            '                If oFactionPlayer.lSlotID(Y) = oPlayer.ObjectID Then
            '                    oFactionPlayer.lSlotID(Y) = -1
            '                    oFactionPlayer.ReverifySlots()
            '                    Exit For
            '                End If
            '            Next Y
            '        End If
            '    End If
            '    oPlayer.lFactionID(X) = -1
            'Next X
            'oPlayer.ReverifySlots()

            'oPlayer.PlayerIsDead = False

            ''If oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase AndAlso yRespawnFlag <> 255 Then
            ''	'Ok, player is in the initial tutorial phase...
            ''	' - add Aurelius as a diplomatic contact and set Aurelius to blood war but my rel towards aurelius to neutral
            ''	Dim oPlayer2 As Player = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
            ''	If oPlayer2 Is Nothing = False Then
            ''		Dim oTmpRel As PlayerRel = New PlayerRel
            ''		oTmpRel.oPlayerRegards = oPlayer
            ''		oTmpRel.oThisPlayer = oPlayer2
            ''		oTmpRel.WithThisScore = elRelTypes.eNeutral
            ''		oPlayer.SetPlayerRel(gl_HARDCODE_PIRATE_PLAYER_ID, oTmpRel)

            ''		Dim yMsg2() As Byte = GetAddPlayerRelMessage(oTmpRel)
            ''		For X As Int32 = 0 To mlDomainUB
            ''			moDomains(X).SendData(yMsg2)
            ''		Next X
            ''	End If

            ''	' - store the time played so far and reset it to 0
            ''	oPlayer.PlayedTimeInTutorialOne = lSeconds
            ''	oPlayer.TotalPlayTime = 0

            ''	' - update phase to phase 2
            ''	oPlayer.yPlayerPhase = eyPlayerPhase.eSecondPhase

            ''	' - spawn the player in tutorial system 2
            ''	' Because of the above, when the player logs back in, they will be in tutorial system 2

            ''ElseIf oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
            'If yRespawnFlag = 254 Then oPlayer.yRespawnWithGuild = 1
            'If oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
            '    'Player is in second or higher tutorial phase, from here...
            '    ' - remove all technologies of the player and all Mineral Knowledge and any score-based values, gns values, running tallies, etc...
            '    oPlayer.WipePlayer()
            '    ' - remove all death budget stuff

            '    oPlayer.lLastViewedEnvir = 0
            '    oPlayer.iLastViewedEnvirType = 0
            '    oPlayer.lStartedEnvirID = -1
            '    oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase
            '    oPlayer.DeathBudgetEndTime = -1
            '    oPlayer.DeathBudgetBalance = 0
            '    oPlayer.DeathBudgetFundsRemaining = 0

            '    ''The player is going to get disconnected soon...
            '    'Dim yResp(2) As Byte
            '    'Dim lPos As Int32 = 0
            '    'System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yResp, 0) : lPos += 2
            '    'yResp(lPos) = 2 : lPos += 1
            '    'oPlayer.SendPlayerMessage(yResp, False, 0)

            '    'With the above command, the player will disconnect, we do not call initialize player
            '    'InitializePlayer(oPlayer, -1)
            'Else
            '    'normal death
            '    InitializePlayer(oPlayer, oPlayer.lStartedEnvirID)
            'End If

        End If
	End Sub
	Private Sub HandlePlayerRestart(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        Dim bSet As Boolean = yData(2) <> 0
        If oPlayer Is Nothing = False Then

            'Dim bSetForcedRestart As Boolean = True
            'If oPlayer.PlayerIsDead = True Then
            '    If oPlayer.bForcedRestart = False Then
            '        bSetForcedRestart = False
            '    End If
            'End If

            oPlayer.PlayerIsDead = bSet
            'If bSetForcedRestart = True Then oPlayer.bForcedRestart = bSet
            If bSet = False Then
                oPlayer.HandleCheckForPlayerDeath()
            End If
        End If
	End Sub
	Private Sub HandleProposeGuildVote(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

		If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ProposeVotes) = False Then
			LogEvent(LogEventType.PossibleCheat, "Player proposing votes without permission: " & mlClientPlayer(lIndex))
			Return
		End If

		Dim lPos As Int32 = 2		'for msgcode
		Dim yTypeOfVote As Byte = yData(lPos) : lPos += 1
		Dim lSelectedItem As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lNewValueNumber As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim sNewValueText As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim lDuration As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim sSummary As String = GetStringFromBytes(yData, lPos, 255) : lPos += 255

        'If oPlayer.oGuild.VoteAlreadyInProgress(yTypeOfVote, lSelectedItem) = True Then
        '    moClients(lIndex).SendData(CreateChatMsg(-1, "A vote for that is already in progress.", ChatMessageType.eNotificationMessage))
        '    Return
        'End If

        Dim oSB As New System.Text.StringBuilder()
        Dim oVote As New GuildVote()
		With oVote
			.dtVoteStarts = Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
			.lNewValue = lNewValueNumber
			.lSelectedItem = lSelectedItem
			.lVoteDuration = lDuration
			.ProposedByID = lPlayerID
			.VoteID = -1
			.yNewValueText = System.Text.ASCIIEncoding.ASCII.GetBytes(sNewValueText)
			.ySummary = System.Text.ASCIIEncoding.ASCII.GetBytes(sSummary)
			.yTypeOfVote = yTypeOfVote

            If .SaveObject(oPlayer.oGuild.ObjectID) = False Then Return

            'A Guild proposal has been submitted by <name>. The proposal starts on <start> and lasts <duration> hours. Proposal Summary: <summary>
            oSB.AppendLine("A guild proposal has been submitted by " & oPlayer.sPlayerNameProper & ". The proposal starts on " & .dtVoteStarts.ToString & " (GMT) and lasts " & .lVoteDuration.ToString & " hours. Proposal Summary: ")
            oSB.AppendLine(sSummary)
        End With
		oPlayer.oGuild.AddVote(oVote)

        oPlayer.oGuild.SendMsgToGuildMembersWithOutbound(oVote.GetVoteObjectAsAdd(), oSB.ToString, "Guild Vote Proposal", oPlayer.ObjectID)

	End Sub
	Private Sub HandleRemoveEventAttachment(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

		Dim lPos As Int32 = 2	'for msgcode
		Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lAttachID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oEvent As GuildEvent = oPlayer.oGuild.GetEvent(lEventID)
		If oEvent Is Nothing = False Then
			If oEvent.lPostedBy = oPlayer.ObjectID OrElse (oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.DeleteEvents) = True) Then
				oEvent.DeleteAttachment(lAttachID)
				oPlayer.oGuild.SendMsgToGuildMembers(yData)
			Else
				LogEvent(LogEventType.PossibleCheat, "Player attempting to delete event attachment without permission: " & mlClientPlayer(lIndex))
				Return
			End If
		End If

	End Sub
	Private Sub HandleRemoveFormation(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2	'for msgcode
		Dim lFormationID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		For X As Int32 = 0 To glFormationDefUB
			If glFormationDefIdx(X) = lFormationID Then
				Dim oFormation As FormationDef = goFormationDefs(X)
				glFormationDefIdx(X) = -1
				If oFormation Is Nothing = False Then
					oFormation.DeleteMe()
				End If
				Exit For
			End If
		Next X

		moClients(lIndex).SendData(yData)
	End Sub
	Private Sub HandleRemoveFromFleet(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, 6)

		Dim oFleet As UnitGroup = CType(GetEpicaObject(lFleetID, ObjectType.eUnitGroup), UnitGroup)

		If oFleet Is Nothing = False Then
			If oFleet.oOwner.ObjectID = mlClientPlayer(lIndex) OrElse (oFleet.oOwner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eModifyBattleGroups) <> 0) Then
				If oFleet.CanAlterComposition = True Then
					oFleet.RemoveUnit(lUnitID, False, False)
				Else
					moClients(lIndex).SendData(GetAddObjectMessage(oFleet, GlobalMessageCode.eAddObjectCommand))
				End If
			Else
				LogEvent(LogEventType.PossibleCheat, "HandleAddToFleet Owner Mismatch: " & mlClientPlayer(lIndex))
			End If
		End If
	End Sub
	Private Sub HandleRemoveRouteItem(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oUnit As Unit = GetEpicaUnit(lUnitID)
		If oUnit Is Nothing Then Return
		If oUnit.Owner.ObjectID = mlClientPlayer(lIndex) OrElse (oUnit.Owner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eMoveUnits) <> 0) Then
			oUnit.RemoveRouteItem(lValue)
			moClients(lIndex).SendData(yData)
		Else : LogEvent(LogEventType.PossibleCheat, "HandleUpdateRouteStatus: Owner Mismatch. PlayerID: " & mlClientPlayer(lIndex))
		End If
	End Sub
	Private Sub HandleRepairOrder(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPos As Int32 = 2
		Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lRepairItem As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'If iParentTypeID <> ObjectType.eFacility Then Return

        'Dim oParent As Facility = GetEpicaFacility(lParentID)
        Dim oParent As Epica_Entity = GetEpicaEntity(lParentID, iParentTypeID)
		Dim bRepairMyself As Boolean = False

		If oParent Is Nothing OrElse (oParent.Owner.ObjectID <> mlClientPlayer(lIndex) AndAlso (oParent.Owner.ObjectID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eAddProduction) = 0)) Then Return

		'Only production-capable facilities can repair, unless the repair order is myself...
		'If (oParent.yProductionType And ProductionType.eProduction) = 0 Then
		If lParentID = lEntityID AndAlso iParentTypeID = iEntityTypeID Then
			bRepairMyself = True
        ElseIf (oParent.yProductionType And ProductionType.eProduction) = 0 AndAlso oParent.yProductionType <> ProductionType.eCommandCenterSpecial AndAlso oParent.yProductionType <> ProductionType.eSpaceStationSpecial Then
            Return
		End If
		'End If

		Dim oEntity As Epica_Entity = Nothing
		Dim oDef As Epica_Entity_Def = Nothing

		If iEntityTypeID = ObjectType.eUnit Then
			oEntity = GetEpicaUnit(lEntityID)
			oDef = CType(oEntity, Unit).EntityDef
		ElseIf iEntityTypeID = ObjectType.eFacility Then
			oEntity = GetEpicaFacility(lEntityID)
            oDef = CType(oEntity, Facility).EntityDef
        ElseIf iEntityTypeID = ObjectType.eUnitDef Then
            For X As Int32 = 0 To oParent.lHangarUB
                If oParent.lHangarIdx(X) > -1 Then
                    oEntity = CType(oParent.oHangarContents(X), Epica_Entity)
                    If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
                        If CType(oEntity, Unit).EntityDef.ObjectID = lEntityID Then
                            oParent.HandleRepairItem(oEntity, CType(oEntity, Unit).EntityDef, lRepairItem)
                        End If
                    End If
                End If
            Next X
            Return
        ElseIf lEntityID = -1 AndAlso iEntityTypeID = -1 Then
            For X As Int32 = 0 To oParent.lHangarUB
                If oParent.lHangarIdx(X) > -1 Then
                    oEntity = CType(oParent.oHangarContents(X), Epica_Entity)
                    If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
                        oParent.HandleRepairItem(oEntity, CType(oEntity, Unit).EntityDef, lRepairItem)
                    End If
                End If
            Next X
            Return
        End If

		If oEntity Is Nothing OrElse oDef Is Nothing Then Return
		If oDef.oPrototype Is Nothing Then Return

		'Ok, verify that the entity is in the parent (for facility repair)
		Dim bFound As Boolean = False
		If bRepairMyself = False Then
			For X As Int32 = 0 To oParent.lHangarUB
				If oParent.lHangarIdx(X) = lEntityID AndAlso oParent.oHangarContents(X).ObjTypeID = iEntityTypeID Then
					bFound = True
					Exit For
				End If
			Next X
		Else : bFound = bRepairMyself
		End If
		If bFound = False Then
			LogEvent(LogEventType.PossibleCheat, "Parent ordered to repair entity that is not in hangar. PlayerID: " & mlClientPlayer(lIndex))
		Else
			oParent.HandleRepairItem(oEntity, oDef, lRepairItem)
		End If
	End Sub
	Private Sub HandleRequestAliasConfigs(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		Dim bAliased As Boolean = GetEpicaPlayer(mlClientPlayer(lIndex)).lAliasingPlayerID > 0
        'If bAliased = True Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		Dim yCache(200000) As Byte
		Dim yFinal() As Byte
		Dim lPos As Int32
		Dim lSingleMsgLen As Int32
		Dim yTemp() As Byte

		lPos = 0
        lSingleMsgLen = -1
        If bAliased = False Then
            For Y As Int32 = 0 To oPlayer.lAliasUB
                If oPlayer.lAliasIdx(Y) <> -1 Then
                    Dim yMsg(66) As Byte
                    Dim lTempPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAliasConfig).CopyTo(yMsg, lTempPos) : lTempPos += 2
                    yMsg(lTempPos) = 254 : lTempPos += 1        'indicates to add to the client's alias list
                    oPlayer.oAliases(Y).PlayerName.CopyTo(yMsg, lTempPos) : lTempPos += 20
                    With oPlayer.uAliasLogin(Y)
                        .yUserName.CopyTo(yMsg, lTempPos) : lTempPos += 20
                        .yPassword.CopyTo(yMsg, lTempPos) : lTempPos += 20
                        System.BitConverter.GetBytes(.lRights).CopyTo(yMsg, lTempPos) : lTempPos += 4
                    End With
                    yTemp = yMsg
                    lSingleMsgLen = yTemp.Length
                    'Ok, before we continue, check if we need to increase our cache
                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                        ReDim Preserve yCache(yCache.Length + 200000)
                    End If
                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                    lPos += 2
                    yTemp.CopyTo(yCache, lPos)
                    lPos += lSingleMsgLen
                End If
            Next Y
        End If

        For Y As Int32 = 0 To oPlayer.lAllowanceUB
            If oPlayer.lAllowanceIdx(Y) <> -1 Then
                Dim yMsg(66) As Byte
                Dim lTempPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAliasConfig).CopyTo(yMsg, lTempPos) : lTempPos += 2
                yMsg(lTempPos) = 253 : lTempPos += 1        'indicates to add to the client's alias list
                Dim oTmp As Player = GetEpicaPlayer(oPlayer.lAllowanceIdx(Y))
                If oTmp Is Nothing Then Return
                oTmp.PlayerName.CopyTo(yMsg, lTempPos) : lTempPos += 20
                With oPlayer.uAllowanceLogin(Y)
                    .yUserName.CopyTo(yMsg, lTempPos) : lTempPos += 20
                    .yPassword.CopyTo(yMsg, lTempPos) : lTempPos += 20
                    System.BitConverter.GetBytes(.lRights).CopyTo(yMsg, lTempPos) : lTempPos += 4
                End With
                yTemp = yMsg
                lSingleMsgLen = yTemp.Length
                'Ok, before we continue, check if we need to increase our cache
                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                    ReDim Preserve yCache(yCache.Length + 200000)
                End If
                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                lPos += 2
                yTemp.CopyTo(yCache, lPos)
                lPos += lSingleMsgLen
            End If
        Next Y

        If lPos <> 0 Then
            lPos += 1               'MSC - 07/29/08 - dont know why but this fixes an issue with getting the alias config data... looks like a possible off by one error somewhere in the code
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            moClients(lIndex).SendLenAppendedData(yFinal)
        End If
    End Sub
    Private Sub HandleRequestBidStatus(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewMining Or AliasingRights.eViewMining)) = 0 Then Return
        End If

        Dim lPos As Int32 = 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim yResp(5 + (lCnt * 5)) As Byte
        Dim lRespPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestBidStatus).CopyTo(yResp, lRespPos) : lRespPos += 2
        System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lRespPos) : lRespPos += 4

        For X As Int32 = 0 To lCnt - 1
            Dim lFacID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            If lFacID > 0 Then
                Dim oFac As Facility = GetEpicaFacility(lFacID)
                If oFac Is Nothing = False Then
                    If oFac.oMiningBid Is Nothing = False Then
                        Dim yType As Byte = oFac.oMiningBid.GetPlayerBidStatus(lPlayerID)
                        System.BitConverter.GetBytes(oFac.lCacheID).CopyTo(yResp, lRespPos) : lRespPos += 4
                        yResp(lRespPos) = yType : lRespPos += 1
                    End If
                End If
            End If
        Next X

        moClients(lIndex).SendData(yResp)
    End Sub
	Private Sub HandleRequestBugList(ByVal lClientIndex As Int32)
		'Ok, player wants the bug list...
		'Dim X As Int32
		'Dim yCache(0) As Byte
		'Dim lPos As Int32
		'Dim yTemp() As Byte

		'lPos = 0
		'For X = 0 To glBugUB
		'    yTemp = GetAddObjectMessage(goBugs(X), EpicaMessageCode.eBugList)

		'    ReDim Preserve yCache(yCache.Length + yTemp.Length + 2)

		'    System.BitConverter.GetBytes(CShort(yTemp.Length)).CopyTo(yCache, lPos)
		'    lPos += 2
		'    yTemp.CopyTo(yCache, lPos)
		'    lPos += yTemp.Length
		'Next X

		'If lPos <> 0 Then
		'    moClients(lClientIndex).SendLenAppendedData(yCache)
		'End If

		'For X As Int32 = 0 To glBugUB
		'	If goBugs(X).yStatus <> 7 Then moClients(lClientIndex).SendData(GetAddObjectMessage(goBugs(X), GlobalMessageCode.eBugList))
		'Next X
		Dim yResp() As Byte = EpicaBug.GetBugListResponse()
		If yResp Is Nothing = False Then moClients(lClientIndex).SendLenAppendedData(yResp)

    End Sub
    Private Sub HandleRequestChannelDetails(ByVal yData() As Byte, ByVal lIndex As Int32)
        'Ok, the client gives us room for the player ID
        Dim lPos As Int32 = 2       'msgcode...
        System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yData, lPos) : lPos += 4
        'Now, simply forward to the email server
        moEmailSrvr.SendData(yData)
    End Sub
    Private Sub HandleRequestChannelList(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 0
        'client gives us room for the playerid
        Dim yForward(6) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestChannelList).CopyTo(yForward, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yForward, lPos) : lPos += 4
        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing Then Return
        If oPlayer.yPlayerTitle = Player.PlayerRank.Emperor Then yForward(lPos) = 1 Else yForward(lPos) = 0
        lPos += 1
        moEmailSrvr.SendData(yForward)
    End Sub
    Private Sub HandleRequestColonyTransportPotentials(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim lPos As Int32 = 2
            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)
            If lColonyIdx > -1 AndAlso lColonyIdx <= glColonyUB Then
                Dim oColony As Colony = goColony(lColonyIdx)

                'Create our temporary arrays
                Dim oUnits(-1) As Unit
                Dim lUnitUB As Int32 = -1

                'Now, loop thru facilities looking in their hangar for units
                Dim lUB As Int32 = -1
                If oColony.lChildrenIdx Is Nothing = False Then lUB = Math.Min(oColony.ChildrenUB, oColony.lChildrenIdx.GetUpperBound(0))
                For X As Int32 = 0 To lUB
                    If oColony.lChildrenIdx(X) > -1 Then
                        Dim oFac As Facility = oColony.oChildren(X)
                        If oFac Is Nothing = False Then
                            'Now, go thru the hangar...
                            If (oFac.CurrentStatus And (elUnitStatus.eFacilityPowered Or elUnitStatus.eHangarOperational)) = (elUnitStatus.eFacilityPowered Or elUnitStatus.eHangarOperational) Then
                                Dim lHUB As Int32 = -1
                                If oFac.lHangarIdx Is Nothing = False Then lHUB = Math.Min(oFac.lHangarUB, oFac.lHangarIdx.GetUpperBound(0))
                                For Y As Int32 = 0 To oFac.lHangarUB
                                    If oFac.lHangarIdx(Y) > -1 Then
                                        Dim oUnit As Unit = CType(oFac.oHangarContents(Y), Unit)
                                        If oUnit Is Nothing = False Then
                                            If oUnit.EntityDef.Cargo_Cap > 0 AndAlso (oUnit.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                                                'ok, let's add it
                                                lUnitUB += 1
                                                If lUnitUB > oUnits.GetUpperBound(0) Then
                                                    ReDim Preserve oUnits(lUnitUB + 100)
                                                End If
                                                oUnits(lUnitUB) = oUnit
                                            End If
                                        End If
                                    End If
                                Next Y
                            End If
                        End If
                    End If
                Next X

                'Now, we have our full list... let's create our response
                Dim yResp(5 + ((lUnitUB + 1) * 28)) As Byte
                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestColonyTransportPotentials).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(lUnitUB + 1).CopyTo(yResp, lPos) : lPos += 4
                For X As Int32 = 0 To lUnitUB
                    If oUnits(X) Is Nothing = False Then
                        With oUnits(X)
                            System.BitConverter.GetBytes(.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                            .EntityName.CopyTo(yResp, lPos) : lPos += 20
                            System.BitConverter.GetBytes(.EntityDef.Cargo_Cap).CopyTo(yResp, lPos) : lPos += 4
                        End With
                    End If
                Next X
                moClients(lIndex).SendData(yResp)
            End If
        End If
    End Sub
    Private Sub HandleRequestContactWithGuild(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim oGuild As Guild = GetEpicaGuild(lGuildID)
        If oGuild Is Nothing = False Then
            'Contact someone online about this player wanting to get in touch with someone
            oGuild.SendMsgToGuildMembers(CreateChatMsg(-1, oPlayer.sPlayerNameProper & " is requesting contact with someone within the guild.", ChatMessageType.eSysAdminMessage))
        End If
    End Sub
	Private Sub HandleRequestDXDiag(ByRef yData() As Byte, ByVal lIndex As Int32)
		Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

		If oPlayer Is Nothing Then Return
		If oPlayer.bDXDiagRequested = False Then
			LogEvent(LogEventType.PossibleCheat, "HandleRequestDXDiag was unrequested: " & lPlayerID)
			Return
		End If

		'now, let's add our stuff
		Dim lPos As Int32 = 2	'for msgcode
		Dim lPack As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lTotalPacks As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim lMsgLen As Int32 = yData.Length - lPos
		Dim sValue As String = GetStringFromBytes(yData, lPos, lMsgLen)
		If lPack = 1 Then
			oPlayer.sDXDiag = sValue
		Else : oPlayer.sDXDiag &= sValue
		End If
		If lPack = lTotalPacks Then oPlayer.sDXDiag = "READY" & oPlayer.sDXDiag

		'attempt to save the player now
		Try
			oPlayer.SaveDXDiag()
		Catch
		End Try
	End Sub
    Private Sub HandleRequestEmail(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eViewEmail) = 0 Then
                Return
            End If
        End If
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        If oPlayer.InMyDomain = False Then
            'need to route this request thru to the operator
            SendPassThruMsg(yData, lPlayerID, oPlayer.ObjectID, oPlayer.ObjTypeID)
            Return
        End If

        Dim yResp() As Byte = DoRequestEmail(oPlayer, yData, 0)
        If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)

    End Sub
    Private Sub HandleRequestEmailSummarys(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eViewEmail) = 0 Then
                Return
            End If
        End If
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim yCache(200000) As Byte
        Dim yFinal() As Byte
        Dim lPos As Int32
        Dim lSingleMsgLen As Int32
        Dim yTemp() As Byte

        lPos = 0
        lSingleMsgLen = -1
        Dim lNewEmailCnt As Int32 = 0
        For Y As Int32 = 0 To oPlayer.EmailFolderUB
            If oPlayer.EmailFolderIdx(Y) <> -1 Then
                With oPlayer.EmailFolders(Y)
                    For Z As Int32 = 0 To .PlayerMsgsUB
                        If .PlayerMsgsIdx(Z) > 0 Then

                            Dim lTempLen As Int32 = 37
                            Dim lTempPos As Int32 = 0

                            If .PlayerMsgs(Z).MsgTitle Is Nothing = False Then
                                lTempLen += .PlayerMsgs(Z).MsgTitle.Length
                            End If
                            ReDim yTemp(lTempLen - 1)

                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestEmailSummarys).CopyTo(yTemp, lTempPos) : lTempPos += 2
                            .PlayerMsgs(Z).GetGUIDAsString.CopyTo(yTemp, lTempPos) : lTempPos += 6

                            If .PlayerMsgs(Z).oSender Is Nothing Then .PlayerMsgs(Z).SentByID = .PlayerMsgs(Z).PlayerID

                            .PlayerMsgs(Z).oSender.PlayerName.CopyTo(yTemp, lTempPos) : lTempPos += 20
                            yTemp(lTempPos) = .PlayerMsgs(Z).MsgRead : lTempPos += 1
                            System.BitConverter.GetBytes(.PCF_ID).CopyTo(yTemp, lTempPos) : lTempPos += 4
                            If .PlayerMsgs(Z).MsgTitle Is Nothing = False Then
                                System.BitConverter.GetBytes(.PlayerMsgs(Z).MsgTitle.Length).CopyTo(yTemp, lTempPos) : lTempPos += 4
                                .PlayerMsgs(Z).MsgTitle.CopyTo(yTemp, lTempPos) : lTempPos += .PlayerMsgs(Z).MsgTitle.Length
                            Else : System.BitConverter.GetBytes(0I).CopyTo(yTemp, lTempPos) : lTempPos += 4
                            End If

                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    Next Z
                End With
            End If
        Next Y

        If lPos <> 0 Then
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            moClients(lIndex).SendLenAppendedData(yFinal)
        End If

    End Sub
    Private Sub HandleRequestEntityAmmo(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim oEntity As Epica_Entity = Nothing

        If iObjTypeID = ObjectType.eUnit Then
            oEntity = GetEpicaUnit(lObjID)
        ElseIf iObjTypeID = ObjectType.eFacility Then
            oEntity = GetEpicaFacility(lObjID)
        End If

        If oEntity Is Nothing = False AndAlso oEntity.Owner.ObjectID = mlClientPlayer(lIndex) OrElse (oEntity.Owner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewUnitsAndFacilities) <> 0) Then
            'Ok, got it... make sure the parent is a unit or facility
            With CType(oEntity.ParentObject, Epica_GUID)
                If .ObjTypeID <> ObjectType.eUnit AndAlso .ObjTypeID <> ObjectType.eFacility Then Return
            End With

            'Dim yResp() As Byte = oEntity.GetRequestEntityAmmoResponse()
            'If yResp Is Nothing = False Then
            '	moClients(lIndex).SendData(yResp)
            'End If
        End If
    End Sub
    Private Sub HandleRequestEntityContents(ByRef yData() As Byte, ByRef oSocket As NetSock)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = Nothing
        If iObjTypeID = ObjectType.eUnit Then
            oEntity = GetEpicaUnit(lObjID)
        Else : oEntity = GetEpicaFacility(lObjID)
        End If
        If oEntity Is Nothing Then Return

        'Now, figure out how many items we'll be sending
        Try
            Dim lCnt As Int32 = 0
            If (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                For X As Int32 = 0 To oEntity.lCargoUB
                    If oEntity.lCargoIdx(X) <> -1 Then
                        Dim bGood As Boolean = True
                        If oEntity.oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
                            If CType(oEntity.oCargoContents(X), MineralCache).Quantity < 1 Then bGood = False
                        End If
                        If bGood = True Then lCnt += 1
                    End If
                Next X
            End If
            If (oEntity.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                For X As Int32 = 0 To oEntity.lHangarUB
                    If oEntity.lHangarIdx(X) <> -1 Then lCnt += 1
                Next X
            End If

            Dim yResp(29 + (lCnt * 17)) As Byte
            Dim lPos As Int32 = 0

            System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityContents).CopyTo(yResp, lPos) : lPos += 2
            oEntity.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
            If iObjTypeID = ObjectType.eUnit Then
                System.BitConverter.GetBytes(CType(oEntity, Unit).EntityDef.ObjectID).CopyTo(yResp, lPos) : lPos += 4

                System.BitConverter.GetBytes(CType(oEntity, Unit).EntityDef.Cargo_Cap).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CType(oEntity, Unit).EntityDef.Hangar_Cap).CopyTo(yResp, lPos) : lPos += 4
            Else
                System.BitConverter.GetBytes(CType(oEntity, Facility).EntityDef.ObjectID).CopyTo(yResp, lPos) : lPos += 4

                System.BitConverter.GetBytes(CType(oEntity, Facility).EntityDef.Cargo_Cap).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(CType(oEntity, Facility).EntityDef.Hangar_Cap).CopyTo(yResp, lPos) : lPos += 4
            End If
            CType(oEntity.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6

            System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lPos) : lPos += 4
            If (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                For X As Int32 = 0 To oEntity.lCargoUB
                    If oEntity.lCargoIdx(X) <> -1 Then

                        If oEntity.oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
                            If CType(oEntity.oCargoContents(X), MineralCache).Quantity < 1 Then Continue For
                        End If

                        yResp(lPos) = 1 : lPos += 1     '1 = cargo
                        oEntity.oCargoContents(X).GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6

                        Dim oObj As Epica_GUID = oEntity.oCargoContents(X)
                        If oObj.ObjTypeID = ObjectType.eMineralCache Then
                            With CType(oObj, MineralCache)
                                System.BitConverter.GetBytes(.oMineral.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                                System.BitConverter.GetBytes(ObjectType.eMineralCache).CopyTo(yResp, lPos) : lPos += 2
                                System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4
                            End With
                        ElseIf oObj.ObjTypeID = ObjectType.eUnit Then
                            CType(oObj, Unit).EntityDef.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
                            System.BitConverter.GetBytes(CInt(1)).CopyTo(yResp, lPos) : lPos += 4
                        ElseIf oObj.ObjTypeID = ObjectType.eComponentCache Then
                            With CType(oObj, ComponentCache)
                                .GetComponent.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
                                System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4
                            End With
                        ElseIf oObj.ObjTypeID = ObjectType.eColonists OrElse oObj.ObjTypeID = ObjectType.eEnlisted OrElse oObj.ObjTypeID = ObjectType.eOfficers Then
                            'Personnel...
                            oObj.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
                            System.BitConverter.GetBytes(oObj.ObjectID + 1).CopyTo(yResp, lPos) : lPos += 4
                        ElseIf oObj.ObjTypeID = ObjectType.eAmmunition Then
                            'An ammunition cache
                            With CType(oObj, AmmunitionCache)
                                System.BitConverter.GetBytes(.oWeaponTech.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                                System.BitConverter.GetBytes(ObjectType.eAmmunition).CopyTo(yResp, lPos) : lPos += 2
                                System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4
                            End With
                        End If
                    End If
                Next X
            End If
            If (oEntity.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                For X As Int32 = 0 To oEntity.lHangarUB
                    If oEntity.lHangarIdx(X) <> -1 Then
                        yResp(lPos) = 0 : lPos += 1     '0 = hangar
                        oEntity.oHangarContents(X).GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6

                        Dim oObj As Epica_GUID = oEntity.oHangarContents(X)
                        If oObj.ObjTypeID = ObjectType.eUnit Then
                            CType(oObj, Unit).EntityDef.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
                        Else : CType(oObj, Facility).EntityDef.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
                        End If

                        'TODO: Not terribly pleased with this concept here...
                        Dim lVal As Int32 = GetQueueDelay(oObj.ObjectID, oObj.ObjTypeID)
                        If lVal = -1 Then
                            'check for a LaunchAll
                            lVal = GetQueueDelay(oEntity.ObjectID, oEntity.ObjTypeID)
                        End If

                        If lVal = -1 Then
                            If oObj.ObjTypeID = ObjectType.eUnit Then
                                lVal = -(CType(oObj, Unit).EntityDef.HullSize)
                            End If
                        End If

                        'Quantity for a Hangar is the number of seconds before able to launch
                        System.BitConverter.GetBytes(lVal).CopyTo(yResp, lPos) : lPos += 4
                    End If
                Next X
            End If

            oSocket.SendData(yResp)

        Catch ex As Exception
            LogEvent(LogEventType.Warning, "HandleRequestEntityContents: " & ex.Message)
            'do nothing else in particular as the client will request the contents again
        End Try
    End Sub
    Private Sub HandleRequestEntityDefenses(ByRef yData() As Byte, ByVal lIndex As Int32)
        'GUID
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = Nothing

        If iTypeID = ObjectType.eUnit Then
            oEntity = GetEpicaUnit(lObjID)
        ElseIf iTypeID = ObjectType.eFacility Then
            oEntity = GetEpicaFacility(lObjID)
        End If

        If oEntity Is Nothing = False Then

            If oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.ParentObject Is Nothing = False Then
                Dim iTemp As Int16 = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
                If iTemp = ObjectType.ePlanet Then
                    CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
                ElseIf iTemp = ObjectType.eSolarSystem Then
                    CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
                End If
            End If

            Dim yMsg(45) As Byte
            Dim lPos As Int32 = 0

            System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityDefenses).CopyTo(yMsg, lPos) : lPos += 2

            With oEntity
                .GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                System.BitConverter.GetBytes(.Q1_HP).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.Q2_HP).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.Q3_HP).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.Q4_HP).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.Structure_HP).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.Shield_HP).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.CurrentStatus).CopyTo(yMsg, lPos) : lPos += 4
            End With

            Dim lDefCrits As Int32 = 0
            Dim oDef As Epica_Entity_Def = Nothing

            If iTypeID = ObjectType.eUnit Then
                oDef = CType(oEntity, Unit).EntityDef
            Else : oDef = CType(oEntity, Facility).EntityDef
            End If

            With oDef
                For X As Int32 = 0 To .lSideCrits.GetUpperBound(0)
                    lDefCrits = lDefCrits Or .lSideCrits(X)
                Next X

                System.BitConverter.GetBytes(lDefCrits).CopyTo(yMsg, lPos) : lPos += 4
                .GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            End With

            moClients(lIndex).SendData(yMsg)
        End If

    End Sub
    Private Sub HandleRequestEnvironmentDomain(ByVal yData() As Byte, ByRef oSocket As NetSock)
        'ok, the data has the EnvirID and the EnvirTypeID
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim oObj As Object = GetEpicaObject(lObjID, iObjTypeID)

        Dim yResp(7) As Byte

        If oObj Is Nothing = False Then
            System.BitConverter.GetBytes(GlobalMessageCode.eEnvironmentDomain).CopyTo(yResp, 0)
            If CType(oObj, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                With CType(oObj, Planet)
                    System.BitConverter.GetBytes(.oDomain.ClientListenPort).CopyTo(yResp, 2)
                    .oDomain.ClientListenIP.CopyTo(yResp, 4)
                End With
            ElseIf CType(oObj, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                With CType(oObj, SolarSystem)
                    System.BitConverter.GetBytes(.oDomain.ClientListenPort).CopyTo(yResp, 2)
                    .oDomain.ClientListenIP.CopyTo(yResp, 4)
                End With

            End If

            oSocket.SendData(yResp)
            Erase yResp
            oObj = Nothing
        End If

    End Sub
    Private Sub HandleRequestGuildEvents(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To oPlayer.oGuild.lEventUB
            If oPlayer.oGuild.oEvents(X) Is Nothing = False Then
                lCnt += 1
            End If
        Next X

        Dim yMsg(5 + (lCnt * 9)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestGuildEvents).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        With oPlayer.oGuild
            For X As Int32 = 0 To .lEventUB
                If .oEvents(X) Is Nothing = False Then
                    .oEvents(X).GetSmallString().CopyTo(yMsg, lPos) : lPos += 9
                End If
            Next X
        End With
        moClients(lIndex).SendData(yMsg)
    End Sub
    Private Sub HandleRequestGuildVoteProposals(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lType As Int32 = System.BitConverter.ToInt32(yData, 2)
        'history = 1, in progress = 0

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        Dim dtNow As DateTime = Now
        With oPlayer.oGuild
            For X As Int32 = 0 To .lVoteUB
                If .oVotes(X) Is Nothing = False Then
                    'If .oVotes(X).dtVoteStarts < dtNow AndAlso .oVotes(X).dtVoteStarts.AddHours(.oVotes(X).lVoteDuration) > dtNow Then
                    moClients(lIndex).SendData(.oVotes(X).GetVoteObjectAsAdd())
                    'End If
                End If
            Next X
        End With
    End Sub
    Private Sub HandleRequestLoadAmmo(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2

        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lWpnID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yType As Byte = yData(lPos) : lPos += 1

        Dim oEntity As Epica_Entity = Nothing
        Dim oParent As Epica_Entity = Nothing

        If iTypeID = ObjectType.eUnit Then
            oEntity = GetEpicaUnit(lID)
        ElseIf iTypeID = ObjectType.eFacility Then
            oEntity = GetEpicaFacility(lID)
        End If

        If oEntity Is Nothing Then Return

        'Ok, validate that oEntity's parent is a unit or facility
        With CType(oEntity.ParentObject, Epica_GUID)
            If .ObjTypeID <> ObjectType.eUnit AndAlso .ObjTypeID <> ObjectType.eFacility Then Return
        End With
        oParent = CType(oEntity.ParentObject, Epica_Entity)

        If oEntity.Owner.ObjectID <> mlClientPlayer(lIndex) OrElse oParent.Owner.ObjectID <> mlClientPlayer(lIndex) Then
            If oEntity.Owner.ObjectID <> mlAliasedAs(lIndex) OrElse oParent.Owner.ObjectID <> mlAliasedAs(lIndex) Then
                Return
            ElseIf (mlAliasedRights(lIndex) And AliasingRights.eTransferCargo) = 0 Then
                Return
            End If
        End If

        'If oParent Is Nothing = False Then
        '	If oParent.HandleLoadAmmo(oEntity, lWpnID, lQty, yType) = True Then
        '		'Something changed
        '		moClients(lIndex).SendData(oEntity.GetRequestEntityAmmoResponse())
        '	End If
        'End If
    End Sub
    Private Sub HandleRequestMaxWpnRng(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim oEntity As Epica_Entity = GetEpicaEntity(lEntityID, iTypeID)

        If oEntity Is Nothing = False Then
            Dim lPlayerID As Int32 = oEntity.Owner.ObjectID
            If lPlayerID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = lPlayerID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewUnitsAndFacilities) <> 0) Then
                Dim yResp(11) As Byte
                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestMaxWpnRng).CopyTo(yResp, lPos) : lPos += 2
                oEntity.GetGUIDAsString().CopyTo(yResp, lPos) : lPos += 6

                Dim oDef As Epica_Entity_Def = Nothing
                Dim lMaxRng As Int32 = 0
                If oEntity.ObjTypeID = ObjectType.eUnit Then
                    oDef = CType(oEntity, Unit).EntityDef
                ElseIf oEntity.ObjTypeID = ObjectType.eFacility Then
                    oDef = CType(oEntity, Facility).EntityDef
                End If
                If oDef Is Nothing Then Return

                For X As Int32 = 0 To oDef.WeaponDefUB
                    If oDef.WeaponDefs(X) Is Nothing = False Then
                        lMaxRng = Math.Max(lMaxRng, oDef.WeaponDefs(X).oWeaponDef.Range)
                    End If
                Next X
                System.BitConverter.GetBytes(lMaxRng).CopyTo(yResp, lPos) : lPos += 4
                moClients(lIndex).SendData(yResp)
            End If
        End If
    End Sub
    Private Sub HandleRequestMineral(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lMinID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewMining Or AliasingRights.eViewTechDesigns)) = 0 Then Return
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        For X As Int32 = 0 To oPlayer.lPlayerMineralUB
            If oPlayer.oPlayerMinerals(X).lMineralID = lMinID Then
                moClients(lIndex).SendData(GetAddObjectMessage(oPlayer.oPlayerMinerals(X), GlobalMessageCode.eAddObjectCommand))
                Return
            End If
        Next X

        LogEvent(LogEventType.PossibleCheat, "Player requesting details of unknown mineral. PlayerID: " & mlClientPlayer(lIndex))
    End Sub
    Private Sub HandleRequestPlanetOwnership(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlanet As Planet = GetEpicaPlanet(lID)
        If oPlanet Is Nothing = False Then
            Dim yResp(23) As Byte
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlanetOwnership).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oPlanet.OwnerID).CopyTo(yResp, lPos) : lPos += 4
            If oPlanet.ParentSystem Is Nothing = False Then System.BitConverter.GetBytes(oPlanet.ParentSystem.ObjectID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oPlanet.RingMineralID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oPlanet.RingMineralConcentration).CopyTo(yResp, lPos) : lPos += 4

            Dim lTempVal As Int32 = oPlanet.lColonysHereUB + 1
            If lTempVal < 0 Then lTempVal = 0
            If lTempVal > 255 Then lTempVal = 255
            yResp(lPos) = CByte(lTempVal) : lPos += 1
            yResp(lPos) = oPlanet.PlanetSizeID : lPos += 1

            moClients(lIndex).SendData(yResp)
        End If
    End Sub
    Private Sub HandleRequestPlayerBudget(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

        If lPlayerID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = lPlayerID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eViewBudget) <> 0) Then
            'The player can always request their budget
            'moclients(lindex)
            Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing = False Then
                If oPlayer.InMyDomain = True Then
                    moClients(lIndex).SendData(GetAddObjectMessage(oPlayer.oBudget, GlobalMessageCode.eAddObjectCommand))
                Else
                    moOperator.SendData(yData)
                End If
            End If
        Else
            LogEvent(LogEventType.PossibleCheat, "HandleRequestPlayerBudget PlayerID Mismatch: " & mlClientPlayer(lIndex))
        End If

    End Sub
    Private Sub HandleRequestPlayerDetails(ByRef oSocket As NetSock, ByVal yData() As Byte, ByVal lIndex As Int32)
        'Dim yObj() As Byte
        'Dim yFinal() As Byte
        'Dim iLen As Int16

        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lType As Int32 = System.BitConverter.ToInt32(yData, 6)

        'check whether the player ID is the same... if it is not... then we treat our responses differently
        If lID <> mlClientPlayer(lIndex) AndAlso lID <> mlAliasedAs(lIndex) Then
            LogEvent(LogEventType.PossibleCheat, "RequestPlayerDetails Owner Mismatch: " & mlClientPlayer(lIndex))
            Return
        End If

        If lIndex <> -1 Then
            If (mlClientStatusFlags(lIndex) And ClientStatusFlags.eRequestedPlayerDetails) <> 0 Then
                LogEvent(LogEventType.PossibleCheat, "Multiple Request Player Details calls. PlayerID: " & mlClientPlayer(lIndex))
                Return
                'Else : mlClientStatusFlags(lIndex) = mlClientStatusFlags(lIndex) Or ClientStatusFlags.eRequestedPlayerDetails
            End If
        End If

        Dim bAliased As Boolean = GetEpicaPlayer(mlClientPlayer(lIndex)).lAliasingPlayerID > 0

        'MSC - 11/27/07 - rewrote to handle login process better and more reliably

        Dim oPlayer As Player = GetEpicaPlayer(lID)
        If oPlayer Is Nothing = False Then

            Dim yCache(200000) As Byte
            Dim yFinal() As Byte
            Dim lPos As Int32
            Dim lSingleMsgLen As Int32
            Dim yTemp() As Byte

            'Ok, mark our in player request details. This will help us avoid sending msgs that may confuse the client...
            oPlayer.bInPlayerRequestDetails = True

            'Now, based on the type, what do we send?
            Select Case CType(lType, elPlayerDetailsType)
                Case elPlayerDetailsType.eAgents
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlAgentUB
                            If oPlayer.mlAgentIdx(Y) <> -1 AndAlso oPlayer.mlAgentID(Y) = glAgentIdx(oPlayer.mlAgentIdx(Y)) Then

                                'Ok, if the agent is captured
                                If (goAgent(oPlayer.mlAgentIdx(Y)).lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
                                    If (goAgent(oPlayer.mlAgentIdx(Y)).lAgentStatus And AgentStatus.IsDead) = 0 Then
                                        'ok, check if the capturing player has logged in in the last 7 days
                                        If goAgent(oPlayer.mlAgentIdx(Y)).lCapturedBy > -1 Then
                                            Dim oOther As Player = GetEpicaPlayer(goAgent(oPlayer.mlAgentIdx(Y)).lCapturedBy)
                                            If oOther Is Nothing OrElse oOther.LastLogin.AddDays(7) < Now Then
                                                'return the agent to the player
                                                With oOther.oSecurity
                                                    Dim lAgentID As Int32 = goAgent(oPlayer.mlAgentIdx(Y)).ObjectID
                                                    For Z As Int32 = 0 To .lCapturedAgentUB
                                                        If .lCapturedAgentIdx(Z) = lAgentID Then
                                                            .lCapturedAgentIdx(Z) = -1
                                                            .oCapturedAgents(Z) = Nothing
                                                        End If
                                                    Next Z
                                                End With

                                                With goAgent(oPlayer.mlAgentIdx(Y))
                                                    .yHealth = 100
                                                    .lCapturedBy = -1
                                                    .lCapturedOn = -1
                                                    .lPrisonTestCycles = 0
                                                    If (.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then .lAgentStatus = .lAgentStatus Xor AgentStatus.HasBeenCaptured
                                                    .lInterrogatorID = -1
                                                    .yInterrogationState = 0
                                                    goAgentEngine.CancelAgentDismissed(.ObjectID)
                                                    .ReturnHome()
                                                    AddToQueue(glCurrentCycle + 1, QueueItemType.eSaveObject, .ObjectID, .ObjTypeID, -1, -1, -1, -1, -1, -1)
                                                End With
                                            End If
                                        End If
                                    End If
                                End If

                                yTemp = GetAddObjectMessage(goAgent(oPlayer.mlAgentIdx(Y)), GlobalMessageCode.eAddObjectCommand)
                                lSingleMsgLen = yTemp.Length
                                'Ok, before we continue, check if we need to increase our cache
                                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                    ReDim Preserve yCache(yCache.Length + 200000)
                                End If
                                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                lPos += 2
                                yTemp.CopyTo(yCache, lPos)
                                lPos += lSingleMsgLen
                            End If
                        Next Y
                        For Y As Int32 = 0 To oPlayer.oSecurity.lCapturedAgentUB
                            If oPlayer.oSecurity.lCapturedAgentIdx(Y) <> -1 Then
                                Dim oAgent As Agent = oPlayer.oSecurity.oCapturedAgents(Y)
                                If oAgent Is Nothing = False Then
                                    yTemp = oAgent.GetCapturedAgentMsg()
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlAgentUB
                        '	If oPlayer.mlAgentIdx(Y) <> -1 AndAlso oPlayer.mlAgentID(Y) = glAgentIdx(oPlayer.mlAgentIdx(Y)) Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(goAgent(oPlayer.mlAgentIdx(Y)), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eAliasConfigs
                    'NOTE: This is the last requested item of the player request. we set the player requested details flag here so
                    '  that we can stop spammers/cheaters/hackers... if this message should go away, we need to move this functionality
                    '  to the last message requested again
                    'If bAliased = False Then
                    '	For Y As Int32 = 0 To oPlayer.lAliasUB
                    '		If oPlayer.lAliasIdx(Y) <> -1 Then
                    '			Dim yMsg(66) As Byte
                    '			lPos = 0
                    '			System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAliasConfig).CopyTo(yMsg, lPos) : lPos += 2
                    '			yMsg(lPos) = 254 : lPos += 1		'indicates to add to the client's alias list
                    '			oPlayer.oAliases(Y).PlayerName.CopyTo(yMsg, lPos) : lPos += 20
                    '			With oPlayer.uAliasLogin(Y)
                    '				.yUserName.CopyTo(yMsg, lPos) : lPos += 20
                    '				.yPassword.CopyTo(yMsg, lPos) : lPos += 20
                    '				System.BitConverter.GetBytes(.lRights).CopyTo(yMsg, lPos) : lPos += 4
                    '			End With
                    '			'oSocket.SendData(yMsg)
                    '			yTemp = yMsg
                    '			lSingleMsgLen = yTemp.Length
                    '			'Ok, before we continue, check if we need to increase our cache
                    '			If lPos + lSingleMsgLen + 2 > yCache.Length Then
                    '				ReDim Preserve yCache(yCache.Length + 200000)
                    '			End If
                    '			System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                    '			lPos += 2
                    '			yTemp.CopyTo(yCache, lPos)
                    '			lPos += lSingleMsgLen
                    '		End If
                    '	Next Y
                    'End If
                    'NOTE: This is the last requested item of the player request. we set the player requested details flag here so
                    '  that we can stop spammers/cheaters/hackers... if this message should go away, we need to move this functionality
                    '  to the last message requested again

                    Dim yObj() As Byte = oPlayer.GetObjAsString(True)
                    Dim iLen As Int16 = CShort(yObj.Length + 1)
                    ReDim yFinal(iLen)
                    System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yFinal, 0)
                    yObj.CopyTo(yFinal, 2)
                    If yFinal Is Nothing = False AndAlso yFinal.Length > 2 Then oSocket.SendData(yFinal)

                    If oPlayer.oGuild Is Nothing = False Then oSocket.SendData(GetAddObjectMessage(oPlayer.oGuild, GlobalMessageCode.eAddObjectCommand))

                    mlClientStatusFlags(lIndex) = mlClientStatusFlags(lIndex) Or ClientStatusFlags.eRequestedPlayerDetails
                    'Clear our flag for now... so that other messages can be sent
                    oPlayer.bInPlayerRequestDetails = False
                Case elPlayerDetailsType.eEmails
                    'Finally, send the player's msgs
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewEmail) <> 0 Then

                        lPos = 0
                        lSingleMsgLen = -1
                        Dim lNewEmailCnt As Int32 = 0
                        For Y As Int32 = 0 To oPlayer.EmailFolderUB
                            If oPlayer.EmailFolderIdx(Y) <> -1 Then
                                oPlayer.EmailFolders(Y).PlayerID = oPlayer.ObjectID
                                yTemp = oPlayer.EmailFolders(Y).GetAddFolderMsg()
                                lSingleMsgLen = yTemp.Length
                                'Ok, before we continue, check if we need to increase our cache
                                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                    ReDim Preserve yCache(yCache.Length + 200000)
                                End If
                                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                lPos += 2
                                yTemp.CopyTo(yCache, lPos)
                                lPos += lSingleMsgLen

                                With oPlayer.EmailFolders(Y)
                                    If .PCF_ID = PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF Then
                                        For Z As Int32 = 0 To .PlayerMsgsUB
                                            If .PlayerMsgsIdx(Z) > -1 AndAlso .PlayerMsgs(Z) Is Nothing = False AndAlso .PlayerMsgs(Z).MsgRead = 0 Then
                                                lNewEmailCnt += 1
                                            End If
                                        Next Z
                                    End If
                                End With
                            End If
                        Next Y
                        ReDim yTemp(6)
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerDetailsAlert).CopyTo(yTemp, 0)
                        yTemp(2) = 1
                        System.BitConverter.GetBytes(lNewEmailCnt).CopyTo(yTemp, 3)
                        lSingleMsgLen = yTemp.Length
                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                            ReDim Preserve yCache(yCache.Length + 200000)
                        End If
                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                        lPos += 2
                        yTemp.CopyTo(yCache, lPos)
                        lPos += lSingleMsgLen

                        'lPos = 0
                        'lSingleMsgLen = -1
                        'For Y As Int32 = 0 To oPlayer.EmailFolderUB
                        '	If oPlayer.EmailFolderIdx(Y) <> -1 Then
                        '		oPlayer.EmailFolders(Y).PlayerID = oPlayer.ObjectID
                        '		yTemp = oPlayer.EmailFolders(Y).GetAddFolderMsg()
                        '		lSingleMsgLen = yTemp.Length
                        '		'Ok, before we continue, check if we need to increase our cache
                        '		If lPos + lSingleMsgLen + 2 > yCache.Length Then
                        '			ReDim Preserve yCache(yCache.Length + 200000)
                        '		End If
                        '		System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                        '		lPos += 2
                        '		yTemp.CopyTo(yCache, lPos)
                        '		lPos += lSingleMsgLen

                        '		With oPlayer.EmailFolders(Y)
                        '			For Z As Int32 = 0 To .PlayerMsgsUB
                        '				If .PlayerMsgsIdx(Z) > 0 Then
                        '					yTemp = GetAddObjectMessage(.PlayerMsgs(Z), GlobalMessageCode.eAddObjectCommand)
                        '					lSingleMsgLen = yTemp.Length
                        '					'Ok, before we continue, check if we need to increase our cache
                        '					If lPos + lSingleMsgLen + 2 > yCache.Length Then
                        '						ReDim Preserve yCache(yCache.Length + 200000)
                        '					End If
                        '					System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                        '					lPos += 2
                        '					yTemp.CopyTo(yCache, lPos)
                        '					lPos += lSingleMsgLen
                        '				End If
                        '			Next Z
                        '		End With
                        '	End If
                        'Next Y


                        'For Y As Int32 = 0 To oPlayer.EmailFolderUB
                        '	If oPlayer.EmailFolderIdx(Y) <> -1 Then
                        '		'Send the folder down...
                        '		oPlayer.EmailFolders(Y).PlayerID = oPlayer.ObjectID
                        '		Dim yMsg() As Byte = oPlayer.EmailFolders(Y).GetAddFolderMsg()
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)

                        '		'Now, send the messages in the folder down...
                        '		With oPlayer.EmailFolders(Y)
                        '			For Z As Int32 = 0 To .PlayerMsgsUB
                        '				If .PlayerMsgsIdx(Z) > 0 Then
                        '					yMsg = GetAddObjectMessage(.PlayerMsgs(Z), GlobalMessageCode.eAddObjectCommand)
                        '					If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '				End If
                        '			Next Z
                        '		End With
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eFacilityDefs
                    'Now, send our  Facility Defs
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eAddProduction) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To glFacilityDefUB
                            If glFacilityDefIdx(Y) <> -1 AndAlso (goFacilityDef(Y).OwnerID = lID OrElse goFacilityDef(Y).OwnerID = 0) Then
                                If goFacilityDef(Y).oPrototype Is Nothing = False AndAlso goFacilityDef(Y).oPrototype.yArchived <> 0 Then Continue For
                                yTemp = GetAddObjectMessage(goFacilityDef(Y), GlobalMessageCode.eAddObjectCommand)
                                lSingleMsgLen = yTemp.Length
                                'Ok, before we continue, check if we need to increase our cache
                                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                    ReDim Preserve yCache(yCache.Length + 200000)
                                End If
                                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                lPos += 2
                                yTemp.CopyTo(yCache, lPos)
                                lPos += lSingleMsgLen

                                If goFacilityDef(Y).ProductionRequirements Is Nothing = False Then
                                    yTemp = GetAddObjectMessage(goFacilityDef(Y).ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y

                        'For Y As Int32 = 0 To glFacilityDefUB
                        '	If glFacilityDefIdx(Y) <> -1 Then
                        '		If goFacilityDef(Y).OwnerID = lID OrElse goFacilityDef(Y).OwnerID = 0 Then
                        '			Dim yMsg() As Byte = GetAddObjectMessage(goFacilityDef(Y), GlobalMessageCode.eAddObjectCommand)
                        '			If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '			If goFacilityDef(Y).ProductionRequirements Is Nothing = False Then
                        '				yMsg = GetAddObjectMessage(goFacilityDef(Y).ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
                        '				If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '			End If
                        '		End If
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eGlobalAlloys
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then
                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlAlloyUB
                                If goInitialPlayer.mlAlloyIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moAlloy(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moAlloy(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moAlloy(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y

                            'For Y As Int32 = 0 To goInitialPlayer.mlAlloyUB
                            '	If goInitialPlayer.mlAlloyIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moAlloy(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moAlloy(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moAlloy(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eGlobalArmor
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then
                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlArmorUB
                                If goInitialPlayer.mlArmorIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moArmor(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moArmor(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moArmor(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y
                            'For Y As Int32 = 0 To goInitialPlayer.mlArmorUB
                            '	If goInitialPlayer.mlArmorIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moArmor(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moArmor(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moArmor(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eGlobalEngines
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then

                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlEngineUB
                                If goInitialPlayer.mlEngineIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moEngine(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moEngine(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moEngine(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y
                            'For Y As Int32 = 0 To goInitialPlayer.mlEngineUB
                            '	If goInitialPlayer.mlEngineIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moEngine(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moEngine(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moEngine(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eGlobalHull
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then
                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlHullUB
                                If goInitialPlayer.mlHullIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moHull(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moHull(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moHull(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y
                            'For Y As Int32 = 0 To goInitialPlayer.mlHullUB
                            '	If goInitialPlayer.mlHullIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moHull(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moHull(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moHull(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eGlobalPrototype
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then

                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlPrototypeUB
                                If goInitialPlayer.mlPrototypeIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moPrototype(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y


                            'For Y As Int32 = 0 To goInitialPlayer.mlPrototypeUB
                            '	If goInitialPlayer.mlPrototypeIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moPrototype(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		'oSocket.SendData(GetAddObjectMessage(goInitialPlayer.moPrototype(Y).GetResearchCost, EpicaMessageCode.eAddObjectCommand))
                            '		'oSocket.SendData(GetAddObjectMessage(goInitialPlayer.moPrototype(Y).GetProductionCost, EpicaMessageCode.eAddObjectCommand))
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eGlobalRadar
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then
                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlRadarUB
                                If goInitialPlayer.mlRadarIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moRadar(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moRadar(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moRadar(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y
                            'For Y As Int32 = 0 To goInitialPlayer.mlRadarUB
                            '	If goInitialPlayer.mlRadarIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moRadar(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moRadar(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moRadar(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eGlobalShield
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then
                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlShieldUB
                                If goInitialPlayer.mlShieldIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moShield(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moShield(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moShield(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y
                            'For Y As Int32 = 0 To goInitialPlayer.mlShieldUB
                            '	If goInitialPlayer.mlShieldIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moShield(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moShield(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moShield(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eGlobalWeapon
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        If goInitialPlayer Is Nothing = False Then
                            lPos = 0
                            lSingleMsgLen = -1
                            For Y As Int32 = 0 To goInitialPlayer.mlWeaponUB
                                If goInitialPlayer.mlWeaponIdx(Y) <> -1 Then
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moWeapon(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moWeapon(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                    yTemp = GetAddObjectMessage(goInitialPlayer.moWeapon(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y
                            'For Y As Int32 = 0 To goInitialPlayer.mlWeaponUB
                            '	If goInitialPlayer.mlWeaponIdx(Y) <> -1 Then
                            '		Dim yMsg() As Byte = GetAddObjectMessage(goInitialPlayer.moWeapon(Y), GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moWeapon(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '		yMsg = GetAddObjectMessage(goInitialPlayer.moWeapon(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                            '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                            '	End If
                            'Next Y
                        End If
                    End If
                Case elPlayerDetailsType.eMineralProperties
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And (AliasingRights.eViewMining Or AliasingRights.eViewTechDesigns))) <> 0 Then
                        'Send down the known mineral properties to the player
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlMinPropertyUB
                            yTemp = GetAddObjectMessage(oPlayer.moMinProperties(Y), GlobalMessageCode.eAddObjectCommand)
                            If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                lSingleMsgLen = yTemp.Length
                                'Ok, before we continue, check if we need to increase our cache
                                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                    ReDim Preserve yCache(yCache.Length + 200000)
                                End If
                                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                lPos += 2
                                yTemp.CopyTo(yCache, lPos)
                                lPos += lSingleMsgLen
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlMinPropertyUB
                        '	Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moMinProperties(Y), GlobalMessageCode.eAddObjectCommand)
                        '	If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyAlloys
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlAlloyUB
                            If oPlayer.mlAlloyIdx(Y) <> -1 AndAlso oPlayer.moAlloy(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moAlloy(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moAlloy(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moAlloy(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlAlloyUB
                        '	If oPlayer.mlAlloyIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moAlloy(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moAlloy(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moAlloy(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyArmor
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlArmorUB
                            If oPlayer.mlArmorIdx(Y) <> -1 AndAlso oPlayer.moArmor(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moArmor(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moArmor(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moArmor(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlArmorUB
                        '	If oPlayer.mlArmorIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moArmor(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moArmor(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moArmor(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyEngines
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlEngineUB
                            If oPlayer.mlEngineIdx(Y) <> -1 AndAlso oPlayer.moEngine(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moEngine(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moEngine(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moEngine(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlEngineUB
                        '	If oPlayer.mlEngineIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moEngine(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moEngine(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moEngine(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyHull
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlHullUB
                            If oPlayer.mlHullIdx(Y) <> -1 AndAlso oPlayer.moHull(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moHull(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moHull(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moHull(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlHullUB
                        '	If oPlayer.mlHullIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moHull(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moHull(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moHull(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyPrototype
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlPrototypeUB
                            If oPlayer.mlPrototypeIdx(Y) <> -1 AndAlso oPlayer.moPrototype(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moPrototype(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moPrototype(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moPrototype(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlPrototypeUB
                        '	If oPlayer.mlPrototypeIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moPrototype(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moPrototype(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moPrototype(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyRadar
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlRadarUB
                            If oPlayer.mlRadarIdx(Y) <> -1 AndAlso oPlayer.moRadar(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moRadar(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moRadar(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moRadar(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlRadarUB
                        '	If oPlayer.mlRadarIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moRadar(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moRadar(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moRadar(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyShield
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlShieldUB
                            If oPlayer.mlShieldIdx(Y) <> -1 AndAlso oPlayer.moShield(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moShield(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moShield(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moShield(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlShieldUB
                        '	If oPlayer.mlShieldIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moShield(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moShield(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moShield(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMySpecials
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.oSpecials.mlSpecialTechUB
                            If oPlayer.oSpecials.mlSpecialTechIdx(Y) <> -1 AndAlso oPlayer.oSpecials.moSpecialTech(Y).bLinked = True Then
                                If oPlayer.oSpecials.moSpecialTech(Y).yArchived = 0 Then
                                    yTemp = GetAddObjectMessage(oPlayer.oSpecials.moSpecialTech(Y), GlobalMessageCode.eAddObjectCommand)
                                Else
                                    yTemp = GetAddObjectMessage(oPlayer.oSpecials.moSpecialTech(Y), GlobalMessageCode.eGetArchivedItems)
                                End If

                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.oSpecials.moSpecialTech(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.oSpecials.moSpecialTech(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlSpecialTechUB
                        '	If oPlayer.mlSpecialTechIdx(Y) <> -1 AndAlso oPlayer.moSpecialTech(Y).bLinked = True Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moSpecialTech(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moSpecialTech(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		'TODO: Production cost on a special tech???
                        '		yMsg = GetAddObjectMessage(oPlayer.moSpecialTech(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eMyWeapon
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And AliasingRights.eViewTechDesigns) <> 0) Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlWeaponUB
                            If oPlayer.mlWeaponIdx(Y) <> -1 AndAlso oPlayer.moWeapon(Y).yArchived = 0 Then
                                yTemp = GetAddObjectMessage(oPlayer.moWeapon(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moWeapon(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                yTemp = GetAddObjectMessage(oPlayer.moWeapon(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.mlWeaponUB
                        '	If oPlayer.mlWeaponIdx(Y) <> -1 Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.moWeapon(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moWeapon(Y).GetResearchCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '		yMsg = GetAddObjectMessage(oPlayer.moWeapon(Y).GetProductionCost, GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eModelDefs
                    ''Send down the model defs for the player
                    'Dim lMaxHull As Int32 = goPlayer(X).oSpecials.lPowerThrustLimit
                    'For Y = 0 To goModelDefs.ModelDefUB
                    '    Dim oModelDef As ModelDef = goModelDefs.GetModelDef(Y)
                    '    If oModelDef Is Nothing = False AndAlso oModelDef.TypeID = 2 OrElse oModelDef.MinHull < lMaxHull Then
                    '        oSocket.SendData(oModelDef.GetModelDefAddString())
                    '    End If
                    'Next Y
                Case elPlayerDetailsType.ePlayerIntel
                    'and player intel
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.mlPlayerIntelUB
                            If oPlayer.mlPlayerIntelIdx(Y) <> -1 Then
                                yTemp = GetAddObjectMessage(oPlayer.moPlayerIntel(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y

                        If oPlayer.moItemIntel Is Nothing = False Then
                            For Y As Int32 = 0 To Math.Min(oPlayer.mlItemIntelUB, oPlayer.moItemIntel.GetUpperBound(0))
                                If oPlayer.moItemIntel(Y) Is Nothing = False AndAlso oPlayer.moItemIntel(Y).yArchived = 0 Then
                                    yTemp = GetAddObjectMessage(oPlayer.moItemIntel(Y), GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            Next Y
                        End If

                        'Send the player tech knowledges
                        For Y As Int32 = 0 To oPlayer.mlPlayerTechKnowledgeUB
                            If oPlayer.myPlayerTechKnowledgeUsed(Y) <> 0 AndAlso oPlayer.moPlayerTechKnowledge(Y).yArchived = 0 Then
                                yTemp = oPlayer.moPlayerTechKnowledge(Y).GetAddMsg()
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                    End If
                Case elPlayerDetailsType.ePlayerMinerals
                    'Send down the minerals known to the player
                    If bAliased = False OrElse ((mlAliasedRights(lIndex) And (AliasingRights.eViewMining Or AliasingRights.eViewTechDesigns))) <> 0 Then

                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.lPlayerMineralUB
                            If oPlayer.oPlayerMinerals(Y) Is Nothing OrElse oPlayer.oPlayerMinerals(Y).Mineral Is Nothing Then Continue For
                            'If oPlayer.oPlayerMinerals(Y).bArchived = True Then Continue For
                            ReDim yTemp(29)
                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerDetailsPlayerMin).CopyTo(yTemp, 0)
                            oPlayer.oPlayerMinerals(Y).Mineral.GetGUIDAsString.CopyTo(yTemp, 2)
                            If oPlayer.oPlayerMinerals(Y).bDiscovered = True Then
                                yTemp(8) = 1
                                oPlayer.oPlayerMinerals(Y).Mineral.MineralName.CopyTo(yTemp, 9)
                            Else
                                yTemp(8) = 0
                                System.Text.ASCIIEncoding.ASCII.GetBytes("Unknown Mineral " & oPlayer.oPlayerMinerals(Y).DiscoveryNumber).CopyTo(yTemp, 9)
                            End If
                            'If oPlayer.oPlayerMinerals(Y).yArchived = True Then yTemp(29) = 1 Else yTemp(29) = 0
                            yTemp(29) = oPlayer.oPlayerMinerals(Y).yArchived

                            'yTemp = GetAddObjectMessage(oPlayer.oPlayerMinerals(Y), GlobalMessageCode.eAddObjectCommand)
                            If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                lSingleMsgLen = yTemp.Length
                                'Ok, before we continue, check if we need to increase our cache
                                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                    ReDim Preserve yCache(yCache.Length + 200000)
                                End If
                                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                lPos += 2
                                yTemp.CopyTo(yCache, lPos)
                                lPos += lSingleMsgLen
                            End If
                        Next Y

                        'lPos = 0
                        'lSingleMsgLen = -1
                        'For Y As Int32 = 0 To oPlayer.lPlayerMineralUB
                        '	yTemp = GetAddObjectMessage(oPlayer.oPlayerMinerals(Y), GlobalMessageCode.eAddObjectCommand)
                        '	If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                        '		lSingleMsgLen = yTemp.Length
                        '		'Ok, before we continue, check if we need to increase our cache
                        '		If lPos + lSingleMsgLen + 2 > yCache.Length Then
                        '			ReDim Preserve yCache(yCache.Length + 200000)
                        '		End If
                        '		System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                        '		lPos += 2
                        '		yTemp.CopyTo(yCache, lPos)
                        '		lPos += lSingleMsgLen
                        '	End If
                        'Next Y
                        'For Y As Int32 = 0 To oPlayer.lPlayerMineralUB
                        '	Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.oPlayerMinerals(Y), GlobalMessageCode.eAddObjectCommand)
                        '	If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        'Next Y
                    End If
                Case elPlayerDetailsType.ePlayerMissions
                    'TODO: Remove this and make it on a per request setup
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewAgents) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To glPlayerMissionUB
                            If glPlayerMissionIdx(Y) <> -1 AndAlso goPlayerMission(Y).oPlayer.ObjectID = oPlayer.ObjectID Then
                                If goPlayerMission(Y).yArchived <> 0 Then Continue For
                                yTemp = goPlayerMission(Y).GetAddObjectMessage
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To glPlayerMissionUB
                        '	If glPlayerMissionIdx(Y) <> -1 AndAlso goPlayerMission(Y).oPlayer.ObjectID = oPlayer.ObjectID Then
                        '		Dim yMsg() As Byte = goPlayerMission(Y).GetAddObjectMessage
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.ePlayerObject
                    Dim yObj() As Byte = oPlayer.GetObjAsString(True)
                    Dim iLen As Int16 = CShort(yObj.Length + 1)
                    ReDim yFinal(iLen)
                    System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yFinal, 0)
                    yObj.CopyTo(yFinal, 2)
                    If yFinal Is Nothing = False AndAlso yFinal.Length > 2 Then oSocket.SendData(yFinal)

                    Dim yMsg() As Byte
                    If bAliased = False Then
                        If oPlayer.PlayerIsDead = True Then
                            ReDim ymsg(2)
                            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yMsg, 0)
                            yMsg(2) = 1
                            oSocket.SendData(yMsg)
                        End If
                    End If

                    If oPlayer.lCelebrationEnds > glCurrentCycle Then
                        Dim lTemp As Int32 = oPlayer.lCelebrationEnds - glCurrentCycle
                        oPlayer.CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
                    End If
                    If oPlayer.bInNegativeCashflow = True Then oPlayer.CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eNegativeCashflow, 1)
                    If oPlayer.bInPirateSpawn = True Then oPlayer.CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eInPirateSpawn, 1)
                    oSocket.SendData(yData)
                    If oPlayer.oGuild Is Nothing = False Then oSocket.SendData(GetAddObjectMessage(oPlayer.oGuild, GlobalMessageCode.eAddObjectCommand))

                    ReDim yMsg(10)
                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerCustomTitle).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                    System.BitConverter.GetBytes(oPlayer.lCustomTitlePermission).CopyTo(yMsg, 6)
                    yMsg(10) = oPlayer.yCustomTitle
                    oSocket.SendData(yMsg)


                    Return
                Case elPlayerDetailsType.ePlayerRel
                    'Now, send our rels...
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewDiplomacy) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To oPlayer.PlayerRelUB
                            Dim oTmpRel As PlayerRel = oPlayer.GetPlayerRelByIndex(Y)
                            If oTmpRel Is Nothing = False Then
                                yTemp = GetAddPlayerRelMessage(oTmpRel)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To oPlayer.PlayerRelUB
                        '	Dim oTmpRel As PlayerRel = oPlayer.GetPlayerRelByIndex(Y)
                        '	If oTmpRel Is Nothing = False Then
                        '		Dim yMsg() As Byte = GetAddPlayerRelMessage(oTmpRel)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.ePlayerWormholes
                    lPos = 0
                    lSingleMsgLen = -1
                    For Y As Int32 = 0 To oPlayer.lWormholeUB
                        If oPlayer.oWormholes(Y) Is Nothing = False Then
                            yTemp = GetAddObjectMessage(oPlayer.oWormholes(Y), GlobalMessageCode.eAddObjectCommand)
                            If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                lSingleMsgLen = yTemp.Length
                                'Ok, before we continue, check if we need to increase our cache
                                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                    ReDim Preserve yCache(yCache.Length + 200000)
                                End If
                                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                lPos += 2
                                yTemp.CopyTo(yCache, lPos)
                                lPos += lSingleMsgLen
                            End If
                        End If
                    Next Y
                    'For Y As Int32 = 0 To oPlayer.lWormholeUB
                    '	If oPlayer.oWormholes(Y) Is Nothing = False Then
                    '		Dim yMsg() As Byte = GetAddObjectMessage(oPlayer.oWormholes(Y), GlobalMessageCode.eAddObjectCommand)
                    '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                    '	End If
                    'Next Y
                Case elPlayerDetailsType.eSpecialAttrs
                    'send special attributes msg
                    oSocket.SendData(oPlayer.oSpecials.GetSpecialAttributesMsg())

                    'We'll throw in formations here...
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eMoveUnits) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To glFormationDefUB
                            If glFormationDefIdx(Y) <> -1 AndAlso (goFormationDefs(Y).lOwnerID = lID) Then
                                yTemp = goFormationDefs(Y).GetAsAddMsg()
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                    End If

                Case elPlayerDetailsType.eTrades
                    If goGTC Is Nothing = False AndAlso (bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewTrades) <> 0) Then
                        goGTC.SendPlayerDirectTrades(oPlayer.ObjectID, oSocket)
                    End If
                Case elPlayerDetailsType.eUnitDefs
                    'Send our unit defs
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eAddProduction) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To glUnitDefUB
                            If glUnitDefIdx(Y) <> -1 AndAlso (goUnitDef(Y).OwnerID = lID OrElse goUnitDef(Y).OwnerID = 0) Then
                                If goUnitDef(Y).oPrototype Is Nothing = False AndAlso goUnitDef(Y).oPrototype.yArchived <> 0 Then Continue For
                                yTemp = GetAddObjectMessage(goUnitDef(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                                If goUnitDef(Y).ProductionRequirements Is Nothing = False Then
                                    yTemp = GetAddObjectMessage(goUnitDef(Y).ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
                                    If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                        lSingleMsgLen = yTemp.Length
                                        'Ok, before we continue, check if we need to increase our cache
                                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                            ReDim Preserve yCache(yCache.Length + 200000)
                                        End If
                                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                        lPos += 2
                                        yTemp.CopyTo(yCache, lPos)
                                        lPos += lSingleMsgLen
                                    End If
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To glUnitDefUB
                        '	If glUnitDefIdx(Y) <> -1 Then
                        '		If goUnitDef(Y).OwnerID = lID OrElse goUnitDef(Y).OwnerID = 0 Then
                        '			Dim yMsg() As Byte = GetAddObjectMessage(goUnitDef(Y), GlobalMessageCode.eAddObjectCommand)
                        '			If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '			If goUnitDef(Y).ProductionRequirements Is Nothing = False Then
                        '				yMsg = GetAddObjectMessage(goUnitDef(Y).ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
                        '				If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '			End If
                        '		End If
                        '	End If
                        'Next Y
                    End If
                Case elPlayerDetailsType.eUnitGroup
                    If bAliased = False OrElse (mlAliasedRights(lIndex) And AliasingRights.eViewBattleGroups) <> 0 Then
                        lPos = 0
                        lSingleMsgLen = -1
                        For Y As Int32 = 0 To glUnitGroupUB
                            If glUnitGroupIdx(Y) <> -1 AndAlso goUnitGroup(Y).oOwner.ObjectID = lID Then
                                yTemp = GetAddObjectMessage(goUnitGroup(Y), GlobalMessageCode.eAddObjectCommand)
                                If yTemp Is Nothing = False AndAlso yTemp.Length > 2 Then
                                    lSingleMsgLen = yTemp.Length
                                    'Ok, before we continue, check if we need to increase our cache
                                    If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                        ReDim Preserve yCache(yCache.Length + 200000)
                                    End If
                                    System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                                    lPos += 2
                                    yTemp.CopyTo(yCache, lPos)
                                    lPos += lSingleMsgLen
                                End If
                            End If
                        Next Y
                        'For Y As Int32 = 0 To glUnitGroupUB
                        '	If glUnitGroupIdx(Y) <> -1 AndAlso goUnitGroup(Y).oOwner.ObjectID = lID Then
                        '		Dim yMsg() As Byte = GetAddObjectMessage(goUnitGroup(Y), GlobalMessageCode.eAddObjectCommand)
                        '		If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oSocket.SendData(yMsg)
                        '	End If
                        'Next Y
                    End If
                Case Else
                    LogEvent(LogEventType.PossibleCheat, "Invalid Type for player request details! type: " & lType & ". PlayerID: " & mlClientPlayer(lIndex))
                    'NOTE: returning here causes the client to hang....
                    Return
            End Select

            'Now, send the initial request back to the client so that it can send us a new request (if needed)
            ' Our Player Has REquested Details flag has been set above in the last message sent to the client
            ' so the anti-hack mechanism should be in place already
            'oSocket.SendData(yData)
            lSingleMsgLen = yData.Length
            'Ok, before we continue, check if we need to increase our cache
            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                ReDim Preserve yCache(yCache.Length + 200000)
            End If
            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
            lPos += 2
            yData.CopyTo(yCache, lPos)
            lPos += lSingleMsgLen

            'Now, send it all...
            If lPos <> 0 Then
                ReDim yFinal(lPos - 1)
                Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
                oSocket.SendLenAppendedData(yFinal)
            End If

        End If



    End Sub
    Private Sub HandleRequestTransportDetails(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lTransID As Int32 = System.BitConverter.ToInt32(yData, 2)

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) <> -1 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewBattleGroups Or AliasingRights.eViewColonyStats)) <> (AliasingRights.eViewColonyStats Or AliasingRights.eViewBattleGroups) Then
                LogEvent(LogEventType.PossibleCheat, "RequestTransportDetails with missing Alias Rights: " & mlClientPlayer(lIndex))
                Return
            End If
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim oTrans As Transport = oPlayer.GetTransport(lTransID)
            If oTrans Is Nothing = False Then
                Dim yResp() As Byte = oTrans.HandleRequestTransportDetails()
                If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
            End If
        End If
    End Sub
    Private Sub HandleRequestTransportName(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) <> -1 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewBattleGroups Or AliasingRights.eViewColonyStats)) <> (AliasingRights.eViewColonyStats Or AliasingRights.eViewBattleGroups) Then
                LogEvent(LogEventType.PossibleCheat, "RequestTransportName with missing Alias Rights: " & mlClientPlayer(lIndex))
                Return
            End If
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim lTransID As Int32 = System.BitConverter.ToInt32(yData, 2)
            Dim oTrans As Transport = oPlayer.GetTransport(lTransID)
            If oTrans Is Nothing = False Then
                Dim yResp(25) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransportName).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(lTransID).CopyTo(yResp, lPos) : lPos += 4
                oTrans.UnitName.CopyTo(yResp, lPos) : lPos += 20
                moClients(lIndex).SendData(yResp)
            End If
        End If
    End Sub
    Private Sub HandleRequestTransports(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) <> -1 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewBattleGroups Or AliasingRights.eViewColonyStats)) <> (AliasingRights.eViewColonyStats Or AliasingRights.eViewBattleGroups) Then
                LogEvent(LogEventType.PossibleCheat, "RequestTransports with missing Alias Rights: " & mlClientPlayer(lIndex))
                Return
            End If
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim yResp() As Byte = oPlayer.HandleRequestTransports()
            If yResp Is Nothing Then Return
            moClients(lIndex).SendData(yResp)
        End If
    End Sub
    Private Sub HandleRequestTransportRouteDestList(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) <> -1 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewBattleGroups Or AliasingRights.eViewColonyStats)) <> (AliasingRights.eViewColonyStats Or AliasingRights.eViewBattleGroups) Then
                LogEvent(LogEventType.PossibleCheat, "RequestTransportRouteDestList with missing Alias Rights: " & mlClientPlayer(lIndex))
                Return
            End If
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim lUB As Int32 = -1
            If oPlayer.mlColonyIdx Is Nothing = False Then lUB = Math.Min(oPlayer.mlColonyUB, oPlayer.mlColonyIdx.GetUpperBound(0))
            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To lUB
                If oPlayer.mlColonyIdx(X) > -1 AndAlso oPlayer.mlColonyIdx(X) <= glColonyUB Then
                    If oPlayer.mlColonyID(X) = glColonyIdx(oPlayer.mlColonyIdx(X)) Then
                        lCnt += 1
                    End If
                End If
            Next X

            Dim yResp(5 + (lCnt * 8)) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransportRouteDestList).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lPos) : lPos += 4


            For X As Int32 = 0 To lUB
                If oPlayer.mlColonyIdx(X) > -1 AndAlso oPlayer.mlColonyIdx(X) <= glColonyUB Then
                    If oPlayer.mlColonyID(X) = glColonyIdx(oPlayer.mlColonyIdx(X)) Then
                        Dim oColony As Colony = goColony(oPlayer.mlColonyIdx(X))
                        If oColony Is Nothing = False Then
                            System.BitConverter.GetBytes(oColony.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                            Dim oPGuid As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                            Dim lSysID As Int32 = -1
                            If oPGuid Is Nothing = False Then
                                If oPGuid.ObjTypeID = ObjectType.ePlanet Then
                                    lSysID = CType(oPGuid, Planet).ParentSystem.ObjectID
                                ElseIf oPGuid.ObjTypeID = ObjectType.eFacility Then
                                    oPGuid = CType(CType(oPGuid, Facility).ParentObject, Epica_GUID)
                                    If oPGuid Is Nothing = False Then
                                        If oPGuid.ObjTypeID = ObjectType.eSolarSystem Then lSysID = oPGuid.ObjectID
                                    End If
                                End If
                            End If
                            System.BitConverter.GetBytes(lSysID).CopyTo(yResp, lPos) : lPos += 4
                        End If
                    End If
                End If
            Next X

            moClients(lIndex).SendData(yResp)
        End If
    End Sub
    Private Sub HandleRespawnSelection(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2

        'If the player is aliasing, just do nothing... shouldn't be here
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then Return

        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing Then Return

        'If oPlayer.PlayerIsDead = False Then
        '    LogEvent(LogEventType.PossibleCheat, "RespawnSelection from player that is not dead: " & mlClientPlayer(lIndex))
        '    Return
        'End If

        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lID = -1 Then
            'Ok, player is requesting the respawn system list
            '  all systems marked Respawn or Opened Spawn
            Dim yResp() As Byte = CompileRespawnList()
            If yResp Is Nothing = False Then moClients(lIndex).SendLenAppendedData(yResp)
        ElseIf lID = Int32.MinValue Then
            Dim yChoice As Byte = yData(lPos)
            If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount Then
                LogEvent(LogEventType.PossibleCheat, "Non-Active Account trying to set Spawn setting: " & mlClientPlayer(lIndex))
                Return
            End If
            If (oPlayer.ySpawnSystemSetting <> eySpawnSettings.eOfferedSpawn) Then
                LogEvent(LogEventType.PossibleCheat, "Player in wrong state trying to set Spawn setting: " & mlClientPlayer(lIndex))
                Return
            End If
            If yChoice = 1 Then
                'accepted
                oPlayer.ySpawnSystemSetting = oPlayer.ySpawnSystemSetting Or eySpawnSettings.eSpawnAccept
                'Now, we have to reset the player
                oPlayer.WipePlayer()

                oPlayer.lLastViewedEnvir = 0
                oPlayer.iLastViewedEnvirType = 0
                oPlayer.lStartedEnvirID = -1
                oPlayer.iStartedEnvirTypeID = -1
                oPlayer.bSpawnAtSameLoc = False
            ElseIf yChoice = 2 Then
                'refused
                oPlayer.ySpawnSystemSetting = oPlayer.ySpawnSystemSetting Or eySpawnSettings.eSpawnRefuse Or eySpawnSettings.eSpawnProcessOver
                oPlayer.oSpecials.PerformLinkTest()
            Else
                LogEvent(LogEventType.PossibleCheat, "Invalid choice passed (" & yChoice & ") in Spawn setting: " & mlClientPlayer(lIndex))
                Return
            End If
        Else
            'Player is indicating to spawn in the system of lID

            Dim lSuperEngineerDefID As Int32 = 28344

            Dim oSys As SolarSystem = GetEpicaSystem(lID)
            If oSys Is Nothing = False Then
                'Ok, the selected system should be marked Respawn or Opened
                If oSys.SystemType = SolarSystem.elSystemType.RespawnSystem OrElse oSys.SystemType = SolarSystem.elSystemType.UnlockedSystem Then
                    'Now, let's do it... need to generate the special space engineer at a random point within the solar system (x,y in the 1-3 mills)
                    Dim oUnitDef As Epica_Entity_Def = GetEpicaUnitDef(lSuperEngineerDefID)
                    Dim oTmp As Unit = New Unit()

                    Dim oRandom As New Random()

                    'Ok, populate our values
                    With oTmp
                        .bProducing = False

                        .CurrentProduction = Nothing
                        .CurrentSpeed = 0

                        .CurrentStatus = 0
                        For X As Int32 = 0 To oUnitDef.lSideCrits.Length - 1
                            .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(X)
                        Next X

                        .EntityDef = oUnitDef
                        oUnitDef.DefName.CopyTo(.EntityName, 0)
                        .ExpLevel = 0
                        .Fuel_Cap = oUnitDef.Fuel_Cap

                        .iCombatTactics = 514
                        .iTargetingTactics = 0
                        .LocAngle = 0

                        .LocX = oRandom.Next(1000000, 3000000)
                        .LocZ = oRandom.Next(1000000, 3000000)
                        If oRandom.Next(0, 100) > 50 Then .LocX = -.LocX
                        If oRandom.Next(0, 100) > 50 Then .LocZ = -.LocZ

                        .ObjTypeID = ObjectType.eUnit
                        .Owner = oPlayer
                        .ParentObject = oSys
                        .Q1_HP = oUnitDef.Q1_MaxHP
                        .Q2_HP = oUnitDef.Q2_MaxHP
                        .Q3_HP = oUnitDef.Q3_MaxHP
                        .Q4_HP = oUnitDef.Q4_MaxHP
                        .Shield_HP = oUnitDef.Shield_MaxHP
                        .Structure_HP = oUnitDef.Structure_MaxHP
                        .yProductionType = oUnitDef.ProductionTypeID
                        .DataChanged()
                    End With
                    If oTmp.SaveObject() = True Then
                        AddUnitToGlobalArray(oTmp)
                        oSys.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oTmp, GlobalMessageCode.eAddObjectCommand))
                    End If

                    With oPlayer
                        .lLastViewedEnvir = oSys.ObjectID
                        .iLastViewedEnvirType = oSys.ObjTypeID

                        '.lStartedEnvirID = lEnvirID
                        '.iStartedEnvirTypeID = iEnvirTypeID
                        '.lStartLocX = lStartX
                        '.lStartLocZ = lStartZ
                        '.lIronCurtainPlanet = .lStartedEnvirID

                        .DataChanged()
                    End With

                    Dim yResp(5) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yResp, 0)
                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yResp, 2)
                    oPlayer.SendPlayerMessage(yResp, False, 0)

                    oPlayer.SaveObject(False)

                Else
                    LogEvent(LogEventType.PossibleCheat, "RespawnSelection is not a valid system type: " & mlClientPlayer(lIndex))
                End If
            End If
        End If
    End Sub
    Private Sub HandleRestartPlayerTutorial(ByVal yData() As Byte, ByVal lIndex As Int32)
        'mlClientPlayer(lIndex)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False AndAlso oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
            oPlayer.WipePlayer()
            With oPlayer

                Dim lCurUB As Int32 = -1
                If glUnitIdx Is Nothing = False Then lCurUB = Math.Min(glUnitUB, glUnitIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If glUnitIdx(X) <> -1 Then
                        Dim oUnit As Unit = goUnit(X)
                        If oUnit Is Nothing Then Continue For
                        If oUnit.Owner.ObjectID = .ObjectID Then
                            Try
                                oUnit.DeleteEntity(X)
                            Catch
                            End Try
                            glUnitIdx(X) = -1
                        End If
                    End If
                Next X

                lCurUB = -1
                If glFacilityIdx Is Nothing = False Then lCurUB = Math.Min(glFacilityUB, glFacilityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If glFacilityIdx(X) <> -1 Then
                        Dim oFac As Facility = goFacility(X)
                        If oFac Is Nothing = False AndAlso oFac.Owner.ObjectID = .ObjectID Then
                            Try
                                oFac.DeleteEntity(X)
                            Catch
                            End Try
                            glFacilityIdx(X) = -1
                            RemoveLookupFacility(oFac.ObjectID, oFac.ObjTypeID)
                        End If
                    End If
                Next X

                lCurUB = -1
                If glColonyIdx Is Nothing = False Then lCurUB = Math.Min(glColonyIdx.GetUpperBound(0), glColonyUB)
                For X As Int32 = 0 To lCurUB
                    If glColonyIdx(X) <> -1 Then
                        Dim oColony As Colony = goColony(X)
                        If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = .ObjectID Then
                            Try
                                oColony.DeleteColony(Colony.ColonyLostReason.Destruction)
                            Catch
                            End Try
                            glColonyIdx(X) = -1
                        End If
                    End If
                Next X

                'remove agents belonging to me
                goAgentEngine.CancelPlayerAgency(lPlayerID)
                goGTC.CancelPlayersOrders(lPlayerID)

                'Remove our factions
                For X As Int32 = 0 To 4
                    If .lSlotID(X) > 0 Then
                        'ok, need to remove me from the faction of the player
                        Dim oSlotPlayer As Player = GetEpicaPlayer(oPlayer.lSlotID(X))
                        If oSlotPlayer Is Nothing = False Then
                            For Y As Int32 = 0 To 2
                                If oSlotPlayer.lFactionID(Y) = oPlayer.ObjectID Then
                                    oSlotPlayer.lFactionID(Y) = -1
                                    oSlotPlayer.ReverifySlots()
                                    Exit For
                                End If
                            Next Y
                        End If
                    End If
                    .lSlotID(X) = -1
                    .ySlotState(X) = 0
                Next X
                For X As Int32 = 0 To 2
                    If oPlayer.lFactionID(X) > 0 Then
                        Dim oFactionPlayer As Player = GetEpicaPlayer(oPlayer.lFactionID(X))
                        If oFactionPlayer Is Nothing = False Then
                            For Y As Int32 = 0 To 4
                                If oFactionPlayer.lSlotID(Y) = oPlayer.ObjectID Then
                                    oFactionPlayer.lSlotID(Y) = -1
                                    oFactionPlayer.ReverifySlots()
                                    Exit For
                                End If
                            Next Y
                        End If
                    End If
                    .lFactionID(X) = -1
                Next X
                .ReverifySlots()

                'Now, send a death to all domains
                Dim yDeath(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yDeath, 0)
                System.BitConverter.GetBytes(lPlayerID).CopyTo(yDeath, 2)
                yDeath(6) = 0
                yDeath(7) = 1
                BroadcastToDomains(yDeath)


                .lLastViewedEnvir = 0
                .iLastViewedEnvirType = 0
                .lStartedEnvirID = -1
                .yPlayerPhase = eyPlayerPhase.eInitialPhase
                .lTutorialStep = 1
                .DeathBudgetEndTime = -1
                .DeathBudgetBalance = 0
                .DeathBudgetFundsRemaining = 0
                .blCredits = 0
                .lStartLocX = Int32.MinValue
                .lStartLocZ = Int32.MinValue
                .blCrossPrimaryAdjustCredits = 0
                .oBudget.Reset()
                .lIronCurtainPlanet = -1
                .yIronCurtainState = eIronCurtainState.IronCurtainIsDown
                .yRespawnWithGuild = 0
                '.bForcedRestart = False
                .bInNegativeCashflow = False
                .bInPirateSpawn = False
                .lCelebrationEnds = 0
                .yCurrentDoctrine = eDoctrineOfLeadershipSetting.eNoSetting
                .PlayedTimeAtEndOfWaves = Int32.MinValue
                .PlayedTimeInTutorialOne = Int32.MinValue
                .PlayedTimeWhenFirstWave = Int32.MinValue
                '.PlayedTimeWhenTimerStarted = Int32.MinValue
                .TutorialPhaseWaves = 0
                .PhaseOneMoneyClicks = 0
                .dtLastRespawn = Now
                .bDeclaredWarOn = False
                '.bSpawnAtSameLoc = False
 
                .PlayerIsDead = False
            End With
        End If
    End Sub
    Private Sub HandleClientRequestObject(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) AndAlso mlAliasedAs(lIndex) > -1 Then lPlayerID = mlAliasedAs(lIndex)

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If iTypeID = ObjectType.eFacility Then

                Dim blBase As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

                Dim oFac As Facility = GetEpicaFacility(lObjID)
                If oFac Is Nothing = False Then
                    oFac.RecalcProduction()
                    blBase = CLng(blBase * oFac.Owner.fFactionResearchTimeMultiplier)
                    Dim lProdFactor As Int32 = oFac.mlProdPoints

                    Dim yResp(20) As Byte
                    Dim yType As Byte = yData(lPos)

                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yResp, lPos) : lPos += 2
                    oFac.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
                    yResp(lPos) = yType : lPos += 1

                    System.BitConverter.GetBytes(blBase).CopyTo(yResp, lPos) : lPos += 8
                    System.BitConverter.GetBytes(lProdFactor).CopyTo(yResp, lPos) : lPos += 4

                    moClients(lIndex).SendData(yResp)
                End If
            ElseIf iTypeID = ObjectType.eSolarSystem Then
                'Ok, client is requesting the details of the solarsystem for gal map view
                Dim oSystem As SolarSystem = GetEpicaSystem(lObjID)
                If oSystem Is Nothing Then Return

                Dim lPIDs() As Int32 = Nothing
                Dim yCnts() As Byte = Nothing
                Dim lPlayerCnt As Int32 = 0

                For X As Int32 = 0 To oSystem.mlPlanetUB
                    Dim lIdx As Int32 = oSystem.GetPlanetIdx(X)
                    If lIdx > -1 AndAlso lIdx <= glPlanetUB Then
                        Dim oP As Planet = goPlanet(lIdx)
                        If oP Is Nothing = False Then
                            If oP.OwnerID > 0 Then
                                Dim bFound As Boolean = False
                                For Y As Int32 = 0 To lPlayerCnt - 1
                                    If lPIDs(Y) = oP.OwnerID Then
                                        yCnts(Y) = CByte(yCnts(Y) + 1)
                                        bFound = True
                                        Exit For
                                    End If
                                Next Y
                                If bFound = False Then
                                    lPlayerCnt += 1
                                    ReDim Preserve lPIDs(lPlayerCnt - 1)
                                    ReDim Preserve yCnts(lPlayerCnt - 1)
                                    lPIDs(lPlayerCnt - 1) = oP.OwnerID
                                    yCnts(lPlayerCnt - 1) = 1
                                End If
                            End If
                        End If
                    End If
                Next X

                Dim yResp(12 + (lPlayerCnt * 5)) As Byte
                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yResp, lPos) : lPos += 2
                oSystem.GetGUIDAsString().CopyTo(yResp, lPos) : lPos += 6

                Dim lPCnt As Int32 = oSystem.mlPlanetUB + 1
                lPCnt = Math.Max(Math.Min(lPCnt, 255), 0)
                yResp(lPos) = CByte(lPCnt) : lPos += 1

                System.BitConverter.GetBytes(lPlayerCnt).CopyTo(yResp, lPos) : lPos += 4
                For X As Int32 = 0 To lPlayerCnt - 1
                    System.BitConverter.GetBytes(lPIDs(X)).CopyTo(yResp, lPos) : lPos += 4
                    yResp(lPos) = yCnts(X) : lPos += 1
                Next X

                moClients(lIndex).SendData(yResp)
            ElseIf iTypeID = ObjectType.eArena Then
                'Ok, player is requesting an arena object or arena list
                If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then Return 'aliased accts cannot do arenas

                If lObjID < 1 Then
                    'requesting the list
                    Arena.HandleRequestArenaList(moClients(lIndex), oPlayer)
                Else
                    'requesting a specific arena detail
                    Dim oArena As Arena = Arena.GetArena(lObjID)
                    If oArena Is Nothing Then Return
                    moClients(lIndex).SendData(oArena.GetArenaDetailMsg(oPlayer.ObjectID))
                End If
            End If
        End If
        
    End Sub
    Private Sub HandleSearchGuilds(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
        Dim lPos As Int32 = 2   'for msgcode

        Dim lSearchCriteria As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sSearchStr As String = GetStringFromBytes(yData, lPos, 50) : lPos += 50

        sSearchStr = sSearchStr.ToUpper.Trim

        Dim lCnt As Int32 = 0
        Dim lItemIdx() As Int32 = Nothing
        Dim lTotalBillboardLen As Int32 = 0

        'ok, let's search our guilds...
        Dim lCurUB As Int32 = -1
        If glGuildIdx Is Nothing = False Then lCurUB = Math.Min(glGuildUB, glGuildIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            'TODO: put in a better search function... one that takes placement into consideration as well as "buying price"
            If glGuildIdx(X) <> -1 Then
                Dim oGuild As Guild = goGuild(X)
                If oGuild Is Nothing = False Then
                    If sSearchStr = "" OrElse oGuild.sSearchableName = sSearchStr OrElse oGuild.sSearchableName.Contains(sSearchStr) = True Then
                        If (oGuild.iRecruitFlags And lSearchCriteria) = lSearchCriteria Then
                            lCnt += 1
                            ReDim Preserve lItemIdx(lCnt - 1)
                            lItemIdx(lCnt - 1) = X

                            If oGuild.yBillboard Is Nothing = False Then lTotalBillboardLen += oGuild.yBillboard.Length
                        End If
                    End If
                End If
            End If
        Next X

        'ok, send our result back
        Dim yMsg(5 + (lCnt * 68) + lTotalBillboardLen) As Byte
        lPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eSearchGuilds).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

        For X As Int32 = 0 To lCnt - 1
            Dim oGuild As Guild = goGuild(lItemIdx(X))
            If oGuild Is Nothing = False Then
                With oGuild
                    System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    .yGuildName.CopyTo(yMsg, lPos) : lPos += 50
                    System.BitConverter.GetBytes(GetDateAsNumber(.dtFormed)).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iRecruitFlags).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.lIcon).CopyTo(yMsg, lPos) : lPos += 4
                    If .yBillboard Is Nothing = False Then
                        System.BitConverter.GetBytes(.yBillboard.Length).CopyTo(yMsg, lPos) : lPos += 4
                        .yBillboard.CopyTo(yMsg, lPos) : lPos += .yBillboard.Length
                    Else
                        System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
                    End If
                End With
            End If
        Next X
        moClients(lIndex).SendData(yMsg)
    End Sub
    Private Sub HandleSendEmail(ByRef yData() As Byte, ByVal lIndex As Int32)
        'Ok, this is the same msg for SendEmail and SaveDraft...
        Dim lPos As Int32 = 0
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lPlayerID <> mlClientPlayer(lIndex) AndAlso (lPlayerID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eAlterEmail) = 0) Then
            LogEvent(LogEventType.PossibleCheat, "HandleSendEmail: PlayerID <> Socket.PlayerID")
            Return
        End If

        'Now, do our dealio
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

        If oPlayer Is Nothing = False Then
            If oPlayer.InMyDomain = False Then
                SendPassThruMsg(yData, oPlayer.ObjectID, oPlayer.ObjectID, ObjectType.ePlayer)
                Return
            End If

            Dim yResp() As Byte = DoSendEmail(oPlayer, yData, 0)
            If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
        End If

    End Sub
    Private Sub HandleSetAgentStatus(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) = 0 Then Return
        End If

        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim oAgent As Agent = GetEpicaAgent(lAgentID)

        If lPlayerID = oAgent.oOwner.ObjectID Then
            If lStatus = AgentStatus.Dismissed Then             'if the status is dismissed then we are dismissing this agent
                'dismiss the agent, verify the agent can be dismissed
                If (oAgent.lAgentStatus And AgentStatus.IsDead) = 0 AndAlso (oAgent.lAgentStatus And AgentStatus.NewRecruit) = 0 Then
                    If oAgent.lAgentStatus <> AgentStatus.NormalStatus Then
                        System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yData, 6)
                    Else
                        oAgent.lAgentStatus = AgentStatus.Dismissed
                        goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentDismissed, oAgent, Nothing, glCurrentCycle + 864000)

                        'Remove agent assignments currently in place
                        oAgent.RemoveMeFromMissions()
                    End If
                Else
                    oAgent.lAgentStatus = oAgent.lAgentStatus Or AgentStatus.Dismissed
                    oAgent.RemoveMeFromMissions()
                    goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentDismissed, oAgent, Nothing, glCurrentCycle)
                End If

            ElseIf lStatus = AgentStatus.NewRecruit Then            'if status is newrecruit then we are recruiting this agent
                'recruit the agent
                If (oAgent.lAgentStatus And AgentStatus.NewRecruit) <> 0 OrElse (oAgent.lAgentStatus And AgentStatus.Dismissed) <> 0 Then
                    If oAgent.oOwner.blCredits >= oAgent.lUpfrontCost OrElse oAgent.oOwner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                        oAgent.oOwner.blCredits -= oAgent.lUpfrontCost

                        If (oAgent.lAgentStatus And AgentStatus.Dismissed) <> 0 Then
                            'ok, remove the queue item
                            goAgentEngine.CancelAgentDismissed(oAgent.ObjectID)
                            'penalty for dismissing the agent
                            Dim lTemp As Int32 = oAgent.Loyalty - 10
                            If lTemp < 5 Then lTemp = 5
                            oAgent.Loyalty = CByte(lTemp)
                        End If

                        oAgent.dtRecruited = Now
                        oAgent.lAgentStatus = AgentStatus.NormalStatus
                        oAgent.DataChanged()
                    Else : System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yData, 6)
                    End If
                End If
            End If
            moClients(lIndex).SendData(yData)
        Else : LogEvent(LogEventType.PossibleCheat, "HandleSetAgentStatus Player Owner Mismatch: " & mlClientPlayer(lIndex))
        End If

    End Sub
    Private Sub HandleSetArchiveState(ByRef yData() As Byte, ByVal lIndex As Int32)
        If mlAliasedAs(lIndex) > 0 AndAlso mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) Then Return

        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing Then Return

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim yValue As Byte = yData(8)
        Dim bSendDeleteDesign As Boolean = False

        Select Case iObjTypeID
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHangarTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
                Dim oTech As Epica_Tech = oPlayer.GetTech(lObjID, iObjTypeID)
                If oTech Is Nothing = False Then
                    If oTech.Owner.ObjectID = oPlayer.ObjectID Then
                        'If yValue = 0 Then oTech.bArchived = False Else oTech.bArchived = True
                        oTech.yArchived = yValue
                        oTech.SaveObject()
                        If oTech.ObjTypeID = ObjectType.eAlloyTech Then
                            With CType(oTech, AlloyTech)
                                Dim oResult As Mineral = .AlloyResult
                                If oResult Is Nothing = False Then
                                    Dim oPM As PlayerMineral = Nothing
                                    For X As Int32 = 0 To oPlayer.lPlayerMineralUB
                                        If oPlayer.oPlayerMinerals(X) Is Nothing = False Then
                                            If oPlayer.oPlayerMinerals(X).lMineralID = oResult.ObjectID Then
                                                Dim yTmpVal As Byte = yValue
                                                If yTmpVal = 255 Then yTmpVal = 1
                                                oPlayer.oPlayerMinerals(X).yArchived = yTmpVal
                                                oPlayer.oPlayerMinerals(X).SaveObject(oPlayer.ObjectID)
                                                Exit For
                                            End If
                                        End If
                                    Next X
                                End If
                            End With
                        End If
                        bSendDeleteDesign = yValue = 255
                    End If
                End If
            Case ObjectType.eMineral
                For X As Int32 = 0 To oPlayer.lPlayerMineralUB
                    If oPlayer.oPlayerMinerals(X) Is Nothing = False AndAlso oPlayer.oPlayerMinerals(X).lMineralID = lObjID Then
                        'If yValue = 0 Then oPlayer.oPlayerMinerals(X).bArchived = False Else oPlayer.oPlayerMinerals(X).bArchived = True
                        oPlayer.oPlayerMinerals(X).yArchived = yValue
                        oPlayer.oPlayerMinerals(X).SaveObject(oPlayer.ObjectID)

                        If oPlayer.oPlayerMinerals(X).Mineral.lAlloyTechID > -1 Then
                            'ok, archive the alloytech too
                            Dim oTech As Epica_Tech = oPlayer.GetTech(oPlayer.oPlayerMinerals(X).Mineral.lAlloyTechID, ObjectType.eAlloyTech)
                            If oTech Is Nothing = False AndAlso oTech.Owner.ObjectID = oPlayer.ObjectID Then
                                Dim yTmpVal As Byte = yValue
                                If yTmpVal = 255 Then yTmpVal = 1
                                oTech.yArchived = yTmpVal
                                oTech.SaveObject()
                            End If
                        End If
                        Exit For
                    End If
                Next X
            Case ObjectType.eMission
                Dim oPM As PlayerMission = GetEpicaPlayerMission(lObjID)
                If oPM Is Nothing = False Then
                    If oPM.oPlayer Is Nothing = False AndAlso oPM.oPlayer.ObjectID = oPlayer.ObjectID Then
                        oPM.yArchived = yValue
                        oPM.SaveObject()
                    End If
                End If
            Case ObjectType.ePlayerItemIntel
                Dim iExt As Int16 = System.BitConverter.ToInt16(yData, 9)
                For X As Int32 = 0 To oPlayer.mlItemIntelUB
                    Dim oPII As PlayerItemIntel = oPlayer.moItemIntel(X)
                    If oPII Is Nothing = False Then
                        If oPII.lItemID = lObjID AndAlso oPII.iItemTypeID = iExt Then
                            'If yValue = 0 Then oPII.bArchived = False Else oPII.bArchived = True
                            oPII.yArchived = yValue
                            oPII.SaveObject()
                            Exit For
                        End If
                    End If
                Next X
            Case ObjectType.ePlayerTechKnowledge
                Dim iExt As Int16 = System.BitConverter.ToInt16(yData, 9)
                Dim oPTK As PlayerTechKnowledge = oPlayer.GetPlayerTechKnowledge(lObjID, iExt)
                If oPTK Is Nothing = False Then
                    If oPTK.oPlayer Is Nothing = False AndAlso oPTK.oPlayer.ObjectID = oPlayer.ObjectID Then
                        'If yValue = 0 Then oPTK.bArchived = False Else oPTK.bArchived = True
                        oPTK.yArchived = yValue
                        oPTK.SaveObject()
                    End If
                End If

            Case Else
                LogEvent(LogEventType.PossibleCheat, "HandleSetArchiveState item mismatch: " & mlClientPlayer(lIndex))
        End Select

        If bSendDeleteDesign = True Then
            Dim yResp(7) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
            moClients(lIndex).SendData(yResp)
        End If
    End Sub
    Private Sub HandleSetColonyResearchQueue(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eAddResearch) = 0 Then Return
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)
            If lColonyIdx > -1 AndAlso glColonyIdx(lColonyIdx) > -1 Then
                Dim oColony As Colony = goColony(lColonyIdx)
                If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = oPlayer.ObjectID Then

                    Dim yAction As Byte = yData(lPos) : lPos += 1
                    Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iTechTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim yCnt As Byte = yData(lPos) : lPos += 1
                    Dim yQty As Byte = yData(lPos) : lPos += 1
                    If yAction = 0 Then
                        oColony.AddColonyQueueItem(lTechID, iTechTypeID, yCnt, yQty)
                    Else
                        oColony.RemoveColonyQueueItem(lTechID, iTechTypeID)
                    End If
                    AddToQueue(glCurrentCycle + 30, QueueItemType.eCheckColonyResearchQueue, oColony.ObjectID, 0, 0, 0, 0, 0, 0, 0)

                    Dim yResp() As Byte = oColony.HandleGetColonyResearchList()
                    If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
                End If
            End If
        End If

    End Sub
    Private Sub HandleSetEntityName(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        If lObjID < 1 Then Return


        If iObjTypeID = ObjectType.eUnit OrElse iObjTypeID = ObjectType.eFacility Then
            Dim oObj As Object = GetEpicaObject(lObjID, iObjTypeID)
            If oObj Is Nothing Then Return
            With CType(oObj, Epica_Entity)
                If .Owner.ObjectID = mlClientPlayer(lIndex) Then
                    ReDim .EntityName(19)
                    Array.Copy(yData, 8, .EntityName, 0, 20)

                    .bNeedsSaved = True

                    Dim iTemp As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID
                    If iTemp = ObjectType.ePlanet Then
                        CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
                    ElseIf iTemp = ObjectType.eSolarSystem Then
                        CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
                    End If
                Else
                    LogEvent(LogEventType.PossibleCheat, "Name Change for Entity that does not belong to socket owner. EntityID: " & .ObjectID & " (of " & .ObjTypeID & "), OwnerID: " & .Owner.ObjectID & ", SocketPlayerID: " & mlClientPlayer(lIndex))
                End If
            End With
        ElseIf iObjTypeID = ObjectType.eColony Then
            Dim oObj As Object = GetEpicaObject(lObjID, iObjTypeID)
            If oObj Is Nothing Then Return
            With CType(oObj, Colony)
                If .Owner.ObjectID = mlClientPlayer(lIndex) Then
                    ReDim .ColonyName(19)
                    Array.Copy(yData, 8, .ColonyName, 0, 20)
                Else
                    LogEvent(LogEventType.PossibleCheat, "Name Change for Colony that does not belong to socket owner. ColonyID: " & .ObjectID & ", OwnerID: " & .Owner.ObjectID & ", SocketPlayerID: " & mlClientPlayer(lIndex))
                End If
            End With
        ElseIf iObjTypeID = ObjectType.eUnitGroup Then
            Dim oObj As Object = GetEpicaObject(lObjID, iObjTypeID)
            If oObj Is Nothing Then Return
            With CType(oObj, UnitGroup)
                If .oOwner.ObjectID = mlClientPlayer(lIndex) Then
                    Dim sNewVal As String = GetStringFromBytes(yData, 8, 20).ToUpper
                    For X As Int32 = 0 To glUnitGroupUB
                        If glUnitGroupIdx(X) <> -1 AndAlso goUnitGroup(X).oOwner.ObjectID = mlClientPlayer(lIndex) Then
                            If goUnitGroup(X).ObjectID <> .ObjectID AndAlso BytesToString(goUnitGroup(X).UnitGroupName).ToUpper = sNewVal Then
                                Return
                            End If
                        End If
                    Next X
                    ReDim CType(oObj, UnitGroup).UnitGroupName(19)
                    Array.Copy(yData, 8, CType(oObj, UnitGroup).UnitGroupName, 0, 20)
                Else
                    LogEvent(LogEventType.PossibleCheat, "Name Change for Unit Group that does not belong to socket owner. UnitGroupID: " & .ObjectID & ", OwnerID: " & .oOwner.ObjectID & ", SocketPlayerID: " & mlClientPlayer(lIndex))
                End If
            End With
        Else
            'Check for components
            Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
            If oPlayer.blCredits >= 10000000 Then
                oPlayer.blCredits -= 10000000
            Else : Return
            End If

            Dim oTech As Epica_Tech = oPlayer.GetTech(lObjID, iObjTypeID)
            If oTech Is Nothing = False AndAlso oTech.Owner.ObjectID = oPlayer.ObjectID Then
                Dim sNewVal As String = GetStringFromBytes(yData, 8, 20).ToUpper
                oTech.SetTechName(sNewVal)

                If iObjTypeID = ObjectType.eAlloyTech Then
                    If CType(oTech, AlloyTech).AlloyResult Is Nothing = False Then
                        CType(oTech, AlloyTech).AlloyResult.MineralName = oTech.GetTechName
                        CType(oTech, AlloyTech).AlloyResult.SaveObject()
                    End If
                End If
                oTech.SaveObject()
            End If
        End If

    End Sub
    Private Sub HandleSetEntityPersonnel(ByVal yData() As Byte, ByVal lClientIndex As Int32)
        'Ok, from the client, it should have the GUID of the facility to return...
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        'We don't really care too much about the typeid
        'Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim oFac As Facility = GetEpicaFacility(lID)

        Dim yAutoLaunch As Byte = 2
        If yData.GetUpperBound(0) >= 8 Then
            yAutoLaunch = yData(8)
        End If

        If oFac Is Nothing = False Then
            Dim yResp(20) As Byte
            With oFac
                If yAutoLaunch = 1 Then
                    .AutoLaunch = True
                ElseIf yAutoLaunch = 0 Then
                    If .yProductionType <> ProductionType.eMining Then .AutoLaunch = False
                End If

                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityPersonnel).CopyTo(yResp, 0)
                .GetGUIDAsString.CopyTo(yResp, 2)

                'Pos 8 is the Max Colonists...
                If .yProductionType = ProductionType.eColonists OrElse .yProductionType = ProductionType.eCommandCenterSpecial Then
                    System.BitConverter.GetBytes(.EntityDef.ProdFactor).CopyTo(yResp, 8)
                Else : System.BitConverter.GetBytes(CInt(0)).CopyTo(yResp, 8)
                End If

                'Pos 12 is the Max Workers...
                System.BitConverter.GetBytes(.MaxWorkers).CopyTo(yResp, 12)

                System.BitConverter.GetBytes(.PowerConsumption).CopyTo(yResp, 16)

                If .AutoLaunch = True Then yResp(20) = 1 Else yResp(20) = 0
            End With
            moClients(lClientIndex).SendData(yResp)
        End If
    End Sub
    Private Sub HandleSetFleetDest(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, 6)

        Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lFleetID)
        If oUnitGroup Is Nothing = False Then
            If oUnitGroup.SetDestSystem(lSystemID, False) = False Then
                'Alert the player unable to move there
                Dim yMsg(9) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetDest).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(lFleetID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(CInt(-1)).CopyTo(yMsg, 6)
                moClients(lIndex).SendData(yMsg)
            Else
                'Alert the player move order accepted
                oUnitGroup.SendPlayerSetFleetDestMsg()
            End If
        End If
    End Sub
    Private Sub HandleSetFleetFormation(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lUnitGroupID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFormationID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lUnitGroupID)
        If oUnitGroup Is Nothing = False Then
            If oUnitGroup.oOwner.ObjectID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = oUnitGroup.oOwner.ObjectID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eModifyBattleGroups) <> 0) Then
                If lFormationID < 1 Then
                    oUnitGroup.lDefaultFormationID = lFormationID
                Else
                    Dim oFormation As FormationDef = GetEpicaFormation(lFormationID)
                    If oFormation Is Nothing = False Then
                        If oFormation.lOwnerID <> oUnitGroup.oOwner.ObjectID Then
                            LogEvent(LogEventType.PossibleCheat, "HandleSetFleetFormation player using another player's formation: " & mlClientPlayer(lIndex))
                            Return
                        End If
                    Else : lFormationID = -1
                    End If
                    oUnitGroup.lDefaultFormationID = lFormationID
                End If
            Else
                LogEvent(LogEventType.PossibleCheat, "HandleSetFleetFormation player without rights: " & mlClientPlayer(lIndex))
                Return
            End If
        End If
    End Sub
    Private Sub HandleSetFleetReinforcer(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'msgcode
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFacilityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oFleet As UnitGroup = GetEpicaUnitGroup(lFleetID)
        Dim yResp() As Byte

        If oFleet Is Nothing = False Then
            If oFleet.oOwner.ObjectID <> mlClientPlayer(lIndex) AndAlso (oFleet.oOwner.ObjectID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eModifyBattleGroups) = 0) Then
                LogEvent(LogEventType.PossibleCheat, "HandleSetFleetReinforcer Owner mismatch: " & mlClientPlayer(lIndex))
                Return
            End If

            Dim bRemove As Boolean = lFacilityID < 0
            lFacilityID = Math.Abs(lFacilityID)

            Dim oFac As Facility = GetEpicaFacility(lFacilityID)
            If oFac Is Nothing = False Then
                If oFac.Owner.ObjectID <> oFleet.oOwner.ObjectID Then
                    LogEvent(LogEventType.PossibleCheat, "HandleSetFleetReinforcer Fac Owner Mismatches Fleet Owner: " & mlClientPlayer(lIndex))
                    Return
                End If

                If bRemove = False Then
                    oFleet.AddReinforcer(oFac.ObjectID)
                    oFac.lReinforcingUnitGroupID = oFleet.ObjectID

                    ReDim yResp(15)
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetReinforcer).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lFleetID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lFacilityID).CopyTo(yResp, lPos) : lPos += 4
                    With CType(oFac.ParentObject, Epica_GUID)
                        .GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
                    End With
                Else
                    oFleet.RemoveReinforcer(oFac.ObjectID)
                    Return
                End If
            Else
                ReDim yResp(9)
                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetReinforcer).CopyTo(yResp, lPos) : lPos += 2
                lFacilityID = -1
                System.BitConverter.GetBytes(lFleetID).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(lFacilityID).CopyTo(yResp, lPos) : lPos += 4
            End If
        Else
            ReDim yResp(9)
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetReinforcer).CopyTo(yResp, lPos) : lPos += 2
            lFleetID = -1
            lFacilityID = -1
            System.BitConverter.GetBytes(lFleetID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lFacilityID).CopyTo(yResp, lPos) : lPos += 4
        End If

        moClients(lIndex).SendData(yResp)
    End Sub
    Private Sub HandleSetGuildRel(ByRef yData() As Byte, ByVal lIndex As Int32)
        'Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        'If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
        'Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        'If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        'If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ModifyGuildRelation) = False Then
        '	LogEvent(LogEventType.PossibleCheat, "Player attempting to set relationship without permission: " & mlClientPlayer(lIndex))
        '	Return
        'End If

        'Dim lPos As Int32 = 2	 'for msgcode
        'Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'Dim yRel As Byte = yData(lPos) : lPos += 1

        'Dim oRel As GuildRel = oPlayer.oGuild.GetRel(lEntityID, iEntityTypeID)
        'If oRel Is Nothing = False Then
        '	oRel.yRelTowardsThem = yRel

        '	'TODO: Send to domain servers and to the player/guild that it relates to (however, reverse it or something)

        '	oPlayer.oGuild.SendMsgToGuildMembers(yData)
        'End If
    End Sub
    Private Sub HandleSetInfiltrateSettings(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oAgent As Agent = GetEpicaAgent(lAgentID)

        If oAgent Is Nothing Then Return
        Dim lPlayerID As Int32 = oAgent.oOwner.ObjectID

        If lPlayerID = mlClientPlayer(lIndex) OrElse (mlAliasedAs(lIndex) = lPlayerID AndAlso (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) <> 0) Then
            Dim lInfType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lInfTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iInfTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lFreq As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'Ok, now, is the agent already infiltrated?
            If (oAgent.lAgentStatus And (AgentStatus.IsInfiltrated Or AgentStatus.Infiltrating Or AgentStatus.HasBeenCaptured Or AgentStatus.IsDead)) <> 0 Then
                '  if so, is the infiltration targets the same?
                If oAgent.lTargetID <> lInfTargetID OrElse oAgent.iTargetTypeID <> iInfTargetTypeID Then
                    If (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead)) = 0 AndAlso (oAgent.lTargetID = -1 OrElse oAgent.iTargetTypeID = -1) Then
                        If (oAgent.lAgentStatus And AgentStatus.IsInfiltrated) <> 0 Then oAgent.lAgentStatus = oAgent.lAgentStatus Xor AgentStatus.IsInfiltrated
                        If (oAgent.lAgentStatus And AgentStatus.Infiltrating) <> 0 Then
                            oAgent.lAgentStatus = oAgent.lAgentStatus Xor AgentStatus.Infiltrating
                            goAgentEngine.CancelAgentEvent(AgentEngine.EventTypeID.AgentFirstInfiltrate, oAgent.ObjectID)
                        End If
                    Else
                        '  if not, return invalid target
                        '2 for msgcode, 4 for agentid
                        System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yData, 6)
                        moClients(lIndex).SendData(yData)
                        Return
                    End If
                End If
            End If

            'set report freq
            If lFreq < 0 Then lFreq = 0
            If lFreq > 255 Then lFreq = 255
            oAgent.yReportFreq = CByte(lFreq)

            If oAgent.InfiltrationType <> lInfType AndAlso oAgent.InfiltrationLevel > 50 Then
                oAgent.InfiltrationLevel = 50
                goAgentEngine.CancelAgentEvent(AgentEngine.EventTypeID.AgentCheckIn, oAgent.ObjectID)
                goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentCheckIn, oAgent, Nothing, glCurrentCycle + Agent.GetCycleDelayByFreq(oAgent.yReportFreq))
            End If
            oAgent.InfiltrationType = CType(lInfType, eInfiltrationType)

            Dim bAdd As Boolean = (oAgent.lAgentStatus And (AgentStatus.IsInfiltrated Or AgentStatus.Infiltrating)) = 0

            'otherwise, begin infiltration process (if needed)
            If oAgent.lTargetID <> lInfTargetID OrElse oAgent.iTargetTypeID <> iInfTargetTypeID Then
                If (oAgent.lAgentStatus And AgentStatus.IsInfiltrated) = 0 Then oAgent.lAgentStatus = oAgent.lAgentStatus Or AgentStatus.Infiltrating
                oAgent.lTargetID = lInfTargetID
                oAgent.iTargetTypeID = iInfTargetTypeID
            End If

            If bAdd = True Then
                Dim lCycle As Int32 = glCurrentCycle + Agent.ml_DEINFILTRATE_TIME
                If oAgent.iTargetTypeID = ObjectType.ePlayer AndAlso oAgent.lTargetID = oAgent.oOwner.ObjectID Then lCycle = glCurrentCycle
                If (oAgent.lAgentStatus And AgentStatus.IsInfiltrated) = 0 Then
                    If oAgent.oOwner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                        lCycle = glCurrentCycle + 300
                        oAgent.lArrivalCycles = 300
                    Else
                        oAgent.lArrivalCycles = Agent.ml_DEINFILTRATE_TIME
                    End If
                    oAgent.lArrivalStart = glCurrentCycle
                    ReDim Preserve yData(yData.Length + 4)
                    System.BitConverter.GetBytes(oAgent.lArrivalCycles).CopyTo(yData, lPos) : lPos += 4
                    goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentFirstInfiltrate, oAgent, Nothing, lCycle)
                Else
                    goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.AgentReInfiltrate, oAgent, Nothing, lCycle)
                End If
            End If

            'send response to client
            moClients(lIndex).SendData(yData)

        Else : LogEvent(LogEventType.PossibleCheat, "HandleSetInfiltrateSettings Player Owner Mismatch: " & mlClientPlayer(lIndex))
        End If
    End Sub
    Private Sub HandleSetIronCurtain(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yValue As Byte = yData(lPos) : lPos += 1

        If mlClientPlayer(lIndex) <> lPlayerID OrElse (mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) <> -1) Then
            LogEvent(LogEventType.PossibleCheat, "HandleSetIronCurtain msg received from aliased player. " & mlClientPlayer(lIndex))
            Return
        End If

        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing = False Then
            If lPlanetID = Int32.MinValue Then
                If yValue = 0 Then
                    If (oPlayer.lStatusFlags And elPlayerStatusFlag.AlwaysRaiseFullInvul) <> 0 Then
                        oPlayer.lStatusFlags = oPlayer.lStatusFlags Xor elPlayerStatusFlag.AlwaysRaiseFullInvul
                    End If
                Else
                    oPlayer.lStatusFlags = oPlayer.lStatusFlags Or elPlayerStatusFlag.AlwaysRaiseFullInvul
                    If oPlayer.yIronCurtainState = eIronCurtainState.IronCurtainIsUpOnSelectedPlanet Then
                        AddToQueue(glCurrentCycle + 1, QueueItemType.eIronCurtainRaise, lPlayerID, 0, 0, 0, 0, 0, 0, 0)
                    End If
                End If
            Else
                If oPlayer.oBudget Is Nothing = False Then
                    If oPlayer.oBudget.IsInEnvironment(lPlanetID, ObjectType.ePlanet) = False Then
                        LogEvent(LogEventType.PossibleCheat, "HandleSetIronCurtain references invalid planet target (not in budget): " & mlClientPlayer(lIndex))
                        Return
                    End If
                    oPlayer.lIronCurtainPlanet = lPlanetID
                    moClients(lIndex).SendData(GetAddObjectMessage(oPlayer.oBudget, GlobalMessageCode.eAddObjectCommand))
                End If
            End If

        End If
    End Sub
    Private Sub HandleSetMineralBid(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewMining Or AliasingRights.eAddProduction)) = 0 Then
                LogEvent(LogEventType.PossibleCheat, "HandleSetMineralBid: Player does not have permission to set value: " & mlClientPlayer(lIndex))
                Return
            End If
        End If

        Dim lPos As Int32 = 2
        Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFacID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lBidAmt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMaxQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oFac As Facility = GetEpicaFacility(lFacID)

        If oFac Is Nothing = False Then
            If oFac.lCacheID <> lCacheID Then
                LogEvent(LogEventType.PossibleCheat, "HandleSetMineralBid: Facility to CacheID mismatch: " & mlClientPlayer(lIndex))
                Return
            End If

            If oFac.oMiningBid Is Nothing = True Then
                oFac.oMiningBid = New MineBuyOrderManager()
                oFac.oMiningBid.oMineralCache = GetEpicaMineralCache(oFac.lCacheID)
                oFac.oMiningBid.oParentFacility = oFac
            End If

            Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
            oFac.oMiningBid.AddBid(lBidAmt, lMaxQty, oPlayer, True)

            Dim yMsg() As Byte = oFac.oMiningBid.GetBidAsMsg(lPlayerID)
            moClients(lIndex).SendData(yMsg)
        End If

    End Sub
    Private Sub HandleSetSkipStatus(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPM_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lGoalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) = 0 Then Return
        End If
        Dim oPM As PlayerMission = GetEpicaPlayerMission(lPM_ID)
        If oPM Is Nothing Then Return
        If oPM.oPlayer.ObjectID <> lPlayerID Then
            LogEvent(LogEventType.PossibleCheat, "HandleSetSkipStatus: Player mismatch. Player: " & mlClientPlayer(lIndex))
            Return
        End If

        If lGoalID > 0 Then
            Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lSkillID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim ySkipVal As Byte = yData(lPos) : lPos += 1

            If oPM.lCurrentPhase > eMissionPhase.ePreparationTime Then Return
            If oPM.oMission Is Nothing = False Then
                Dim oPMG As PlayerMissionGoal = Nothing
                For X As Int32 = 0 To oPM.oMission.GoalUB
                    If oPM.oMission.MethodIDs(X) = oPM.lMethodID AndAlso oPM.oMission.Goals(X).ObjectID = lGoalID Then
                        oPMG = oPM.oMissionGoals(X)
                        Exit For
                    End If
                Next X
                If oPMG Is Nothing Then Return
                If oPMG.oGoal.MissionPhase = eMissionPhase.ePreparationTime Then
                    For X As Int32 = 0 To oPMG.lAssignmentUB
                        If oPMG.oAssignments(X) Is Nothing = False Then
                            If oPMG.oAssignments(X).oAgent.ObjectID = lAgentID AndAlso oPMG.oAssignments(X).oSkill.ObjectID = lSkillID Then
                                If ySkipVal = 0 Then
                                    'clearing skipped
                                    If (oPMG.oAssignments(X).yStatus And AgentAssignmentStatus.Skipped) <> 0 Then
                                        oPMG.oAssignments(X).yStatus = oPMG.oAssignments(X).yStatus Or AgentAssignmentStatus.Skipped
                                    End If
                                Else
                                    'setting skipped
                                    oPMG.oAssignments(X).yStatus = oPMG.oAssignments(X).yStatus Or AgentAssignmentStatus.Skipped
                                End If

                                'send our message back to the client so it can update itself
                                moClients(lIndex).SendData(yData)

                                Exit For
                            End If
                        End If
                    Next X
                Else
                    LogEvent(LogEventType.PossibleCheat, "HandleSetSkipStatus: Goal Phase Mismatch: " & mlClientPlayer(lIndex))
                End If
            End If
        Else
            'Ok, special command for the mission: Cancel, Pause or Resume Mission
            If lGoalID = -1 Then
                'cancel mission...
                Dim lTemp As Int32 = oPM.lCurrentPhase
                If (oPM.lCurrentPhase And eMissionPhase.eMissionPaused) <> 0 Then
                    lTemp = lTemp Xor eMissionPhase.eMissionPaused
                End If
 
                goAgentEngine.CancelMissionEvent(AgentEngine.EventTypeID.ProcessMission, oPM.PM_ID)
                If lTemp > eMissionPhase.ePreparationTime Then
                    If oPM.lCurrentPhase <> eMissionPhase.eMissionOverFailure AndAlso oPM.lCurrentPhase <> eMissionPhase.eMissionOverSuccess Then
                        oPM.lCurrentPhase = eMissionPhase.eReinfiltrationPhase
                        oPM.ProcessTests()
                    End If
                End If
                oPM.lCurrentPhase = eMissionPhase.eCancelled
                oPM.MarkAgentsOffMission()
                moClients(lIndex).SendData(yData)
            ElseIf lGoalID = -2 Then
                'pause/resume
                If (oPM.lCurrentPhase And eMissionPhase.eMissionPaused) = 0 Then
                    'ok, mission is not paused
                    '  remove the event from the agent engine (if one exists)
                    goAgentEngine.CancelMissionEvent(AgentEngine.EventTypeID.ProcessMission, oPM.PM_ID)

                    '  set our mission status
                    oPM.lCurrentPhase = oPM.lCurrentPhase Or eMissionPhase.eMissionPaused
                Else
                    'mission is paused
                    goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ProcessMission, Nothing, oPM, glCurrentCycle + AgentEngine.ml_MISSION_CHECK_DELAY)
                    'clear our paused status
                    oPM.lCurrentPhase = oPM.lCurrentPhase Xor eMissionPhase.eMissionPaused
                End If
                ReDim Preserve yData(lPos)
                yData(lPos) = oPM.lCurrentPhase
                moClients(lIndex).SendData(yData)
            End If
        End If

    End Sub
    Private Sub HandleSetSpecialTechThrowback(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2

        'Validate that the player belongs to the socket and that the researcher's owner is the player
        If (mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eAddResearch) <> 0 Then
            LogEvent(LogEventType.PossibleCheat, "Player without rights to SetSpecTechThrowback: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then lPlayerID = mlAliasedAs(lIndex)

        Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTechTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim yValue As Byte = yData(lPos) : lPos += 1

        If iTechTypeID <> ObjectType.eSpecialTech Then
            LogEvent(LogEventType.PossibleCheat, "Player trying to throwback non spec tech: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim oTech As PlayerSpecialTech = CType(oPlayer.GetTech(lTechID, iTechTypeID), PlayerSpecialTech)
        If oTech Is Nothing Then
            LogEvent(LogEventType.PossibleCheat, "Player trying to throwback tech they dont have: " & mlClientPlayer(lIndex))
            Return
        End If

        If oTech.bLinked = False Then
            LogEvent(LogEventType.PossibleCheat, "Player trying to throwback unlinked spec tech: " & mlClientPlayer(lIndex))
            Return
        End If

        If oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
            LogEvent(LogEventType.PossibleCheat, "Player trying a throwback command on a research tech: " & mlClientPlayer(lIndex))
            Return
        End If

        If yValue = PlayerSpecialTech.eyThrowBackType.ePushToBypass Then
            'ok, put techid, techtypeid into the bypass queue... then start a timer for the next link event
            If oTech.bInTheTank = True Then
                LogEvent(LogEventType.PossibleCheat, "Player trying to put a tech in the bypass queue that is already there: " & mlClientPlayer(lIndex))
                Return
            End If
            If oTech.GetResearcherCount() > 0 Then
                LogEvent(LogEventType.PossibleCheat, "Player trying to put tech in bypass that is being researched: " & mlClientPlayer(lIndex))
                Return
            End If

            oTech.yFlags = CByte(oTech.yFlags Or 2)
            oTech.SaveObject()

            Dim lCheckQueue As Int32 = 2592000         '24 hours of cycle time
            If oTech.GetResearchCost.PointsRequired < 238464000 Then        '24 hrs at 92 res time per cycle
                Dim lCnt As Int32 = 0
                Dim lIntel As Int32 = 0
                For X As Int32 = 0 To oPlayer.mlColonyUB
                    If oPlayer.mlColonyIdx(X) > -1 AndAlso glColonyIdx(oPlayer.mlColonyIdx(X)) = oPlayer.mlColonyID(X) Then
                        lCnt += 1
                        lIntel += CInt(goColony(oPlayer.mlColonyIdx(X)).GetColonyIntelligence())
                    End If
                Next X

                lIntel = lIntel \ lCnt

                lCheckQueue = 480 - ((lIntel - 100) * 6)   'in minutes
                If lCheckQueue < 0 Then lCheckQueue = 1 'minimum 1 minute
                lCheckQueue *= 60                  'now in seconds
                lCheckQueue *= 30                  'now in cycles
            End If
            AddToQueue(glCurrentCycle + lCheckQueue, QueueItemType.ePerformLinkTest, oPlayer.ObjectID, oTech.ObjectID, -1, -1, -1, -1, -1, -1)
        ElseIf yValue = PlayerSpecialTech.eyThrowBackType.eSwapFromBypass Then
            If oTech.bInTheTank = False Then
                LogEvent(LogEventType.PossibleCheat, "Player trying to pull out of the tank a tech not in the tank: " & mlClientPlayer(lIndex))
                Return
            End If

            'the techid we are putting into the bypass queue
            Dim lWithID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim oWith As PlayerSpecialTech = CType(oPlayer.GetTech(lWithID, ObjectType.eSpecialTech), PlayerSpecialTech)

            If oWith Is Nothing = False Then
                If oWith.bLinked = False Then
                    LogEvent(LogEventType.PossibleCheat, "Player trying to swap with unlinked tech: " & mlClientPlayer(lIndex))
                    Return
                End If
                If oWith.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                    LogEvent(LogEventType.PossibleCheat, "Player trying to swap with researched tech: " & mlClientPlayer(lIndex))
                    Return
                End If
                If oWith.GetResearcherCount > 0 Then
                    LogEvent(LogEventType.PossibleCheat, "Player trying to swap with tech that is being researched: " & mlClientPlayer(lIndex))
                    Return
                End If

                oPlayer.blCredits -= (oTech.GetResearchCost.CreditCost \ 10)
                oWith.yFlags = oWith.yFlags Or PlayerSpecialTech.eySpecialTechFlags.eInTheTank
                If oTech.bInTheTank = True Then oTech.yFlags = oTech.yFlags Xor PlayerSpecialTech.eySpecialTechFlags.eInTheTank

                oWith.SaveObject()
                oTech.SaveObject()


            ElseIf lWithID = -1 Then
                With oPlayer.oSpecials
                    Dim lInQueue As Int32 = 0
                    For X As Int32 = 0 To .mlSpecialTechUB
                        If .mlSpecialTechIdx(X) > -1 Then
                            Dim oTmpTech As PlayerSpecialTech = .moSpecialTech(X)
                            If oTmpTech Is Nothing = False AndAlso oTmpTech.bLinked = True AndAlso oTmpTech.bInTheTank = False AndAlso oTmpTech.ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                lInQueue += 1
                            End If
                        End If
                    Next X

                    If lInQueue < 3 Then
                        oPlayer.blCredits -= (oTech.GetResearchCost.CreditCost \ 10)
                        If oTech.bInTheTank = True Then oTech.yFlags = oTech.yFlags Xor PlayerSpecialTech.eySpecialTechFlags.eInTheTank
                        oTech.SaveObject()
                    Else
                        LogEvent(LogEventType.PossibleCheat, "Player trying to swap with non tech: " & mlClientPlayer(lIndex))
                        Return
                    End If
                End With
            End If


        End If

    End Sub
    Private Sub HandleSetSpecTechGuarantee(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.GuaranteedSpecialTechID > 0 Then Return

        oPlayer.GuaranteedSpecialTechID = System.BitConverter.ToInt32(yData, 2)
        oPlayer.SaveObject(True)
    End Sub
    Private Sub HandleSetRallyPoint(ByRef yData() As Byte, ByVal lIndex As Int32)
        '2 bytes for msg code
        Dim lPos As Int32 = 2
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To lCnt - 1
            Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim oEntity As Epica_Entity = CType(GetEpicaObject(lID, iTypeID), Epica_Entity)
            If oEntity Is Nothing = False Then
                With oEntity
                    .RallyPointX = lDestX
                    .RallyPointZ = lDestZ
                    .RallyPointA = iDestA
                    .RallyPointEnvirID = lDestID
                    .RallyPointEnvirTypeID = iDestTypeID
                End With
            End If
        Next X

    End Sub
    Private Sub HandleSetRouteMineral(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2 'for msg code
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lRouteIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim yFlag As Byte = yData(lPos) : lPos += 2

        Dim oUnit As Unit = GetEpicaUnit(lUnitID)
        If oUnit Is Nothing Then Return

        If oUnit.Owner.ObjectID = mlClientPlayer(lIndex) OrElse (oUnit.Owner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eMoveUnits) <> 0) Then
            If oUnit.lRouteUB >= lRouteIdx Then
                If oUnit.uRoute(lRouteIdx).oDest Is Nothing OrElse (oUnit.uRoute(lRouteIdx).oDest.ObjectID = lDestID AndAlso oUnit.uRoute(lRouteIdx).oDest.ObjTypeID = iDestTypeID) Then
                    'Load All situation
                    oUnit.uRoute(lRouteIdx).SetLoadItem(lObjID, iObjTypeID, oUnit.Owner, yFlag)
                    moClients(lIndex).SendData(yData)
                End If
            End If
        Else : LogEvent(LogEventType.PossibleCheat, "HandleSetRouteMineral: Owner Mismatch. PlayerID: " & mlClientPlayer(lIndex))
        End If
    End Sub
    Private Sub HandleSetTaxRate(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lColonyID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim yNewRate As Byte = yData(6)
        Dim yControlGrowth As Byte = yData(7)
        Dim yControlMorale As Byte = yData(8)

        Dim oColony As Colony = GetEpicaColony(lColonyID)
        If oColony Is Nothing Then Return

        Dim lOwnerID As Int32 = oColony.Owner.ObjectID

        If lOwnerID = mlClientPlayer(lIndex) OrElse (lOwnerID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eAlterColonyStats) <> 0) Then
            If yControlGrowth <> 255 Then
                Dim lRealVal As Int32 = CInt(yControlGrowth) - 128I
                If oColony.Owner.oSpecials.yControlledGrowthLimit < Math.Abs(lRealVal) Then
                    LogEvent(LogEventType.PossibleCheat, "Player attempting to control growth more than allowed. " & oColony.Owner.ObjectID)
                    Return
                End If
                oColony.iControlledGrowth = CShort(lRealVal)
            End If
            If yControlMorale <> 255 Then
                Dim lRealVal As Int32 = CInt(yControlMorale) - 128I
                If oColony.Owner.oSpecials.yControlledMoraleLimit < Math.Abs(lRealVal) Then
                    LogEvent(LogEventType.PossibleCheat, "Player attempting to control morale more than allowed. " & oColony.Owner.ObjectID)
                    Return
                End If
                oColony.iControlledMorale = CShort(lRealVal)
            End If

            oColony.TaxRate = yNewRate
            oColony.DataChanged()

            If yNewRate > 79 Then
                LogEvent(LogEventType.ExtensiveLogging, "Tax Rate set to " & yNewRate & " for colonyID: " & lColonyID & " by " & mlClientPlayer(lIndex))
            End If
        End If

    End Sub
    Private Sub HandleSetTransportStatus(ByVal yData() As Byte, ByVal lIndex As Int32)

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) <> -1 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then
            If (mlAliasedRights(lIndex) And (AliasingRights.eViewBattleGroups Or AliasingRights.eViewColonyStats Or AliasingRights.eModifyBattleGroups)) <> (AliasingRights.eViewColonyStats Or AliasingRights.eViewBattleGroups Or AliasingRights.eModifyBattleGroups) Then
                LogEvent(LogEventType.PossibleCheat, "SetTransportStatus with missing Alias Rights: " & mlClientPlayer(lIndex))
                Return
            End If
            lPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            Dim lPos As Int32 = 2
            Dim lTransID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim oTrans As Transport = oPlayer.GetTransport(lTransID)
            If oTrans Is Nothing Then Return

            Dim yType As Byte = yData(lPos) : lPos += 1

            Select Case yType
                Case 1          'begin route
                    If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then Return
                    oTrans.HandleBegin()
                    oTrans.Owner.ClearNextTransportEvent()

                    Dim yResp(18) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lTransID).CopyTo(yResp, lPos) : lPos += 4
                    yResp(lPos) = yType : lPos += 1
                    yResp(lPos) = 255 : lPos += 1
                    System.BitConverter.GetBytes(oTrans.DestinationID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oTrans.DestinationTypeID).CopyTo(yResp, lPos) : lPos += 2
                    yResp(lPos) = oTrans.TransFlags : lPos += 1
                    Dim lSecs As Int32 = 0
                    If oTrans.ETA <> DateTime.MinValue Then
                        lSecs = CInt(oTrans.ETA.Subtract(DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime).TotalSeconds)
                        If lSecs < 0 Then lSecs = 0
                    End If
                    System.BitConverter.GetBytes(lSecs).CopyTo(yResp, lPos) : lPos += 4

                    moClients(lIndex).SendData(yResp)
                Case 2          'clear route
                    If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then Return
                    oTrans.lRouteUB = -1
                    ReDim Preserve oTrans.oRoute(oTrans.lRouteUB)
                    oTrans.SaveObject()

                    Dim yResp(7) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lTransID).CopyTo(yResp, lPos) : lPos += 4
                    yResp(lPos) = yType : lPos += 1
                    yResp(lPos) = 255 : lPos += 1
                    moClients(lIndex).SendData(yResp)
                Case 3          'delete item
                    Dim yRoute As Byte = yData(lPos) : lPos += 1
                    Dim yAction As Byte = yData(lPos) : lPos += 1

                    If (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 AndAlso yRoute = oTrans.CurrentWaypoint Then Return

                    If yAction = 255 Then
                        'delete a route
                        If yRoute <= oTrans.lRouteUB Then
                            If oTrans.oRoute(yRoute) Is Nothing = False Then
                                'delete the route item
                                For X As Int32 = CInt(yRoute) To oTrans.lRouteUB - 1
                                    oTrans.oRoute(X) = oTrans.oRoute(X + 1)
                                    oTrans.oRoute(X).OrderNum -= CByte(1)
                                Next X
                                oTrans.lRouteUB -= 1
                                ReDim Preserve oTrans.oRoute(oTrans.lRouteUB)
                                'k, once done deleting, if our currentwaypoint > yRoute, currentwaypoint -= 1
                                If oTrans.CurrentWaypoint > yRoute Then
                                    Dim lNewVal As Int32 = oTrans.CurrentWaypoint
                                    lNewVal -= 1
                                    If lNewVal < 0 Then lNewVal = 0
                                    oTrans.CurrentWaypoint = CByte(lNewVal)
                                End If
                                oTrans.SaveObject()
                            End If
                        End If
                    Else
                        'delete an action
                        If yRoute <= oTrans.lRouteUB Then
                            If oTrans.oRoute(yRoute) Is Nothing = False Then
                                Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                                If yAction > oRoute.lActionUB Then Return
                                If oRoute.oActions(yAction) Is Nothing Then Return
                                For X As Int32 = CInt(yAction) To oRoute.lActionUB - 1
                                    oRoute.oActions(X) = oRoute.oActions(X + 1)
                                    oRoute.oActions(X).ActionOrderNum -= CByte(1)
                                Next X
                                oRoute.lActionUB -= 1
                                ReDim Preserve oRoute.oActions(oRoute.lActionUB)
                                oRoute.SaveObject(False)
                            End If
                        End If
                    End If

                    moClients(lIndex).SendData(yData)
                Case 4          'move item down
                    Dim yRoute As Byte = yData(lPos) : lPos += 1
                    Dim yAction As Byte = yData(lPos) : lPos += 1

                    If oTrans.lRouteUB < yRoute Then Return
                    If oTrans.oRoute(yRoute) Is Nothing Then Return

                    'check if we can do this
                    If (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                        If oTrans.CurrentWaypoint = yRoute Then Return

                        If yAction = 255 Then
                            If oTrans.CurrentWaypoint = CInt(yRoute) + 1 Then Return
                            If oTrans.lRouteUB = yRoute Then Return
                        Else
                            If oTrans.oRoute(yRoute).lActionUB < yAction Then Return
                        End If
                    End If

                    'Ok, we can - do it... moving down is increasing the order num
                    If yAction = 255 Then
                        'moving the route item down
                        Dim lFromIdx As Int32 = yRoute
                        Dim lToIdx As Int32 = yRoute
                        lToIdx += 1

                        Dim oTo As TransportRoute = oTrans.oRoute(lToIdx)
                        oTrans.oRoute(lToIdx) = oTrans.oRoute(lFromIdx)
                        oTrans.oRoute(lFromIdx) = oTo

                        oTrans.oRoute(lToIdx).OrderNum = CByte(lToIdx)
                        oTrans.oRoute(lFromIdx).OrderNum = CByte(lFromIdx)

                        oTrans.oRoute(lToIdx).SaveObject(False)
                        oTrans.oRoute(lFromIdx).SaveObject(False)
                    Else
                        'moving the action item down
                        Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                        If oRoute Is Nothing Then Return

                        Dim lFromIdx As Int32 = yAction
                        Dim lToIdx As Int32 = yAction
                        lToIdx += 1

                        Dim oTo As TransportRouteAction = oRoute.oActions(lToIdx)
                        oRoute.oActions(lToIdx) = oRoute.oActions(lFromIdx)
                        oRoute.oActions(lFromIdx) = oTo

                        oRoute.oActions(lToIdx).ActionOrderNum = CByte(lToIdx)
                        oRoute.oActions(lFromIdx).ActionOrderNum = CByte(lFromIdx)

                        oRoute.SaveObject(False)
                    End If

                    moClients(lIndex).SendData(yData)
                Case 5          'move item up
                    Dim yRoute As Byte = yData(lPos) : lPos += 1
                    Dim yAction As Byte = yData(lPos) : lPos += 1

                    If oTrans.lRouteUB < yRoute Then Return
                    If oTrans.oRoute(yRoute) Is Nothing Then Return

                    'check if we can do this
                    If (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                        If oTrans.CurrentWaypoint = yRoute Then Return

                        If yAction = 255 Then
                            If oTrans.CurrentWaypoint = CInt(yRoute) - 1 Then Return
                            If oTrans.lRouteUB = yRoute Then Return
                        Else
                            If oTrans.oRoute(yRoute).lActionUB < yAction Then Return
                        End If
                    End If

                    'Ok, we can - do it... moving down is increasing the order num
                    If yAction = 255 Then
                        'moving the route item down
                        Dim lFromIdx As Int32 = yRoute
                        Dim lToIdx As Int32 = yRoute
                        lToIdx -= 1

                        Dim oTo As TransportRoute = oTrans.oRoute(lToIdx)
                        oTrans.oRoute(lToIdx) = oTrans.oRoute(lFromIdx)
                        oTrans.oRoute(lFromIdx) = oTo

                        oTrans.oRoute(lToIdx).OrderNum = CByte(lToIdx)
                        oTrans.oRoute(lFromIdx).OrderNum = CByte(lFromIdx)

                        oTrans.oRoute(lToIdx).SaveObject(False)
                        oTrans.oRoute(lFromIdx).SaveObject(False)
                    Else
                        'moving the action item down
                        Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                        If oRoute Is Nothing Then Return

                        Dim lFromIdx As Int32 = yAction
                        Dim lToIdx As Int32 = yAction
                        lToIdx -= 1

                        Dim oTo As TransportRouteAction = oRoute.oActions(lToIdx)
                        oRoute.oActions(lToIdx) = oRoute.oActions(lFromIdx)
                        oRoute.oActions(lFromIdx) = oTo

                        oRoute.oActions(lToIdx).ActionOrderNum = CByte(lToIdx)
                        oRoute.oActions(lFromIdx).ActionOrderNum = CByte(lFromIdx)

                        oRoute.SaveObject(False)
                    End If

                    moClients(lIndex).SendData(yData)
                Case 6          'pause route
                    If (oTrans.TransFlags And Transport.elTransportFlags.ePaused) <> 0 Then
                        oTrans.TransFlags = oTrans.TransFlags Xor Transport.elTransportFlags.ePaused
                        If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) <> 0 Then oTrans.TransFlags = oTrans.TransFlags Xor Transport.elTransportFlags.eIdle
                        If (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) = 0 Then
                            oTrans.TransportArrived()
                        End If
                    Else
                        oTrans.TransFlags = oTrans.TransFlags Or Transport.elTransportFlags.ePaused
                    End If
                    oTrans.Owner.ClearNextTransportEvent()

                    Dim yResp(7) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lTransID).CopyTo(yResp, lPos) : lPos += 4
                    yResp(lPos) = yType : lPos += 1
                    If (oTrans.TransFlags And Transport.elTransportFlags.ePaused) <> 0 Then yResp(lPos) = 255 Else yResp(lPos) = 0
                    lPos += 1
                    moClients(lIndex).SendData(yResp)
                Case 7          'recall
                    oTrans.HandleRecall()
                    oTrans.Owner.ClearNextTransportEvent()

                    Dim yResp(13) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lTransID).CopyTo(yResp, lPos) : lPos += 4
                    yResp(lPos) = yType : lPos += 1
                    yResp(lPos) = oTrans.TransFlags : lPos += 1
                    System.BitConverter.GetBytes(oTrans.DestinationID).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oTrans.DestinationTypeID).CopyTo(yResp, lPos) : lPos += 2
                    moClients(lIndex).SendData(yResp)
                Case 8          'loop orders
                    If (oTrans.TransFlags And Transport.elTransportFlags.eLoop) <> 0 Then
                        oTrans.TransFlags = oTrans.TransFlags Xor Transport.elTransportFlags.eLoop
                    Else
                        oTrans.TransFlags = oTrans.TransFlags Or Transport.elTransportFlags.eLoop
                    End If

                    Dim yResp(7) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetTransportStatus).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lTransID).CopyTo(yResp, lPos) : lPos += 4
                    yResp(lPos) = yType : lPos += 1
                    If (oTrans.TransFlags And Transport.elTransportFlags.eLoop) <> 0 Then yResp(lPos) = 255 Else yResp(lPos) = 0
                    lPos += 1
                    moClients(lIndex).SendData(yResp)
                Case 9          'add dest
                    Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    'let's confirm the dest
                    If iDestTypeID <> ObjectType.eColony Then
                        LogEvent(LogEventType.PossibleCheat, "Player setting transport route to non colony! Player: " & mlClientPlayer(lIndex))
                        Return
                    End If
                    Dim oColony As Colony = GetEpicaColony(lDestID)
                    If oColony Is Nothing OrElse oColony.Owner Is Nothing OrElse oColony.Owner.ObjectID <> oTrans.OwnerID Then Return

                    Dim oRoute As New TransportRoute()
                    With oRoute
                        .oTransport = oTrans
                        .lActionUB = -1
                        ReDim .oActions(-1)
                        .DestinationID = lDestID
                        .DestinationTypeID = iDestTypeID
                        .WaypointFlags = 0
                        .OrderNum = CByte(oTrans.lRouteUB + 1)
                    End With
                    ReDim Preserve oTrans.oRoute(oTrans.lRouteUB + 1)
                    oTrans.oRoute(oTrans.lRouteUB + 1) = oRoute
                    oTrans.lRouteUB += 1
                    oTrans.SaveObject()

                    moClients(lIndex).SendData(yData)
                Case 10         'discard cargo
                    If (oTrans.TransFlags And Transport.elTransportFlags.eIdle) = 0 Then Return

                    Dim lCargoID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iCargoTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim lCargoOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    Dim oCargo As TransportCargo = oTrans.GetCargo(lCargoID, iCargoTypeID, lCargoOwnerID)
                    If oCargo Is Nothing = False Then
                        oCargo.Quantity = 0
                        oCargo.SaveObject(False)
                    End If

                    moClients(lIndex).SendData(yData)
                Case 11         'add action
                    Dim yRoute As Byte = yData(lPos) : lPos += 1

                    If yRoute > oTrans.lRouteUB Then Return
                    Dim oRoute As TransportRoute = oTrans.oRoute(yRoute)
                    If oRoute Is Nothing Then Return
                    If oTrans.CurrentWaypoint = yRoute AndAlso (oTrans.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then Return
                    If oRoute.lActionUB > 8 Then Return

                    Dim oNew As New TransportRouteAction
                    With oNew
                        .ActionOrderNum = CByte(oRoute.lActionUB + 1)
                        .ActionTypeID = CType(yData(lPos), TransportRouteAction.TransportRouteActionType) : lPos += 1
                        .Extended1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .Extended2 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        .Extended3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .oParentRoute = oRoute
                    End With
                    ReDim Preserve oRoute.oActions(oRoute.lActionUB + 1)
                    oRoute.oActions(oRoute.lActionUB + 1) = oNew
                    oRoute.lActionUB += 1
                    oRoute.SaveObject(False)

                    moClients(lIndex).SendData(yData)
                Case 12 'Rename Transport
                    Dim sName As String = GetStringFromBytes(yData, lPos, 20)
                    If sName <> "" Then
                        oTrans.UnitName = StringToBytes(sName)
                        oTrans.SaveObject()
                    End If
                    Dim yResp(25) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransportName).CopyTo(yResp, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lTransID).CopyTo(yResp, lPos) : lPos += 4
                    oTrans.UnitName.CopyTo(yResp, lPos) : lPos += 20
                    moClients(lIndex).SendData(yResp)
            End Select
        End If
    End Sub
    Private Sub HandleSubmitComponentDesignMsg(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
        'Several Component msgs... but same header...
        Dim lPos As Int32 = 2       'using pos to make my life easier

        'get the objecttypeid
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Get the researcher guid
        Dim lResID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iResTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oTech As Epica_Tech = Nothing

        Dim bPreExist As Boolean = False

        Dim oResearcher As Object = Nothing
        oResearcher = GetEpicaObject(lResID, iResTypeID)
        If oResearcher Is Nothing Then
            'TODO: Return failure msg
            Return
        End If

        'Validate that the player belongs to the socket and that the researcher's owner is the player
        If (mlClientPlayer(lSocketIndex) <> CType(oResearcher, Epica_Entity).Owner.ObjectID) AndAlso (mlAliasedAs(lSocketIndex) <> CType(oResearcher, Epica_Entity).Owner.ObjectID OrElse (mlAliasedRights(lSocketIndex) And AliasingRights.eCreateDesigns) = 0) Then
            LogEvent(LogEventType.PossibleCheat, "Player without rights to researcher submitting design: " & mlClientPlayer(lSocketIndex))
            Return
        End If

        'Now, create our new tech object (if needed)
        If lTechID <> -1 Then
            oTech = CType(oResearcher, Epica_Entity).Owner.GetTech(lTechID, iObjTypeID)
            If oTech.Owner.ObjectID = 0 Then oTech = Nothing
            bPreExist = True
        End If

        If oTech Is Nothing = False AndAlso bPreExist = True AndAlso oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
            oTech = Nothing
        End If

        If oTech Is Nothing Then
            bPreExist = False
            Select Case iObjTypeID
                Case ObjectType.eAlloyTech
                    oTech = New AlloyTech
                Case ObjectType.eEngineTech
                    oTech = New EngineTech
                Case ObjectType.eShieldTech
                    oTech = New ShieldTech
                Case ObjectType.eArmorTech
                    oTech = New ArmorTech
                Case ObjectType.eRadarTech
                    oTech = New RadarTech
                    'Case ObjectType.eHangarTech
                    '    oTech = New HangarTech
                Case ObjectType.eHullTech
                    oTech = New HullTech
                Case ObjectType.ePrototype
                    oTech = New Prototype
                Case ObjectType.eWeaponTech
                    oTech = BaseWeaponTech.CreateWeaponClass(yData(lPos))
            End Select
        End If

        If oTech Is Nothing = False Then
            If bPreExist = False Then
                oTech.ObjectID = -1
            Else
                If oTech.GetResearcherCount > 0 Then
                    oTech = Nothing
                    Dim yAlert(13) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yAlert, 0)
                    CType(oResearcher, Epica_Entity).GetGUIDAsString.CopyTo(yAlert, 2)
                    System.BitConverter.GetBytes(CInt(-50)).CopyTo(yAlert, 8)
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yAlert, 12)
                    CType(oResearcher, Epica_Entity).Owner.SendPlayerMessage(yAlert, True, AliasingRights.eAddProduction Or AliasingRights.eCancelProduction)
                    'Now, break out before anything else...
                    Return
                End If
            End If

            oTech.ObjTypeID = iObjTypeID

            'Set the tech's owner
            oTech.Owner = CType(oResearcher, Epica_Entity).Owner

            If oTech.SetFromDesignMsg(yData) = False Then
                oTech = Nothing
                Dim yAlert(13) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yAlert, 0)
                CType(oResearcher, Epica_Entity).GetGUIDAsString.CopyTo(yAlert, 2)
                'TODO: is this right??? 0 for the id and the objtypeid...
                System.BitConverter.GetBytes(CInt(0)).CopyTo(yAlert, 8)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yAlert, 12)
                CType(oResearcher, Epica_Entity).Owner.SendPlayerMessage(yAlert, True, AliasingRights.eAddProduction Or AliasingRights.eCancelProduction)
                'Now, break out before anything else...
                Return
            End If

            If bPreExist = False Then
                Select Case oTech.ObjTypeID
                    Case ObjectType.eAlloyTech
                        oTech.Owner.mlAlloyUB += 1
                        ReDim Preserve oTech.Owner.mlAlloyIdx(oTech.Owner.mlAlloyUB)
                        ReDim Preserve oTech.Owner.moAlloy(oTech.Owner.mlAlloyUB)
                        oTech.Owner.moAlloy(oTech.Owner.mlAlloyUB) = CType(oTech, AlloyTech)
                    Case ObjectType.eArmorTech
                        oTech.Owner.mlArmorUB += 1
                        ReDim Preserve oTech.Owner.mlArmorIdx(oTech.Owner.mlArmorUB)
                        ReDim Preserve oTech.Owner.moArmor(oTech.Owner.mlArmorUB)
                        oTech.Owner.moArmor(oTech.Owner.mlArmorUB) = CType(oTech, ArmorTech)
                    Case ObjectType.eEngineTech
                        oTech.Owner.mlEngineUB += 1
                        ReDim Preserve oTech.Owner.mlEngineIdx(oTech.Owner.mlEngineUB)
                        ReDim Preserve oTech.Owner.moEngine(oTech.Owner.mlEngineUB)
                        oTech.Owner.moEngine(oTech.Owner.mlEngineUB) = CType(oTech, EngineTech)
                        'Case ObjectType.eHangarTech
                        '    oTech.Owner.mlHangarUB += 1
                        '    ReDim Preserve oTech.Owner.mlHangarIdx(oTech.Owner.mlHangarUB)
                        '    ReDim Preserve oTech.Owner.moHangar(oTech.Owner.mlHangarUB)
                        '    oTech.Owner.moHangar(oTech.Owner.mlHangarUB) = CType(oTech, HangarTech)
                    Case ObjectType.eHullTech
                        oTech.Owner.mlHullUB += 1
                        ReDim Preserve oTech.Owner.mlHullIdx(oTech.Owner.mlHullUB)
                        ReDim Preserve oTech.Owner.moHull(oTech.Owner.mlHullUB)
                        oTech.Owner.moHull(oTech.Owner.mlHullUB) = CType(oTech, HullTech)
                    Case ObjectType.ePrototype
                        oTech.Owner.mlPrototypeUB += 1
                        ReDim Preserve oTech.Owner.mlPrototypeIdx(oTech.Owner.mlPrototypeUB)
                        ReDim Preserve oTech.Owner.moPrototype(oTech.Owner.mlPrototypeUB)
                        oTech.Owner.moPrototype(oTech.Owner.mlPrototypeUB) = CType(oTech, Prototype)
                    Case ObjectType.eRadarTech
                        oTech.Owner.mlRadarUB += 1
                        ReDim Preserve oTech.Owner.mlRadarIdx(oTech.Owner.mlRadarUB)
                        'ReDim Preserve oTech.Owner.moAlloy(oTech.Owner.mlRadarUB)
                        ReDim Preserve oTech.Owner.moRadar(oTech.Owner.mlRadarUB)
                        oTech.Owner.moRadar(oTech.Owner.mlRadarUB) = CType(oTech, RadarTech)
                    Case ObjectType.eShieldTech
                        oTech.Owner.mlShieldUB += 1
                        ReDim Preserve oTech.Owner.mlShieldIdx(oTech.Owner.mlShieldUB)
                        ReDim Preserve oTech.Owner.moShield(oTech.Owner.mlShieldUB)
                        oTech.Owner.moShield(oTech.Owner.mlShieldUB) = CType(oTech, ShieldTech)
                    Case ObjectType.eWeaponTech
                        oTech.Owner.mlWeaponUB += 1
                        ReDim Preserve oTech.Owner.mlWeaponIdx(oTech.Owner.mlWeaponUB)
                        ReDim Preserve oTech.Owner.moWeapon(oTech.Owner.mlWeaponUB)
                        oTech.Owner.moWeapon(oTech.Owner.mlWeaponUB) = CType(oTech, BaseWeaponTech)
                End Select
            Else
                oTech.MajorDesignFlaw = 0
                oTech.ErrorReasonCode = 0

                If oTech.ObjTypeID = ObjectType.eHullTech Then CType(oTech, HullTech).ClearPowerRequired()
            End If

            'Set our component development phase...
            oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eComponentDesign

            oTech.PopIntel = 100
            If CType(oResearcher, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                With CType(oResearcher, Facility)
                    If .ParentColony Is Nothing = False Then
                        oTech.PopIntel = .ParentColony.GetColonyIntelligence()
                    End If
                End With
            End If

            'TODO: If ValidateDesign fails, that is likely a cheat scenario
            If oTech.ValidateDesign = True AndAlso oTech.SaveObject Then
                If bPreExist = False Then
                    Select Case oTech.ObjTypeID
                        Case ObjectType.eAlloyTech
                            oTech.Owner.mlAlloyIdx(oTech.Owner.mlAlloyUB) = oTech.ObjectID
                        Case ObjectType.eArmorTech
                            oTech.Owner.mlArmorIdx(oTech.Owner.mlArmorUB) = oTech.ObjectID
                        Case ObjectType.eEngineTech
                            oTech.Owner.mlEngineIdx(oTech.Owner.mlEngineUB) = oTech.ObjectID
                            'Case ObjectType.eHangarTech
                            '    oTech.Owner.mlHangarIdx(oTech.Owner.mlHangarUB) = oTech.ObjectID
                        Case ObjectType.eHullTech
                            oTech.Owner.mlHullIdx(oTech.Owner.mlHullUB) = oTech.ObjectID
                        Case ObjectType.ePrototype
                            oTech.Owner.mlPrototypeIdx(oTech.Owner.mlPrototypeUB) = oTech.ObjectID
                        Case ObjectType.eRadarTech
                            oTech.Owner.mlRadarIdx(oTech.Owner.mlRadarUB) = oTech.ObjectID
                        Case ObjectType.eShieldTech
                            oTech.Owner.mlShieldIdx(oTech.Owner.mlShieldUB) = oTech.ObjectID
                        Case ObjectType.eWeaponTech
                            oTech.Owner.mlWeaponIdx(oTech.Owner.mlWeaponUB) = oTech.ObjectID
                    End Select
                End If

                Dim oProdCost As ProductionCost = oTech.GetCurrentProductionCost

                'Add the Tech to the oResearcher as CurrentProduction
                If CType(oResearcher, Epica_GUID).ObjTypeID = ObjectType.eFacility Then

                    Dim oFac As Facility = CType(oResearcher, Facility)
                    If oFac.bProducing = False AndAlso oFac.AddProduction(oTech.ObjectID, oTech.ObjTypeID, 0, 1, 0) = True Then
                        'TODO: this sucks, I have to search through the list AGAIN!
                        Dim lIdx As Int32 = -1
                        For lTmpIdx As Int32 = 0 To glFacilityUB
                            If glFacilityIdx(lTmpIdx) = lResID Then
                                lIdx = lTmpIdx
                                Exit For
                            End If
                        Next lTmpIdx
                        If lIdx <> -1 Then AddEntityProducing(lIdx, ObjectType.eFacility, lResID)
                    Else
                        'TODO: Set Entity Prod failed?
                    End If
                Else
                    'TODO: Need to figure that out..,.
                End If

                'Now, send oTech to the player
                moClients(lSocketIndex).SendData(GetAddObjectMessage(oTech, GlobalMessageCode.eAddObjectCommand))
            Else
                oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign
                moClients(lSocketIndex).SendData(GetAddObjectMessage(oTech, GlobalMessageCode.eAddObjectCommand))
            End If
        End If

    End Sub
    Private Sub HandleSubmitMission(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eAlterAgents) = 0 Then Return
        End If

        Dim lPos As Int32 = 2 'for msgcode
        Dim yActionID As Byte = yData(lPos) : lPos += 1
        Dim lPM_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oOwner As Player = GetEpicaPlayer(lPlayerID)

        If yActionID = 2 Then
            'delete the mission by PM_ID
            For X As Int32 = 0 To glPlayerMissionUB
                If glPlayerMissionIdx(X) = lPM_ID Then
                    'moMission.lCurrentPhase = eMissionPhase.eInPlanning OrElse moMission.lCurrentPhase = eMissionPhase.ePreparationTime Then
                    If goPlayerMission(X).lCurrentPhase = eMissionPhase.eInPlanning OrElse goPlayerMission(X).lCurrentPhase = eMissionPhase.ePreparationTime Then
                        goPlayerMission(X).lCurrentPhase = eMissionPhase.eCancelled
                    Else : LogEvent(LogEventType.PossibleCheat, "HandleSubmitMission: Cancel attempt on non-cancel state mission. Player: " & mlClientPlayer(lIndex))
                    End If

                    Exit For
                End If
            Next X
        Else
            'ok, check PM_ID
            Dim oPM As PlayerMission = Nothing
            'MissionID (4)
            Dim lMissionID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'If lPM_ID > 0 Then
            '	oPM = GetEpicaPlayerMission(lPM_ID)
            'Else
            Dim oMission As Mission = GetEpicaMission(lMissionID)
            If oMission Is Nothing Then
                LogEvent(LogEventType.PossibleCheat, "HandleSubmitMission: Mission Not Found. Player: " & mlClientPlayer(lIndex))
                Return
            End If
            oPM = New PlayerMission(oMission, oOwner)
            'End If
            If oPM Is Nothing Then Return

            With oPM
                .PM_ID = lPM_ID
                Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4 'TargetPlayerID (4)
                If lTemp > 0 Then .oTarget = GetEpicaPlayer(lTemp) Else .oTarget = Nothing
                .lTargetID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4 'TargetID1 (4)
                .iTargetTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2 'TargetTypeID1 (2)
                .lTargetID2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4 'TargetID2 (4)
                .iTargetTypeID2 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2 'TargetTypeID2 (2)
                .ySafeHouseSetting = yData(lPos) : lPos += 1                    'SafeHouseSetting (1)
                .lMethodID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4   'methodid (4)

                'ok, first, check our safehousesetting
                If .ySafeHouseSetting <> 0 Then
                    If .oSafeHouseGoal Is Nothing Then
                        .oSafeHouseGoal = New PlayerMissionGoal()
                        .oSafeHouseGoal.oGoal = Goal.GetSafehouseGoal()
                        .oSafeHouseGoal.oMission = oPM
                    End If
                    'skillsetid
                    lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .oSafeHouseGoal.oSkillSet = .oSafeHouseGoal.oGoal.GetOrAddSkillSet(lTemp, True)
                    'assignment1 agentid
                    Dim lAgent1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    'assignment1 skillid
                    Dim lSkill1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    'assignment2 agentid
                    Dim lAgent2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    'assignment2 skillid
                    Dim lSkill2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    If lAgent1 > -1 Then
                        Dim oAgent As Agent = GetEpicaAgent(lAgent1)
                        If oAgent Is Nothing = False Then
                            If oAgent.GetSkillValue(lSkill1, True, 0) = 0 Then
                                If oAgent.GetSkillValue(lSkillHardcodes.eNaturallyTalented, True, 0) = 0 Then
                                    LogEvent(LogEventType.PossibleCheat, "HandleSubmitMission: Agent does not possess assigned skill. Player: " & mlClientPlayer(lIndex))
                                    Return
                                End If
                            End If
                            .oSafeHouseGoal.AddAgentAssignment(GetEpicaAgent(lAgent1), GetEpicaSkill(lSkill1))
                        End If
                    End If
                    If lAgent2 > -1 Then
                        Dim oAgent As Agent = GetEpicaAgent(lAgent2)
                        If oAgent Is Nothing = False Then
                            If oAgent.GetSkillValue(lSkill2, True, 0) = 0 Then
                                If oAgent.GetSkillValue(lSkillHardcodes.eNaturallyTalented, True, 0) = 0 Then
                                    LogEvent(LogEventType.PossibleCheat, "HandleSubmitMission: Agent does not possess assigned skill. Player: " & mlClientPlayer(lIndex))
                                    Return
                                End If
                            End If
                            .oSafeHouseGoal.AddAgentAssignment(GetEpicaAgent(lAgent2), GetEpicaSkill(lSkill2))
                        End If
                    End If
                End If

                lTemp = yData(lPos) : lPos += 1                     'PhaseCnt (1)
                For X As Int32 = 0 To lTemp - 1
                    Dim lPhase As Int32 = yData(lPos) : lPos += 1 '	Phase (1)
                    Dim lCoverAgentCnt As Int32 = yData(lPos) : lPos += 1 '	CoverAgentCnt (1)
                    For Y As Int32 = 0 To lCoverAgentCnt - 1
                        Dim lCA As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4 '		CoverAgentID (4)
                        Dim oCA As Agent = GetEpicaAgent(lCA)
                        If oCA Is Nothing = False Then .AddPhaseCoverAgent(lPhase, oCA, 0)
                    Next Y
                Next X

                'GoalCnt (1)
                lTemp = yData(lPos) : lPos += 1
                For X As Int32 = 0 To lTemp - 1
                    Dim lGoalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4         '   GoalID (4)
                    Dim lSkillSetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4     '	SkillsetID (4)
                    Dim lAssCnt As Int32 = yData(lPos) : lPos += 1                                      '	AssignmentCnt (1)

                    Dim oMG As PlayerMissionGoal = Nothing
                    For Y As Int32 = 0 To .oMission.GoalUB
                        If .oMission.MethodIDs(Y) = .lMethodID AndAlso .oMission.Goals(Y).ObjectID = lGoalID Then
                            oMG = .oMissionGoals(Y)
                            Exit For
                        End If
                    Next Y
                    If oMG Is Nothing = False Then
                        Dim oSkillSet As SkillSet = Nothing
                        For Y As Int32 = 0 To oMG.oGoal.SkillSetUB
                            If oMG.oGoal.SkillSets(Y).SkillSetID = lSkillSetID Then
                                oSkillSet = oMG.oGoal.SkillSets(Y)
                                Exit For
                            End If
                        Next Y
                        If oSkillSet Is Nothing = False Then
                            oMG.oSkillSet = oSkillSet
                        ElseIf yActionID = 1 Then
                            LogEvent(LogEventType.PossibleCheat, "HandleSubmitMission: SkillsetID " & lSkillSetID & " is not part of goal " & lGoalID & ". Player: " & mlClientPlayer(lIndex))
                            Return
                        End If
                    ElseIf yActionID = 1 Then
                        LogEvent(LogEventType.PossibleCheat, "HandleSubmitMission: GoalID " & lGoalID & " not part of mission " & .oMission.ObjectID & ". Player: " & mlClientPlayer(lIndex))
                        Return
                    End If

                    'clear our assignments
                    oMG.lAssignmentUB = -1

                    For Y As Int32 = 0 To lAssCnt - 1
                        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4    '		AgentID (4)
                        Dim lSkillID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4    '		SkillID (4)
                        Dim oAgent As Agent = GetEpicaAgent(lAgentID)
                        Dim oSkill As Skill = GetEpicaSkill(lSkillID)

                        If oAgent.GetSkillValue(lSkillID, False, -1) = -1 Then
                            LogEvent(LogEventType.PossibleCheat, "HandleSubmitMission: Agent does not possess assigned skill. Player: " & mlClientPlayer(lIndex))
                            Return
                        End If

                        Dim oAA As AgentAssignment = oMG.AddAgentAssignment(oAgent, oSkill)
                    Next Y
                Next X

                If yActionID = 1 Then .lCurrentPhase = eMissionPhase.eWaitingToExecute Else .lCurrentPhase = eMissionPhase.eInPlanning
            End With

            'Ok, now, let's validate our data
            Dim bAdd As Boolean = oPM.PM_ID < 1
            Dim lResult As Int32 = oPM.DataIsValid(yActionID = 1)
            If lResult = 1 Then
                'ok, all is good... 
                If bAdd = True Then
                    If oPM.SaveObject() = False Then
                        Dim yResp(9) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitMission).CopyTo(yResp, 0)
                        System.BitConverter.GetBytes(lPM_ID).CopyTo(yResp, 2)
                        System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yResp, 6)
                        moClients(lIndex).SendData(yResp)
                        Return
                    End If

                    Dim lIdx As Int32 = -1
                    For X As Int32 = 0 To glPlayerMissionUB
                        If glPlayerMissionIdx(X) = -1 Then
                            goPlayerMission(X) = oPM
                            glPlayerMissionIdx(X) = oPM.PM_ID
                            lIdx = X
                            Exit For
                        End If
                    Next X
                    If lIdx = -1 Then
                        lIdx = glPlayerMissionUB + 1
                        ReDim Preserve glPlayerMissionIdx(glPlayerMissionUB + 1)
                        ReDim Preserve goPlayerMission(glPlayerMissionUB + 1)
                        goPlayerMission(lIdx) = oPM
                        glPlayerMissionIdx(lIdx) = oPM.PM_ID
                        glPlayerMissionUB += 1
                    End If
                Else
                    For X As Int32 = 0 To glPlayerMissionUB
                        If glPlayerMissionIdx(X) = oPM.PM_ID Then
                            'Do any mission in progress checks here and stuff
                            If goPlayerMission(X).oMission Is Nothing = False AndAlso goPlayerMission(X).oMission.ObjectID <> oPM.oMission.ObjectID Then
                                'cannot change the mission
                                oPM = goPlayerMission(X)
                            Else
                                If goPlayerMission(X).lCurrentPhase <> eMissionPhase.eSettingTheStage AndAlso goPlayerMission(X).lCurrentPhase <> eMissionPhase.eFlippingTheSwitch Then
                                    goPlayerMission(X) = oPM
                                Else
                                    'cannot change the mission
                                    oPM = goPlayerMission(X)
                                End If
                            End If

                            Exit For
                        End If
                    Next X
                    If oPM.SaveObject() = False Then
                        Dim yResp(9) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitMission).CopyTo(yResp, 0)
                        System.BitConverter.GetBytes(lPM_ID).CopyTo(yResp, 2)
                        System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yResp, 6)
                        moClients(lIndex).SendData(yResp)
                        Return
                    End If
                End If

                'Now, send the mission back
                oOwner.SendPlayerMessage(oPM.GetAddObjectMessage(), False, AliasingRights.eViewAgents)
            Else
                'uh... send a message back?
                Dim yResp(9) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSubmitMission).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lPM_ID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(lResult).CopyTo(yResp, 6)
                moClients(lIndex).SendData(yResp)
            End If
        End If

    End Sub
    Private Sub HandleSubmitTrade(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lTradeID As Int32 = System.BitConverter.ToInt32(yData, lPos)

        Dim oTrade As DirectTrade = Nothing

        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 Then
            lPlayerID = mlAliasedAs(lIndex)
            If (mlAliasedRights(lIndex) And AliasingRights.eAlterTrades) = 0 Then Return
        End If

        If lTradeID < 1 Then
            'Creating a new trade object... the TradeID is the target PlayerID (negative though)
            lTradeID = Math.Abs(lTradeID)
            oTrade = New DirectTrade()
            With oTrade
                .ObjectID = -1
                .ObjTypeID = ObjectType.eTrade
                .oPlayer1TP = New TradePlayer()
                .oPlayer2TP = New TradePlayer()
            End With
            With oTrade.oPlayer1TP
                .PlayerID = lPlayerID
                .lItemUB = -1
            End With
            With oTrade.oPlayer2TP
                .PlayerID = lTradeID
                .lItemUB = -1
            End With

            Dim oFromPlayer As Player = GetEpicaPlayer(lPlayerID)
            Dim oOtherPlayer As Player = GetEpicaPlayer(lTradeID)
            If oOtherPlayer Is Nothing = False AndAlso oFromPlayer Is Nothing = False Then
                If oOtherPlayer.lConnectedPrimaryID < 0 AndAlso (oOtherPlayer.iInternalEmailSettings And eEmailSettings.eTradeRequested) <> 0 Then
                    Dim sBody As String = BytesToString(oFromPlayer.PlayerName) & " has submitted a trade agreement for your approval."
                    Dim sTitle As String = "New Trade Agreement"
                    Dim oPC As PlayerComm = oOtherPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, oOtherPlayer.ObjectID, GetDateAsNumber(Now), False, oOtherPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False AndAlso (oOtherPlayer.iEmailSettings And eEmailSettings.eTradeRequested) <> 0 Then goMsgSys.SendOutboundEmail(oPC, oOtherPlayer, GlobalMessageCode.eSubmitTrade, 0, 0, 0, 0, 0, 0, "")
                End If
            End If

            If oTrade.SaveObject() = False Then
                'TODO: Alert the player
                Return
            Else
                If goGTC Is Nothing Then goGTC = New GalacticTradeSystem()
                goGTC.AddNewDirectTrade(oTrade)
            End If
        Else
            oTrade = goGTC.GetDirectTrade(lTradeID)
        End If

        If oTrade Is Nothing = False Then
            oTrade.HandleSubmitTradeMsg(yData, lPlayerID)
        End If

    End Sub
    Private Sub HandleTransferContents(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lFromID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iFromTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lToID As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim iToTypeID As Int16 = System.BitConverter.ToInt16(yData, 18)
        Dim lQty As Int32 = System.BitConverter.ToInt32(yData, 20)

        'Now, dim a colony object
        Dim oColony As Colony = Nothing

        'And do our transfer
        If iFromTypeID = ObjectType.eColony Then
            Select Case iObjTypeID
                Case ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eMineral, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers
                Case Else
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony's iObjTypeID not acceptable. Object Type: " & iObjTypeID & ", PlayerID: " & mlClientPlayer(lIndex))
                    Return
            End Select

            'transfers from a colony to a target...
            oColony = GetEpicaColony(lFromID)

            If oColony Is Nothing = False Then
                If oColony.Owner.ObjectID <> mlClientPlayer(lIndex) AndAlso (oColony.Owner.ObjectID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eTransferCargo) = 0) Then
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony Owner is not Socket Owner! FromColony Owner: " & oColony.Owner.ObjectID & ", Socket: " & mlClientPlayer(lIndex))
                    Return
                End If
            Else
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony is nothing. PlayerID: " & mlClientPlayer(lIndex))
                Return
            End If

            Dim oTo As Epica_Entity = Nothing
            If iToTypeID = ObjectType.eUnit Then
                oTo = GetEpicaUnit(lToID)
            ElseIf iToTypeID = ObjectType.eFacility Then
                oTo = GetEpicaFacility(lToID)
            Else
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: From Colony, ToType is not unit or facility. PlayerID: " & mlClientPlayer(lIndex))
                Return
            End If
            If oTo Is Nothing Then
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: From Colony. To Entity not found. PlayerID: " & mlClientPlayer(lIndex))
                Return
            End If

            If oTo.yProductionType = ProductionType.eMining Then
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: From Colony. To Entity Production is Mining. PlayerID: " & mlClientPlayer(lIndex))
            End If

            If (oTo.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then
                'Indicate that the destination's cargo bay is inoperable 
                lQty = -1
                System.BitConverter.GetBytes(lQty).CopyTo(yData, 20)
                moClients(lIndex).SendData(yData)
                Return
            End If

            If oTo.ObjTypeID = ObjectType.eUnit Then
                If CType(oTo.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    If (CType(oTo.ParentObject, Facility).CurrentStatus And elUnitStatus.eHangarOperational) = 0 Then
                        LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony. To Is Unit. Parent is Facility. Facility's hangar is inoperable. PlayerID: " & mlClientPlayer(lIndex))
                        Return
                    End If

                    If CType(oTo.ParentObject, Facility).ParentColony Is Nothing OrElse CType(oTo.ParentObject, Facility).ParentColony.ObjectID <> oColony.ObjectID Then
                        LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony. To Is unit. Parent is Facility. Facility's colony is not source colony. PlayerID: " & mlClientPlayer(lIndex))
                        Return
                    End If
                Else
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony. To is Unit that is not in facility. PlayerID: " & mlClientPlayer(lIndex))
                    Return
                End If
            Else
                If CType(oTo, Facility).ParentColony Is Nothing OrElse CType(oTo, Facility).ParentColony.ObjectID <> oColony.ObjectID Then
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony. To Is Facility. Facility's colony is not source colony. PlayerID: " & mlClientPlayer(lIndex))
                    Return
                End If
            End If

            'Now, check cargo space available
            Dim lToCargoCap As Int32 = oTo.Cargo_Cap
            If iObjTypeID = ObjectType.eColonists OrElse iObjTypeID = ObjectType.eOfficers OrElse iObjTypeID = ObjectType.eEnlisted Then
                lToCargoCap \= CInt(oTo.Owner.oSpecials.yPersonnelCargoUsage)
            End If
            If lToCargoCap < lQty Then lQty = lToCargoCap

            Dim oProdCost As ProductionCost = New ProductionCost()
            'Ok, quantity should be adjusted... now, let's create a production cost
            With oProdCost
                .ColonistCost = 0 : .CreditCost = 0 : .EnlistedCost = 0 : .PointsRequired = 0 : .ProductionCostType = 0 : .OfficerCost = 0
                If iObjTypeID = ObjectType.eColonists Then
                    .ColonistCost = lQty
                ElseIf iObjTypeID = ObjectType.eOfficers Then
                    .OfficerCost = lQty
                ElseIf iObjTypeID = ObjectType.eEnlisted Then
                    .EnlistedCost = lQty
                Else : .AddProductionCostItem(lObjID, iObjTypeID, lQty)
                End If
                'Set the objectypeid = specialtech so that if the player is in death budget, nothing weird happens
                .ObjTypeID = ObjectType.eSpecialTech
                .ProductionCostType = 1
            End With

            'Now, do a production cost assessment to the colony
            If oColony.HasRequiredResources(oProdCost, Nothing, eyHasRequiredResourcesFlags.DoNotWait) = eResourcesResult.Sufficient Then
                If iObjTypeID = ObjectType.eMineral Then
                    oTo.AddMineralCacheToCargo(lObjID, lQty)
                ElseIf iObjTypeID = ObjectType.eOfficers OrElse iObjTypeID = ObjectType.eEnlisted OrElse iObjTypeID = ObjectType.eColonists Then
                    oTo.AddPersonnelCacheToCargo(iObjTypeID, lQty)
                Else
                    Dim oTech As Epica_Tech = oTo.Owner.GetTech(lObjID, iObjTypeID)
                    If oTech Is Nothing Then oTech = QuickLookupTechnology(lObjID, iObjTypeID)
                    If oTech Is Nothing Then
                        LogEvent(LogEventType.CriticalError, "HandleTransferContents: FromColony. Colony used resoures. Tech was not found! oTo: " & oTo.ObjectID & ", GUID: " & lObjID & ", " & iObjTypeID)
                        'refund the cargo
                        'oTo.Cargo_Cap += lQty
                        Return
                    Else : oTo.AddComponentCacheToCargo(lObjID, iObjTypeID, lQty, oTech.Owner.ObjectID)
                    End If
                End If
            Else
                'refund the to cargo cap. the low resources alert should have indicated to the player the issue so we are done
                'oTo.Cargo_Cap += lQty
            End If
        ElseIf iToTypeID = ObjectType.eColony Then
            Select Case iObjTypeID
                Case ObjectType.eComponentCache, ObjectType.eMineralCache, ObjectType.eColonists, ObjectType.eOfficers, ObjectType.eEnlisted
                Case Else
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony's iObjTypeID not acceptable. Object Type: " & iObjTypeID & ", PlayerID: " & mlClientPlayer(lIndex))
                    Return
            End Select

            'transferring to the colony
            oColony = GetEpicaColony(lToID)

            If oColony Is Nothing = False Then
                If oColony.Owner.ObjectID <> mlClientPlayer(lIndex) AndAlso (oColony.Owner.ObjectID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eTransferCargo) = 0) Then
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony Owner is not Socket Owner! ToColony Owner: " & oColony.Owner.ObjectID & ", Socket: " & mlClientPlayer(lIndex))
                    Return
                End If
            Else
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony is nothing. PlayerID: " & mlClientPlayer(lIndex))
                Return
            End If

            Dim oFrom As Epica_Entity = Nothing
            If iFromTypeID = ObjectType.eUnit Then
                oFrom = GetEpicaUnit(lFromID)
            ElseIf iFromTypeID = ObjectType.eFacility Then
                oFrom = GetEpicaFacility(lFromID)
            Else
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony, FromType is not unit or facility. PlayerID: " & mlClientPlayer(lIndex))
                Return
            End If
            If oFrom Is Nothing Then
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony. From Entity not found. PlayerID: " & mlClientPlayer(lIndex))
                Return
            End If

            If oFrom.yProductionType = ProductionType.eMining Then
                LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony. From Entity Production is Mining. PlayerID: " & mlClientPlayer(lIndex))
            End If

            If (oFrom.CurrentStatus And elUnitStatus.eCargoBayOperational) = 0 Then
                'Indicate that the destination's cargo bay is inoperable 
                lQty = -1
                System.BitConverter.GetBytes(lQty).CopyTo(yData, 20)
                moClients(lIndex).SendData(yData)
                Return
            End If

            If oFrom.ObjTypeID = ObjectType.eUnit Then
                If CType(oFrom.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    If (CType(oFrom.ParentObject, Facility).CurrentStatus And elUnitStatus.eHangarOperational) = 0 Then
                        LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony. From Is Unit. Parent is Facility. Facility's hangar is inoperable. PlayerID: " & mlClientPlayer(lIndex))
                        Return
                    End If

                    If CType(oFrom.ParentObject, Facility).ParentColony Is Nothing OrElse CType(oFrom.ParentObject, Facility).ParentColony.ObjectID <> oColony.ObjectID Then
                        LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony. From Is Unit. Parent is Facility. Facility's colony is not dest colony. PlayerID: " & mlClientPlayer(lIndex))
                        Return
                    End If
                Else
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: ToColony. From is Unit that is not in facility. PlayerID: " & mlClientPlayer(lIndex))
                    Return
                End If
            Else
                If CType(oFrom, Facility).ParentColony Is Nothing OrElse CType(oFrom, Facility).ParentColony.ObjectID <> oColony.ObjectID Then
                    LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: FromColony. From Is Facility. Facility's colony is not dest colony. PlayerID: " & mlClientPlayer(lIndex))
                    Return
                End If
            End If

            'Does the source have the items? we will reduce them as we find them
            Dim lQtyFound As Int32 = 0
            Dim oDetailItem As Epica_GUID = Nothing

            If iObjTypeID = ObjectType.eColonists OrElse iObjTypeID = ObjectType.eEnlisted OrElse iObjTypeID = ObjectType.eOfficers Then
                'lQty = Math.Min(oFrom.oCargoContents(lCIdx).ObjectID + 1, lQty)
                For X As Int32 = 0 To oFrom.lCargoUB
                    If oFrom.lCargoIdx(X) <> -1 AndAlso oFrom.oCargoContents(X).ObjTypeID = iObjTypeID Then
                        Dim lTemp As Int32 = oFrom.oCargoContents(X).ObjectID + 1
                        lQtyFound += lTemp
                        If lQtyFound > lQty Then
                            Dim lTotal As Int32 = (lQtyFound - lQty)
                            oFrom.oCargoContents(X).ObjectID = lTotal - 1
                            oFrom.lCargoIdx(X) = oFrom.oCargoContents(X).ObjectID
                            lQtyFound = lQty
                        Else
                            oFrom.oCargoContents(X).ObjectID = -1
                            oFrom.lCargoIdx(X) = -1
                        End If

                        If lQtyFound = lQty Then Exit For
                    End If
                Next X
                oDetailItem = New Epica_GUID
                oDetailItem.ObjectID = lQtyFound
                oDetailItem.ObjTypeID = iObjTypeID
            Else
                For X As Int32 = 0 To oFrom.lCargoUB
                    If oFrom.lCargoIdx(X) <> -1 AndAlso oFrom.oCargoContents(X).ObjectID = lObjID AndAlso oFrom.oCargoContents(X).ObjTypeID = iObjTypeID Then
                        If iObjTypeID = ObjectType.eMineralCache Then
                            lQtyFound += CType(oFrom.oCargoContents(X), MineralCache).Quantity
                            If lQtyFound > lQty Then
                                Dim lTotal As Int32 = (lQtyFound - lQty)
                                oDetailItem = CType(oFrom.oCargoContents(X), MineralCache).oMineral
                                CType(oFrom.oCargoContents(X), MineralCache).Quantity = lTotal
                                lQtyFound = lQty
                            Else
                                oDetailItem = CType(oFrom.oCargoContents(X), MineralCache).oMineral
                                CType(oFrom.oCargoContents(X), MineralCache).Quantity = 0
                                oFrom.lCargoIdx(X) = -1
                            End If
                        ElseIf iObjTypeID = ObjectType.eComponentCache Then
                            lQtyFound += CType(oFrom.oCargoContents(X), ComponentCache).Quantity
                            If lQtyFound > lQty Then
                                Dim lTotal As Int32 = (lQtyFound - lQty)
                                oDetailItem = CType(oFrom.oCargoContents(X), ComponentCache).GetComponent
                                CType(oFrom.oCargoContents(X), ComponentCache).Quantity = lTotal
                                lQtyFound = lQty
                            Else
                                oDetailItem = CType(oFrom.oCargoContents(X), ComponentCache).GetComponent
                                CType(oFrom.oCargoContents(X), ComponentCache).Quantity = 0
                                oFrom.lCargoIdx(X) = -1
                            End If
                        End If
                        If lQtyFound = lQty Then Exit For
                    End If
                Next X
            End If

            If oDetailItem Is Nothing AndAlso lQtyFound <> 0 Then
                LogEvent(LogEventType.CriticalError, "HandleTransferContents: DetailItem is nothing")
                Return
            End If

            'If oFrom.ObjTypeID = ObjectType.eFacility AndAlso oFrom.yProductionType = ProductionType.eTradePost Then
            '    'k, verify that this item is not in a trade...
            '    Dim blSellQty As Int64 = goGTC.GetItemsBeingTraded(oFrom.ObjectID, lObjID, iObjTypeID)
            '    If blSellQty > Int32.MaxValue Then
            '        lQtyFound = 0
            '    ElseIf blSellQty > 0 Then
            '        lQtyFound -= CInt(blSellQty)
            '    End If
            'End If

            'ok, qty found indicates how much we have found...
            If lQtyFound <> lQty Then lQty = lQtyFound
            If lQty <> 0 Then
                'Go ahead and refund the from's cargo
                'oFrom.Cargo_Cap += lQty

                'Ok, now, attempt to put the cargo into the colony
                Dim lRemaining As Int32 = lQty
                If oDetailItem Is Nothing = False Then
                    If iObjTypeID = ObjectType.eColonists Then
                        oColony.Population += oDetailItem.ObjectID
                        lRemaining = 0
                    ElseIf iObjTypeID = ObjectType.eEnlisted Then
                        oColony.ColonyEnlisted += oDetailItem.ObjectID
                        lRemaining = 0
                    ElseIf iObjTypeID = ObjectType.eOfficers Then
                        oColony.ColonyOfficers += oDetailItem.ObjectID
                        lRemaining = 0
                    Else
                        lRemaining = oColony.AddObjectCaches(oDetailItem.ObjectID, oDetailItem.ObjTypeID, lQty, False)
                        If oDetailItem.ObjTypeID = ObjectType.eMineral AndAlso lRemaining <> lQty Then
                            oColony.Owner.CheckFirstContactWithMineral(oDetailItem.ObjectID)
                        End If
                    End If
                End If

                If lRemaining <> 0 Then
                    'Ok, remaining came back so the colony could not place everything... reduce our quantity transferred by remaining
                    lQty -= lRemaining

                    'add our cache
                    If oDetailItem Is Nothing = False Then
                        If oDetailItem.ObjTypeID = ObjectType.eMineral Then
                            If oFrom.AddMineralCacheToCargo(oDetailItem.ObjectID, lRemaining) Is Nothing = False Then
                                'we're going to be using cargo capacity so reduce it by remaining
                                'oFrom.Cargo_Cap -= lRemaining
                            End If
                        ElseIf oDetailItem.ObjTypeID = ObjectType.eColonists OrElse oDetailItem.ObjTypeID = ObjectType.eEnlisted OrElse oDetailItem.ObjTypeID = ObjectType.eOfficers Then
                            oFrom.AddPersonnelCacheToCargo(iObjTypeID, lRemaining)
                        Else
                            If oFrom.AddComponentCacheToCargo(oDetailItem.ObjectID, oDetailItem.ObjTypeID, lRemaining, CType(oDetailItem, Epica_Tech).Owner.ObjectID) Is Nothing = False Then
                                'we're going to be using cargo capacity so reduce it by remaining
                                'oFrom.Cargo_Cap -= lRemaining
                            End If
                        End If
                    Else : LogEvent(LogEventType.CriticalError, "Could not determine Detail Item to refund entity. Materials were lost! Entity: " & oFrom.ObjectID & ", " & oFrom.ObjTypeID & ". Object: " & lObjID & ", " & iObjTypeID)
                    End If
                End If
            End If
        Else
            'TODO: probably should allow entity to entity transfers in the future
            'for now, we will mark as possible cheat
            LogEvent(LogEventType.PossibleCheat, "HandleTransferContents: From/To Not Colony. PlayerID: " & mlClientPlayer(lIndex))
            Return
        End If

        'if we are here, the transfer worked (more or less) so we will send back the result
        System.BitConverter.GetBytes(lQty).CopyTo(yData, 20)
        moClients(lIndex).SendData(yData)
    End Sub
    Private Sub HandleTutorialGiveCredits(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer.yPlayerPhase <> eyPlayerPhase.eInitialPhase Then
            LogEvent(LogEventType.PossibleCheat, "Bannable Offense: Player Request Give Credits outside of Initial Phase: " & mlClientPlayer(lIndex))
            Return
        End If
        oPlayer.blCredits += 1000000
        oPlayer.PhaseOneMoneyClicks += 1
    End Sub
    Private Sub HandleTutorialProdFinish(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lStepID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'first, is the player in a tutorial phase
        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing Then Return
        If oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
            LogEvent(LogEventType.PossibleCheat, "Bannable Offense. HandleTutorialProdFinish received while player in full live. PlayerID: " & mlClientPlayer(lIndex))
            Return
        End If

        'ok, we are here, check the step....
        'Step		ID		TypeID
        '23			5		11
        '60			54		11
        '60			45		11
        '101		-1		7
        '155		-1		15
        '170		-1		15
        '170		-1		23
        '187		47		11
        '198		-1		45,44,23
        '208		3		34
        '208		-1		7
        '265		999		18

        'Select Case lStepID
        '	Case 23
        '		If lProdID <> 5 OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 63
        '		If (lProdID <> 54 AndAlso lProdID <> 45) OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 65
        '		If (lProdID <> 54 AndAlso lProdID <> 45) OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 103
        '		If lProdID <> 52 OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 109
        '		If iProdTypeID <> 7 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 164
        '		If iProdTypeID <> 15 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 174
        '		If iProdTypeID <> 23 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 179
        '		If iProdTypeID <> 15 AndAlso iProdTypeID <> 23 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 191
        '		If iProdTypeID <> 23 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 196
        '		If lProdID <> 47 OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 200
        '		If lProdID <> 53 OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 201
        '		If lProdID <> 51 OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 204
        '		If lProdID <> 49 OrElse iProdTypeID <> 11 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 209
        '		If iProdTypeID <> 23 AndAlso iProdTypeID <> 44 AndAlso iProdTypeID <> 45 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 220
        '		If iProdTypeID <> 7 AndAlso (lProdID <> 3 OrElse iProdTypeID <> 34) Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case 279
        '		If iProdTypeID <> 18 AndAlso lProdID <> 999 Then
        '			LogEvent(LogEventType.PossibleCheat, "HandleTutorialProdFinish unexpected prodguid for stepid. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '			Return
        '		End If
        '	Case Else
        '		LogEvent(LogEventType.PossibleCheat, "Unexpected HandleTutorialProdFinish for Step. StepID = " & lStepID & ", PlayerID = " & mlClientPlayer(lIndex))
        '		Return
        'End Select

        'Ok, now, we are here, so everything is ok
        If iProdTypeID = 18 Then
            'mission... we get a mission ID not a PM_ID...
            Dim lCurUB As Int32 = glPlayerMissionUB
            If glPlayerMissionIdx Is Nothing = False Then lCurUB = Math.Min(lCurUB, glPlayerMissionIdx.GetUpperBound(0)) Else lCurUB = -1
            For X As Int32 = 0 To lCurUB
                If glPlayerMissionIdx(X) <> -1 Then
                    Dim oPM As PlayerMission = goPlayerMission(X)
                    If oPM Is Nothing = False AndAlso oPM.oPlayer.ObjectID = oPlayer.ObjectID Then
                        If oPM.oMission Is Nothing = False Then 'AndAlso oPM.oMission.ObjectID = lProdID Then
                            goAgentEngine.FastForwardMissionEvent(oPM.PM_ID)
                            Exit For
                        End If
                    End If
                End If
            Next X
        Else
            'production
            If lProdID = -1 AndAlso iProdTypeID = -1 Then
                oPlayer.DeathBudgetEndTime = glCurrentCycle + 60
                oPlayer.DeathBudgetFundsRemaining = CInt(Math.Min(Int32.MaxValue, oPlayer.blCredits))
            End If
            FastForwardProduction(oPlayer.ObjectID, lProdID, iProdTypeID)

        End If

    End Sub
    Private Sub HandleUpdateGuildEvent(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2    'for msgcode
        Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If yData.Length > 10 Then
            'larger version, are we updating or creating?
            Dim oEvent As GuildEvent = Nothing
            If lEventID > 0 Then
                'updating
                oEvent = oPlayer.oGuild.GetEvent(lEventID)
            Else
                'creating
                oEvent = New GuildEvent()
            End If
            If oEvent Is Nothing = False Then
                With oEvent
                    ReDim .yTitle(49)
                    Array.Copy(yData, lPos, .yTitle, 0, 50) : lPos += 50
                    .dtStartsAt = Date.SpecifyKind(GetDateFromNumber(System.BitConverter.ToInt32(yData, lPos)), DateTimeKind.Utc) : lPos += 4
                    .lDuration = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .yEventType = yData(lPos) : lPos += 1
                    .ySendAlerts = yData(lPos) : lPos += 1
                    .yEventIcon = yData(lPos) : lPos += 1
                    .yMembersCanAccept = yData(lPos) : lPos += 1
                    .yRecurrence = yData(lPos) : lPos += 1
                    .dtPostedOn = Now.ToUniversalTime
                    Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    ReDim .yDetails(lLen - 1)
                    Array.Copy(yData, lPos, .yDetails, 0, lLen) : lPos += lLen
                    .lPostedBy = lPlayerID

                    If .SaveEvent(oPlayer.oGuild.ObjectID) = False Then Return
                End With

                If lEventID < 1 Then oPlayer.oGuild.AddEvent(oEvent)
                System.BitConverter.GetBytes(oEvent.EventID).CopyTo(yData, 2)
                ReDim Preserve yData(yData.GetUpperBound(0) + 8)
                System.BitConverter.GetBytes(lPlayerID).CopyTo(yData, lPos) : lPos += 4
                System.BitConverter.GetBytes(GetDateAsNumber(oEvent.dtPostedOn)).CopyTo(yData, lPos) : lPos += 4
                oPlayer.oGuild.SendMsgToGuildMembers(yData)
            End If
        Else
            'smaller version... event is always present
            Dim oEvent As GuildEvent = oPlayer.oGuild.GetEvent(lEventID)
            If oEvent Is Nothing Then Return

            Dim yAcceptance As Byte = yData(lPos) : lPos += 1
            'Ok, I am setting my acceptance
            If yAcceptance = 255 Then
                'cancel the event
                If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.DeleteEvents) OrElse oPlayer.ObjectID = oEvent.lPostedBy Then
                    oPlayer.oGuild.DeleteEvent(lEventID)
                    oPlayer.oGuild.SendMsgToGuildMembers(yData)
                Else
                    LogEvent(LogEventType.PossibleCheat, "Player attempting to cancel event without permission: " & mlClientPlayer(lIndex))
                    Return
                End If
            Else
                If oEvent.yMembersCanAccept = 0 Then Return

                Dim yForward(10) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildEvent).CopyTo(yForward, 0)
                System.BitConverter.GetBytes(lEventID).CopyTo(yForward, 2)
                yForward(6) = yAcceptance
                System.BitConverter.GetBytes(lPlayerID).CopyTo(yForward, 7)

                oEvent.SetPlayerAcceptance(lPlayerID, yAcceptance)
                oPlayer.oGuild.SendMsgToGuildMembers(yForward)
            End If
        End If
    End Sub
    Private Sub HandleUpdateGuildRank(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then Return
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        Dim lRankID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yChanges As Byte = yData(lPos) : lPos += 1

        If (yChanges And 1) <> 0 Then
            'move direction
            If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.CreateRanks) = False Then
                LogEvent(LogEventType.PossibleCheat, "Player attempting to move a rank without permission: " & mlClientPlayer(lIndex))
                Return
            End If
            Dim yMoveDir As Byte = yData(lPos) : lPos += 1
            'ok, now move the rank...
            If yMoveDir = 255 Then
                'remove the rank
                oPlayer.oGuild.RemoveRank(lRankID)
            ElseIf yMoveDir = 128 Then
                Dim oRank As New GuildRank()
                With oRank
                    .lRankID = -1
                    .lRankPermissions = 0
                    .lVoteStrength = 0
                    .ParentGuild = oPlayer.oGuild
                    .TaxRateFlat = 0
                    .TaxRatePercentage = 0
                    .TaxRatePercType = eyGuildTaxPercType.CashFlow
                    Dim lNextPos As Int32 = -1
                    For X As Int32 = 0 To oPlayer.oGuild.lRankUB
                        If oPlayer.oGuild.oRanks(X) Is Nothing = False Then
                            If oPlayer.oGuild.oRanks(X).yPosition >= lNextPos Then
                                lNextPos = CInt(oPlayer.oGuild.oRanks(X).yPosition) + 1
                            End If
                        End If
                    Next X
                    If lNextPos < 0 Then lNextPos = 0
                    If lNextPos > 255 Then lNextPos = 255
                    .yPosition = CByte(lNextPos)
                    ReDim .yRankName(19)
                    Array.Copy(yData, lPos, .yRankName, 0, 20) : lPos += 20

                    If (yChanges And 4) <> 0 Then lPos += 4
                    If (yChanges And 8) <> 0 Then
                        .TaxRateFlat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .TaxRatePercentage = yData(lPos) : lPos += 1
                        .TaxRatePercType = CType(yData(lPos), eyGuildTaxPercType) : lPos += 1
                    End If
                    If .SaveObject(oPlayer.oGuild.ObjectID) = False Then Return
                End With
                oPlayer.oGuild.AddRank(oRank)
                oPlayer.oGuild.SendMsgToGuildMembers(GetAddObjectMessage(oPlayer.oGuild, GlobalMessageCode.eAddObjectCommand))
                Return
            Else
                oPlayer.oGuild.MoveRank(lRankID, yMoveDir)
            End If
        End If
        If (yChanges And 2) <> 0 Then
            'new name
            If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ChangeRankNames) = False Then
                LogEvent(LogEventType.PossibleCheat, "Player attempting to change rank names without permission: " & mlClientPlayer(lIndex))
                Return
            End If
            Dim oRank As GuildRank = oPlayer.oGuild.GetRank(lRankID)
            If oRank Is Nothing = False Then
                ReDim oRank.yRankName(19)
                Array.Copy(yData, lPos, oRank.yRankName, 0, 20)
            End If
            lPos += 20
        End If
        If (yChanges And 4) <> 0 Then
            'new vote str
            If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ChangeRankVotingWeight) = False Then
                LogEvent(LogEventType.PossibleCheat, "Player attempting to change rank vote strength without permission: " & mlClientPlayer(lIndex))
                Return
            End If
            Dim lVoteStr As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim oRank As GuildRank = oPlayer.oGuild.GetRank(lRankID)
            If oRank Is Nothing = False Then oRank.lVoteStrength = lVoteStr
        End If
        If (yChanges And 8) <> 0 Then
            If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ChangeRankPermissions) = False Then
                LogEvent(LogEventType.PossibleCheat, "Player attempting to change taxes without permission: " & mlClientPlayer(lIndex))
                Return
            End If
            Dim lTaxRateFlat As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yTaxRatePerc As Byte = yData(lPos) : lPos += 1
            Dim yTaxRateType As Byte = yData(lPos) : lPos += 1
            Dim oRank As GuildRank = oPlayer.oGuild.GetRank(lRankID)
            If oRank Is Nothing = False Then
                oRank.TaxRateFlat = lTaxRateFlat
                oRank.TaxRatePercentage = yTaxRatePerc
                oRank.TaxRatePercType = CType(yTaxRateType, eyGuildTaxPercType)
            End If
        End If

        Dim oSaveRank As GuildRank = oPlayer.oGuild.GetRank(lRankID)
        If oSaveRank Is Nothing = False Then oSaveRank.SaveObject(oPlayer.oGuild.ObjectID)

        oPlayer.oGuild.SendMsgToGuildMembers(yData)

    End Sub
    Private Sub HandleUpdateGuildRecruitment(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ChangeRecruitment) = False Then
            LogEvent(LogEventType.PossibleCheat, "Player changing recruitment without permission: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim lPos As Int32 = 2   'for msgcode
        Dim iFlags As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lBillboardLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        oPlayer.oGuild.iRecruitFlags = CType(iFlags, eiRecruitmentFlags)
        ReDim oPlayer.oGuild.yBillboard(lBillboardLen - 1)
        Array.Copy(yData, lPos, oPlayer.oGuild.yBillboard, 0, lBillboardLen) : lPos += lBillboardLen

        oPlayer.oGuild.SendMsgToGuildMembers(yData)
    End Sub
    Private Sub HandleUpdateGuildRelNotes(ByRef yData() As Byte, ByVal lIndex As Int32)
        'Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        'If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        'Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        'If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        'Dim lPos As Int32 = 2	 'for msgcode
        'Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'Dim lTextLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Dim oRel As GuildRel = oPlayer.oGuild.GetRel(lEntityID, iEntityTypeID)
        'If oRel Is Nothing = False Then
        '	ReDim oRel.yNote(lTextLen - 1)
        '	Array.Copy(yData, lPos, oRel.yNote, 0, lTextLen) : lPos += lTextLen

        '	oPlayer.oGuild.SendMsgToGuildMembers(yData)
        'End If

    End Sub
    'OLD GUILD BASE STUFF
    'Private Sub HandleUpdateGuildTreasury(ByRef yData() As Byte, ByVal lIndex As Int32)
    '    Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
    '    If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

    '    Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
    '    If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

    '    If oPlayer.oGuild.lGuildHallID > 0 Then Return

    '    If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.BuildGuildBase) = False Then
    '        LogEvent(LogEventType.PossibleCheat, "Player attempting to place guildhall without permission: " & mlClientPlayer(lIndex))
    '        Return
    '    End If

    '    Dim lPos As Int32 = 2
    '    Dim lBuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iBuildTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '    Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iLocA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '    Dim oFacDef As FacilityDef = GetEpicaFacilityDef(lBuildID)
    '    If oFacDef Is Nothing Then Return

    '    Dim lIdx As Int32 = AddFacilityNoProducer(oFacDef, oPlayer, lEnvirID, iEnvirTypeID, lLocX, lLocZ, iLocA, eiBehaviorPatterns.eEngagement_Stand_Ground Or eiBehaviorPatterns.eTactics_Normal, eiTacticalAttrs.eArmedUnit, True)
    '    If lIdx <> -1 Then
    '        Dim oFac As Facility = goFacility(lIdx)
    '        If oFac Is Nothing = False Then
    '            If iEnvirTypeID = ObjectType.ePlanet Then
    '                CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
    '            ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
    '                CType(oFac.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
    '            End If

    '            oPlayer.oGuild.lGuildHallID = oFac.ObjectID
    '            oPlayer.oGuild.SendMsgToGuildMembers(GetAddObjectMessage(oPlayer.oGuild, GlobalMessageCode.eAddObjectCommand))
    '        End If
    '    End If
    'End Sub
    Private Sub HandleUpdateGuildTreasury(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ViewBankLog) = False Then
            LogEvent(LogEventType.PossibleCheat, "Player attempting to view guild log without permission: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim oGuild As Guild = oPlayer.oGuild
        '28 bytes each item...
        'show no more than 100 items
        Dim lCnt As Int32 = Math.Min(100, (oGuild.lBankLogUB + 1))
        'If lCnt < 1 Then Return

        Dim yMsg(13 + (lCnt * 24)) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildTreasury).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oGuild.blLastTaxIncome).CopyTo(yMsg, lPos) : lPos += 8

        For X As Int32 = 0 To Math.Min(99, oGuild.lBankLogUB)
            With oGuild.oBankLog(X)
                System.BitConverter.GetBytes(.lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lTransDate).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.blAmount).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(.blBalance).CopyTo(yMsg, lPos) : lPos += 8
            End With
        Next X
        moClients(lIndex).SendData(yMsg)

    End Sub
    Private Sub HandleUpdateMOTD(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If oPlayer.oGuild Is Nothing = False Then
                If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ChangeMOTD) = True Then
                    Dim yTemp(199) As Byte
                    Array.Copy(yData, 2, yTemp, 0, 200)
                    oPlayer.oGuild.yMOTD = yTemp
                    oPlayer.oGuild.SendMsgToGuildMembers(yData)
                Else
                    LogEvent(LogEventType.PossibleCheat, "Player attempting to change MOTD and lacks permission: " & mlClientPlayer(lIndex))
                End If
            End If
        End If

    End Sub
    Private Sub HandleUpdatePlayerCustomTitle(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lTitlePerm As Int32 = System.BitConverter.ToInt32(yData, 2)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) Then Return

        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing = False Then
            If lTitlePerm = -1 Then
                oPlayer.yCustomTitle = oPlayer.yPlayerTitle
            Else
                If (oPlayer.lCustomTitlePermission And lTitlePerm) <> 0 Then
                    Select Case lTitlePerm
                        Case elCustomRankPermissions.Arbiter
                            oPlayer.yCustomTitle = eyCustomRank.Governor Or eyCustomRank.DiplomacyShift
                        Case elCustomRankPermissions.Broker
                            oPlayer.yCustomTitle = eyCustomRank.Overseer Or eyCustomRank.TraderShift
                        Case elCustomRankPermissions.Chancellor
                            oPlayer.yCustomTitle = eyCustomRank.King Or eyCustomRank.DiplomacyShift
                        Case elCustomRankPermissions.ChiefBroker
                            oPlayer.yCustomTitle = eyCustomRank.Baron Or eyCustomRank.TraderShift
                        Case elCustomRankPermissions.ChiefScientist
                            oPlayer.yCustomTitle = eyCustomRank.Baron Or eyCustomRank.ResearcherShift
                        Case elCustomRankPermissions.CommerceCzar
                            oPlayer.yCustomTitle = eyCustomRank.Emperor Or eyCustomRank.TraderShift
                        Case elCustomRankPermissions.Counselor
                            oPlayer.yCustomTitle = eyCustomRank.Overseer Or eyCustomRank.DiplomacyShift
                        Case elCustomRankPermissions.Diplomat
                            oPlayer.yCustomTitle = eyCustomRank.Magistrate Or eyCustomRank.DiplomacyShift
                        Case elCustomRankPermissions.Explorer
                            oPlayer.yCustomTitle = eyCustomRank.Magistrate Or eyCustomRank.ResearcherShift
                        Case elCustomRankPermissions.HighSenator
                            oPlayer.yCustomTitle = eyCustomRank.Baron Or eyCustomRank.DiplomacyShift
                        Case elCustomRankPermissions.Inquisitor
                            oPlayer.yCustomTitle = eyCustomRank.Duke Or eyCustomRank.ResearcherShift
                        Case elCustomRankPermissions.MasterMerchant
                            oPlayer.yCustomTitle = eyCustomRank.King Or eyCustomRank.TraderShift
                        Case elCustomRankPermissions.MasterScientist
                            oPlayer.yCustomTitle = eyCustomRank.Overseer Or eyCustomRank.ResearcherShift
                        Case elCustomRankPermissions.Merchant
                            oPlayer.yCustomTitle = eyCustomRank.Governor Or eyCustomRank.TraderShift
                        Case elCustomRankPermissions.Preeminence
                            oPlayer.yCustomTitle = eyCustomRank.King Or eyCustomRank.ResearcherShift
                        Case elCustomRankPermissions.Scientist
                            oPlayer.yCustomTitle = eyCustomRank.Governor Or eyCustomRank.ResearcherShift
                        Case elCustomRankPermissions.Senator
                            oPlayer.yCustomTitle = eyCustomRank.Duke Or eyCustomRank.DiplomacyShift
                        Case elCustomRankPermissions.SupremeChancellor
                            oPlayer.yCustomTitle = eyCustomRank.Emperor Or eyCustomRank.DiplomacyShift
                        Case elCustomRankPermissions.TradeLord
                            oPlayer.yCustomTitle = eyCustomRank.Duke Or eyCustomRank.TraderShift
                        Case elCustomRankPermissions.Trader
                            oPlayer.yCustomTitle = eyCustomRank.Magistrate Or eyCustomRank.TraderShift
                        Case elCustomRankPermissions.Transcendent
                            oPlayer.yCustomTitle = eyCustomRank.Emperor Or eyCustomRank.ResearcherShift
                    End Select
                    Dim yOutMsg(10) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerCustomTitle).CopyTo(yOutMsg, 0)
                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yOutMsg, 2)
                    System.BitConverter.GetBytes(oPlayer.lCustomTitlePermission).CopyTo(yOutMsg, 6)
                    yOutMsg(10) = oPlayer.yCustomTitle
                    moClients(lIndex).SendData(yOutMsg)

                    ReDim yOutMsg(6)
                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerCustomTitle).CopyTo(yOutMsg, 0)
                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yOutMsg, 2)
                    yOutMsg(6) = oPlayer.yCustomTitle
                    For X As Int32 = 0 To oPlayer.PlayerRelUB
                        Dim oRel As PlayerRel = oPlayer.GetPlayerRelByIndex(X)
                        If oRel Is Nothing = False Then
                            If oRel.oPlayerRegards Is Nothing = False AndAlso oRel.oThisPlayer Is Nothing = False Then
                                If oRel.oPlayerRegards.ObjectID = oPlayer.ObjectID Then
                                    oRel.oThisPlayer.SendPlayerMessage(yOutMsg, False, AliasingRights.eViewDiplomacy)
                                End If
                            End If
                        End If
                    Next X

                Else
                    LogEvent(LogEventType.PossibleCheat, "Player trying to set title to something they dont have: " & mlClientPlayer(lIndex))
                End If
            End If
        End If

    End Sub
    Private Sub HandleUpdatePlayerTutorialStep(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lStepID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(lIndex))
        If oPlayer Is Nothing = False Then
            If oPlayer.lTutorialStep > lStepID Then
                LogEvent(LogEventType.PossibleCheat, "HandleUpdatePlayerTutorialStep: StepID is less than last reported. Player: " & mlClientPlayer(lIndex))
            Else
                oPlayer.lTutorialStep = lStepID

                'Ok, new idea... if the step id = 49, then the player is to 
                If lStepID = 51 Then
                    '   auto self destruct (remove all units, facilities, etc...)
                    oPlayer.PlayerIsDead = True

                    Dim yForward(7) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yForward, 0)
                    System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yForward, 2)
                    yForward(6) = 255
                    yForward(7) = 0
                    moOperator.SendData(yForward)

                    '   auto change envirs to aurelium
                    '   start the aurelium part of the tutorial
                    Return
                End If

                If lStepID = 231 Then oPlayer.blCredits += 20000000
                If lStepID = 232 Then
                    AddToQueue(glCurrentCycle + 30, QueueItemType.eGenerateNewbieAgent, oPlayer.ObjectID, ObjectType.ePlayer, -1, -1, 0, 0, 0, 0)
                End If
                If lStepID = 233 Then oPlayer.blCredits = Math.Max(oPlayer.blCredits + 5000000, 5000000)
                If lStepID = 293 Then oPlayer.blMaxPopulation = 1000000
                If lStepID = 294 Then oPlayer.blCredits += 100000000
                If lStepID = 297 Then
                    oPlayer.blCredits += 100000000
                    AureliusAI.SpawnNextWave(oPlayer.ObjectID, oPlayer.ObjectID + 500000000)
                End If
                If lStepID = 314 Then
                    Dim lSeconds As Int32 = oPlayer.TotalPlayTime
                    If oPlayer.lConnectedPrimaryID > -1 Then
                        lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
                    End If
                    If oPlayer.PlayedTimeInTutorialOne <> Int32.MinValue Then lSeconds += oPlayer.PlayedTimeInTutorialOne

                    'oPlayer.PlayedTimeWhenTimerStarted = lSeconds
                    'Dim lNumberOfSeconds As Int32 = 14400
                    'AddToQueue(glCurrentCycle + (lNumberOfSeconds * 30), QueueItemType.eTutorialFourHourTimerExpire, oPlayer.ObjectID, oPlayer.ObjTypeID, -1, -1, 0, 0, 0, 0)
                End If
                If lStepID = 281 Then
                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                       "We have found the location of the pirate factory! Your agent was also successful in causing a malfunction in the factory's machinery which resulted in an explosion doing massive damage to the factory. The waypoint is attached to this email. For the honor of your empire!", _
                       "Pirate Factory", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then
                        oPC.AddEmailAttachment(1, oPlayer.ObjectID + 500000000, ObjectType.ePlanet, -15000, -16600, "Factory Loc")
                        oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                    oPC = Nothing
                End If
                If lStepID = 285 Then
                    For X As Int32 = 0 To oPlayer.mlColonyUB
                        If oPlayer.mlColonyIdx(X) > -1 Then
                            If glColonyIdx(oPlayer.mlColonyIdx(X)) = oPlayer.mlColonyID(X) Then
                                Dim oColony As Colony = goColony(oPlayer.mlColonyIdx(X))
                                If oColony Is Nothing = False Then
                                    For Y As Int32 = 0 To oColony.ChildrenUB
                                        Dim oFac As Facility = oColony.oChildren(Y)
                                        If oFac Is Nothing = False AndAlso oColony.oChildren(Y).yProductionType = ProductionType.eCommandCenterSpecial Then
                                            oFac.Structure_HP = oFac.EntityDef.Structure_MaxHP \ 10
                                            Dim yMsg(15) As Byte
                                            System.BitConverter.GetBytes(GlobalMessageCode.eRepairCompleted).CopyTo(yMsg, 0)
                                            oFac.GetGUIDAsString.CopyTo(yMsg, 2)

                                            System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yMsg, 8)
                                            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 12)

                                            If oFac.ParentObject Is Nothing = False Then
                                                With CType(oFac.ParentObject, Epica_GUID)
                                                    If .ObjTypeID = ObjectType.ePlanet Then
                                                        If CType(oFac.ParentObject, Planet).oDomain Is Nothing = False Then CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                                                    ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                                                        If CType(oFac.ParentObject, SolarSystem).oDomain Is Nothing = False Then CType(oFac.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                                                    End If
                                                End With
                                            End If
                                            Exit For
                                        End If
                                    Next Y
                                    'Exit For
                                End If
                            End If
                        End If
                    Next X
                End If
            End If
        End If
    End Sub
    Private Sub HandleUpdatePlayerVote(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlClientPlayer(lIndex) <> mlAliasedAs(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim yVoteType As Byte = yData(lPos) : lPos += 1
        If yVoteType = 0 Then
            'guild
            If oPlayer.oGuild Is Nothing Then Return

            Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yVoteVal As Byte = yData(lPos) : lPos += 1
            Dim oVote As GuildVote = oPlayer.oGuild.GetVote(lProposalID)
            If oVote Is Nothing = False Then
                oVote.SetMemberVote(lPlayerID, CType(yVoteVal, eyVoteValue))
            End If
        Else
            Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yVoteVal As eyVoteValue = CType(yData(lPos), eyVoteValue) : lPos += 1
            System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yData, lPos) : lPos += 4
            moOperator.SendData(yData)
            'Senate.SetPlayerVote(lProposalID, yVoteVal, lPlayerID)
        End If
    End Sub
    Private Sub HandleUpdateRankPermission(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) > 0 AndAlso mlAliasedAs(lIndex) <> lPlayerID Then Return

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing OrElse oPlayer.oGuild Is Nothing Then Return

        If oPlayer.oGuild.RankHasPermission(oPlayer.lGuildRankID, RankPermissions.ChangeRankPermissions) = False Then
            LogEvent(LogEventType.PossibleCheat, "Player attempting to change rank permissions without permission: " & mlClientPlayer(lIndex))
            Return
        End If

        Dim lPos As Int32 = 2   'for msgcode
        Dim lRankID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lValueChangeCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oRank As GuildRank = oPlayer.oGuild.GetRank(lRankID)
        If oRank Is Nothing = False Then
            For X As Int32 = 0 To lValueChangeCnt - 1
                Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If lValue < 0 Then
                    'ok, this means we are removing the value
                    lValue = Math.Abs(lValue)
                    If (oRank.lRankPermissions And lValue) <> 0 Then oRank.lRankPermissions = oRank.lRankPermissions Xor lValue
                Else
                    'ok, this means we are adding the value
                    oRank.lRankPermissions = oRank.lRankPermissions Or lValue
                End If
            Next X
            oPlayer.oGuild.SendMsgToGuildMembers(yData)
        End If
    End Sub
    Private Sub HandleUpdateRouteStatus(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oUnit As Unit = GetEpicaUnit(lUnitID)
        If oUnit Is Nothing Then Return
        If oUnit.Owner.ObjectID = mlClientPlayer(lIndex) OrElse (oUnit.Owner.ObjectID = mlAliasedAs(lIndex) AndAlso (mlAliasedRights(lIndex) And AliasingRights.eMoveUnits) <> 0) Then
            If lValue = Int32.MaxValue Then
                'begin
                oUnit.lCurrentRouteIdx = -1
                oUnit.bRoutePaused = False
                oUnit.bRunRouteOnce = False
                oUnit.ProcessNextRouteItem()
            ElseIf lValue = Int32.MinValue Then
                'force next
                oUnit.bRoutePaused = False
                oUnit.ProcessNextRouteItem()
            ElseIf lValue = -2 Then
                'pause
                oUnit.bRoutePaused = Not oUnit.bRoutePaused
                If oUnit.bRoutePaused = True Then
                    System.BitConverter.GetBytes(-2I).CopyTo(yData, 6)
                Else
                    System.BitConverter.GetBytes(-3I).CopyTo(yData, 6)
                    oUnit.CurrentRouteItemAction()
                End If
            ElseIf lValue = -4 Then
                oUnit.lCurrentRouteIdx = -1
                oUnit.bRoutePaused = False
                oUnit.bRunRouteOnce = True
                oUnit.ProcessNextRouteItem()
            Else

                'Ok, adding/update a route location
                Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                'Now... update our unit
                If oUnit.lRouteUB >= lValue Then
                    With oUnit.uRoute(lValue)
                        .oDest = CType(GetEpicaObject(lEnvirID, iEnvirTypeID), Epica_GUID)
                        .lLocX = lLocX
                        .lLocZ = lLocZ
                        If .oDest.ObjTypeID = ObjectType.eFacility Then
                            If CType(.oDest, Facility).yProductionType = ProductionType.eMining Then
                                .oLoadItem = Nothing
                                .SetLoadItem(-1, -1, oUnit.Owner, .yExtraFlags)
                            End If
                        End If
                    End With
                Else
                    Dim uNewItem As RouteItem
                    With uNewItem
                        .lLocX = lLocX
                        .lLocZ = lLocZ
                        .lOrderNum = oUnit.lRouteUB + 1
                        .oDest = CType(GetEpicaObject(lEnvirID, iEnvirTypeID), Epica_GUID)
                        If .oDest.ObjTypeID = ObjectType.eFacility Then
                            If CType(.oDest, Facility).yProductionType = ProductionType.eMining Then
                                .oLoadItem = Nothing
                                .SetLoadItem(-1, -1, oUnit.Owner, .yExtraFlags)
                            End If
                        End If
                        .oLoadItem = Nothing
                    End With
                    oUnit.AddRouteItem(uNewItem)
                End If

                moClients(lIndex).SendData(yData)
            End If
            moClients(lIndex).SendData(yData)
        Else : LogEvent(LogEventType.PossibleCheat, "HandleUpdateRouteStatus: Owner Mismatch. PlayerID: " & mlClientPlayer(lIndex))
        End If
    End Sub
    Private Sub HandleUpdateSlotStates(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lActualPlayerID As Int32 = mlClientPlayer(lIndex)
        If mlAliasedAs(lIndex) <> mlClientPlayer(lIndex) AndAlso mlAliasedAs(lIndex) > 0 Then
            If (mlAliasedRights(lIndex) And AliasingRights.eAlterDiplomacy) = 0 Then
                LogEvent(LogEventType.PossibleCheat, "HandleUpdateSlotStates Player Mismatch: " & mlClientPlayer(lIndex))
                Return
            End If
            lActualPlayerID = mlAliasedAs(lIndex)
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lActualPlayerID)
        If oPlayer Is Nothing Then Return
        If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount Then
            moClients(lIndex).SendData(CreateChatMsg(-1, "Trial accounts cannot use factions.", ChatMessageType.eSysAdminMessage))
            Return
        End If

        Dim lPos As Int32 = 2 'for msgcode
        Dim yType As Byte = yData(lPos) : lPos += 1
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yState As Byte = yData(lPos) : lPos += 1

        'if the ID is -1, simply return the UpdateSlotState msg...
        If lID > 0 Then
            'if ytype is a 0 then its a slot, otherwise its a faction
            If yType = 0 Then
                'slot... ok, our action (yState) - I can add or remove
                Select Case yState
                    Case eySlotState.ForceRemove
                        For X As Int32 = 0 To 4
                            If oPlayer.lSlotID(X) = lID Then
                                oPlayer.lSlotID(X) = -1
                                Exit For
                            End If
                        Next X
                        Dim oTmp As Player = GetEpicaPlayer(lID)
                        If oTmp Is Nothing = False Then
                            For X As Int32 = 0 To 2
                                If oTmp.lFactionID(X) = oPlayer.ObjectID Then
                                    oTmp.lFactionID(X) = -1
                                End If
                            Next X

                            'send other player notice
                            If (oTmp.iInternalEmailSettings And eEmailSettings.eFactionUpdates) <> 0 Then
                                Dim oPC As PlayerComm = oTmp.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oPlayer.sPlayerNameProper & " has removed you from their faction.", "Faction Removal", oTmp.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                If oPC Is Nothing = False Then oTmp.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If

                            oPlayer.ReverifySlots()
                            oTmp.ReverifySlots()
                            'If oTmp.oSocket Is Nothing = False Then
                            '	oTmp.oSocket.SendData(oTmp.GetSlotStateMsg)
                            'End If
                        End If
                        oTmp = Nothing
                    Case eySlotState.Unaccepted 'add
                        Dim oTmp As Player = GetEpicaPlayer(lID)
                        If oTmp Is Nothing = False Then

                            If oTmp.yPlayerTitle > oPlayer.yPlayerTitle Then Return
                            If oTmp.AccountStatus <> AccountStatusType.eActiveAccount Then
                                moClients(lIndex).SendData(CreateChatMsg(-1, "Trial accounts cannot use factions.", ChatMessageType.eSysAdminMessage))
                                Return
                            End If

                            Dim bGood As Boolean = False
                            For X As Int32 = 0 To 2
                                If oTmp.lFactionID(X) < 1 Then
                                    bGood = True
                                    Exit For
                                End If
                            Next X

                            If bGood = True Then
                                bGood = False
                                For X As Int32 = 0 To 4
                                    If oPlayer.lSlotID(X) < 1 Then
                                        bGood = True
                                        Exit For
                                    End If
                                Next X
                                If bGood = False Then Return
                            End If

                            If bGood = False Then
                                If (oPlayer.iInternalEmailSettings And eEmailSettings.eFactionUpdates) <> 0 Then
                                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your request for faction with " & oTmp.sPlayerNameProper & " was received, however, " & oTmp.sPlayerNameProper & " does not have faction slots remaining.", "Faction Request Denied", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                    If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                End If
                            Else
                                For X As Int32 = 0 To 4
                                    If oPlayer.lSlotID(X) < 1 Then
                                        oPlayer.lSlotID(X) = lID
                                        oPlayer.ySlotState(X) = eySlotState.Unaccepted
                                        Exit For
                                    End If
                                Next X
                                For X As Int32 = 0 To 2
                                    If oTmp.lFactionID(X) < 1 Then
                                        oTmp.lFactionID(X) = oPlayer.ObjectID
                                        Exit For
                                    End If
                                Next X


                                'ok, send a msg to oTmp
                                If (oTmp.iInternalEmailSettings And eEmailSettings.eFactionUpdates) <> 0 Then
                                    Dim oPC As PlayerComm = oTmp.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oPlayer.sPlayerNameProper & " has requested your participation in a faction. To view the request, go to the Diplomacy screen and click on the Factions button.", "Faction Requested", oPlayer.ObjectID, GetDateAsNumber(Now), False, oTmp.sPlayerNameProper, Nothing)
                                    If oPC Is Nothing = False Then oTmp.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                End If

                                oPlayer.ReverifySlots()
                                oTmp.ReverifySlots()
                                'If oTmp.oSocket Is Nothing = False Then
                                '	oTmp.oSocket.SendData(oTmp.GetSlotStateMsg)
                                'End If
                            End If
                        End If
                    Case Else
                        LogEvent(LogEventType.PossibleCheat, "Unknown State (" & yState & ") in HandleUpdateSlotStates. Player: " & mlClientPlayer(lIndex))
                End Select
            Else
                'faction - i can accept or reject
                Select Case yState
                    Case eySlotState.Accepted
                        Dim bGood As Boolean = False
                        For X As Int32 = 0 To 2
                            If oPlayer.lFactionID(X) = lID Then
                                'Now, let's get the other player
                                Dim oTmp As Player = GetEpicaPlayer(lID)
                                If oTmp Is Nothing = False Then
                                    If oTmp.AccountStatus <> AccountStatusType.eActiveAccount Then
                                        moClients(lIndex).SendData(CreateChatMsg(-1, "Trial accounts cannot use factions.", ChatMessageType.eSysAdminMessage))
                                        Return
                                    End If
                                    For Y As Int32 = 0 To 4
                                        If oTmp.lSlotID(Y) = oPlayer.ObjectID Then
                                            'ok... got it
                                            bGood = True

                                            oTmp.ySlotState(Y) = oTmp.ySlotState(Y) Or eySlotState.Accepted

                                            'Now, notify oTmp
                                            If (oTmp.iInternalEmailSettings And eEmailSettings.eFactionUpdates) <> 0 Then
                                                Dim oPC As PlayerComm = oTmp.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oPlayer.sPlayerNameProper & " has accepted your faction request.", "Faction Accepted", oPlayer.ObjectID, GetDateAsNumber(Now), False, oTmp.sPlayerNameProper, Nothing)
                                                If oPC Is Nothing = False Then oTmp.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                            End If

                                            oPlayer.ReverifySlots()
                                            oTmp.ReverifySlots()
                                            'If oTmp.oSocket Is Nothing = False Then
                                            '	oTmp.oSocket.SendData(oTmp.GetSlotStateMsg)
                                            'End If

                                            Exit For
                                        End If
                                    Next Y
                                End If

                                Exit For
                            End If
                        Next X
                        If bGood = False Then
                            LogEvent(LogEventType.PossibleCheat, "HandleUpdateSlotStates: Player Faction does not exist in Accept. " & mlClientPlayer(lIndex))
                        End If
                    Case eySlotState.ForceRemove
                        Dim bGood As Boolean = False
                        For X As Int32 = 0 To 2
                            If oPlayer.lFactionID(X) = lID Then
                                'Now, let's get the other player
                                oPlayer.lFactionID(X) = -1
                                Dim oTmp As Player = GetEpicaPlayer(lID)
                                If oTmp Is Nothing = False Then
                                    For Y As Int32 = 0 To 4
                                        If oTmp.lSlotID(Y) = oPlayer.ObjectID Then
                                            'ok... got it
                                            bGood = True

                                            oTmp.lSlotID(Y) = -1

                                            'Now, notify oTmp
                                            If (oTmp.iInternalEmailSettings And eEmailSettings.eFactionUpdates) <> 0 Then
                                                Dim oPC As PlayerComm = oTmp.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oPlayer.sPlayerNameProper & " has rejected your faction request.", "Faction Rejected", oPlayer.ObjectID, GetDateAsNumber(Now), False, oTmp.sPlayerNameProper, Nothing)
                                                If oPC Is Nothing = False Then oTmp.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                            End If

                                            oPlayer.ReverifySlots()
                                            oTmp.ReverifySlots()
                                            'If oTmp.oSocket Is Nothing = False Then
                                            '	oTmp.oSocket.SendData(oTmp.GetSlotStateMsg)
                                            'End If

                                            Exit For
                                        End If
                                    Next Y
                                End If

                                Exit For
                            End If
                        Next X
                        If bGood = False Then
                            LogEvent(LogEventType.PossibleCheat, "HandleUpdateSlotStates: Player Faction does not exist in Reject. " & mlClientPlayer(lIndex))
                        End If

                    Case Else
                        LogEvent(LogEventType.PossibleCheat, "Unknown State (" & yState & ") in HandleUpdateSlotStates. Player: " & mlClientPlayer(lIndex))
                End Select
            End If

        End If

        moClients(lIndex).SendData(oPlayer.GetSlotStateMsg)
    End Sub
#End Region

End Class
