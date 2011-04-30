Option Strict On

'Manages a DirectTrade - non global market, basically a trade agreement
Public Class DirectTrade
    Inherits Epica_GUID

    Public Enum eTradeStateValues As Byte
        Proposal = 0
        InNegotiation = 1       'synonomous with OPEN
        InVerification = 2      'both players have accepted and the server needs to verify the trade parameters
        InProgress = 4          'trade is in progress, there should be 2 queue action items one at 1/2 time and one at end
        TradeCompleted = 8      'trade is finished
        Player1Accepted = 16
        Player2Accepted = 32

        TradeRejected = 255
    End Enum

    Public Enum eFailureReason As Byte
        NoFailureReason = 0
        Player1DestNotFound = 1
        Player1SourceNotFound = 2
        Player2DestNotFound = 3
        Player2SourceNotFound = 4
        Player1MissingItems = 5
        Player2MissingItems = 6
		FacilityTradeDifferentEnvirs = 7
		Player1AgentsFull = 8
		Player2AgentsFull = 9
    End Enum

    Public TradeState As Byte

    Public FailureReason As Byte

    'The next two values only matter when TradeState is InProgress
    Public TradeCycles As Int32         'number of cycles for the trade to complete
    Public CyclesRemaining As Int32     'guesstimate for cycles remaining for the trade to complete
    Public TradeStarted As Int32        'date in Int32 format for when the trade was initiated

    Public oPlayer1TP As TradePlayer
    Public oPlayer2TP As TradePlayer

    Public Function GetObjAsString() As Byte()
        Dim yMsg() As Byte
        Dim lPos As Int32 = 0

        Dim lP1ItemCnt As Int32 = 0
        Dim lP2ItemCnt As Int32 = 0
        Dim lP1NoteLen As Int32 = 0
        Dim lP2NoteLen As Int32 = 0

        For X As Int32 = 0 To oPlayer1TP.lItemUB
            If oPlayer1TP.lItemIdx(X) <> -1 Then lP1ItemCnt += 1
        Next X
        For X As Int32 = 0 To oPlayer2TP.lItemUB
            If oPlayer2TP.lItemIdx(X) <> -1 Then lP2ItemCnt += 1
        Next X

        If oPlayer1TP.Notes Is Nothing = False Then lP1NoteLen = oPlayer1TP.Notes.Length
        If oPlayer2TP.Notes Is Nothing = False Then lP2NoteLen = oPlayer2TP.Notes.Length

		ReDim yMsg(47 + lP1NoteLen + lP2NoteLen + (lP1ItemCnt * 18) + (lP2ItemCnt * 18))

        Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(oPlayer1TP.PlayerID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oPlayer2TP.PlayerID).CopyTo(yMsg, lPos) : lPos += 4
        'System.BitConverter.GetBytes(TradeStarted).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = TradeState : lPos += 1
        System.BitConverter.GetBytes(TradeCycles).CopyTo(yMsg, lPos) : lPos += 4

        'If TradeStarted <> 0 Then
        '    Dim lTemp As Int32 = CInt(Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime.Subtract(Date.SpecifyKind(GetDateFromNumber(TradeStarted), DateTimeKind.Utc)).TotalMilliseconds / 30.0F)
        '    CyclesRemaining = TradeCycles - lTemp
        'End If
        ''		CyclesRemaining = TradeCycles - (glCurrentCycle - TradeStarted)
        'System.BitConverter.GetBytes(CyclesRemaining).CopyTo(yMsg, lPos) : lPos += 4
        If TradeStarted <> 0 Then
            Dim lTemp As Int32 = CInt(Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime.Subtract(Date.SpecifyKind(GetDateFromNumber(TradeStarted), DateTimeKind.Utc)).TotalSeconds)
            System.BitConverter.GetBytes(lTemp * 30).CopyTo(yMsg, lPos) : lPos += 4
        Else
            System.BitConverter.GetBytes(TradeStarted).CopyTo(yMsg, lPos) : lPos += 4
        End If


        yMsg(lPos) = FailureReason : lPos += 1

        'Player 1 is first...
        With oPlayer1TP
            'System.BitConverter.GetBytes(.DestID).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(.SourceID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(.lTradePostID).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(lP1NoteLen).CopyTo(yMsg, lPos) : lPos += 4
            If lP1NoteLen > 0 Then
                .Notes.CopyTo(yMsg, lPos) : lPos += .Notes.Length
            End If

            'Now the items
            System.BitConverter.GetBytes(lP1ItemCnt).CopyTo(yMsg, lPos) : lPos += 4
            For X As Int32 = 0 To .lItemUB
                If .lItemIdx(X) <> -1 Then
                    System.BitConverter.GetBytes(.oItems(X).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.oItems(X).ObjTypeID).CopyTo(yMsg, lPos) : lPos += 2
					System.BitConverter.GetBytes(.oItems(X).Quantity).CopyTo(yMsg, lPos) : lPos += 8

					Dim lTmpID As Int32 = .oItems(X).lExtendedID
					Dim iTmpTypeID As Int16 = .oItems(X).ObjTypeID
					If iTmpTypeID = ObjectType.eMineralCache Then
						Dim oCache As MineralCache = GetEpicaMineralCache(.oItems(X).ObjectID)
						If oCache Is Nothing = False AndAlso oCache.oMineral Is Nothing = False Then lTmpID = oCache.oMineral.ObjectID
					End If
					System.BitConverter.GetBytes(lTmpID).CopyTo(yMsg, lPos) : lPos += 4
                End If
            Next X
        End With

        'Now, player 2
        With oPlayer2TP
            'System.BitConverter.GetBytes(.DestID).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(.SourceID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(.lTradePostID).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(lP2NoteLen).CopyTo(yMsg, lPos) : lPos += 4
            If lP2NoteLen > 0 Then
                .Notes.CopyTo(yMsg, lPos) : lPos += .Notes.Length
            End If

            'Now the items
            System.BitConverter.GetBytes(lP2ItemCnt).CopyTo(yMsg, lPos) : lPos += 4
            For X As Int32 = 0 To .lItemUB
                If .lItemIdx(X) <> -1 Then
                    System.BitConverter.GetBytes(.oItems(X).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.oItems(X).ObjTypeID).CopyTo(yMsg, lPos) : lPos += 2
					System.BitConverter.GetBytes(.oItems(X).Quantity).CopyTo(yMsg, lPos) : lPos += 8

					Dim lTmpID As Int32 = .oItems(X).lExtendedID
					Dim iTmpTypeID As Int16 = .oItems(X).ObjTypeID
					If iTmpTypeID = ObjectType.eMineralCache Then
						Dim oCache As MineralCache = GetEpicaMineralCache(.oItems(X).ObjectID)
						If oCache Is Nothing = False AndAlso oCache.oMineral Is Nothing = False Then lTmpID = oCache.oMineral.ObjectID
					End If
					System.BitConverter.GetBytes(lTmpID).CopyTo(yMsg, lPos) : lPos += 4
                End If
            Next X
        End With

        Return yMsg

    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID < 1 Then
                'INSERT
                sSQL = "INSERT INTO tblTrade (TradeStarted, TradeState, TradeCycles, CyclesRemaining, FailureReason) VALUES (" & _
                  TradeStarted & ", " & TradeState & ", " & TradeCycles & ", " & CyclesRemaining & ", " & FailureReason & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblTrade SET TradeStarted = " & TradeStarted & ", TradeState = " & TradeState & _
                  ", TradeCycles = " & TradeCycles & ", CyclesRemaining = " & CyclesRemaining & ", FailureReason = " & _
                  FailureReason & " WHERE TradeID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID < 1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(TradeID) FROM tblTrade WHERE TradeStarted = " & TradeStarted
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            oComm = Nothing

            Me.ObjTypeID = ObjectType.eTrade

            oPlayer1TP.TradeID = Me.ObjectID
            oPlayer1TP.SaveObject()
            oPlayer2TP.TradeID = Me.ObjectID
            oPlayer2TP.SaveObject()

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type Trade. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub HandleSubmitTradeMsg(ByRef yData() As Byte, ByVal lPlayerID As Int32)
        'Check our current status...
        'Cannot change the trade object if it is REJECTED or Not in the Proposed or InNegotiation states
        If TradeState = eTradeStateValues.TradeRejected OrElse (TradeState <> eTradeStateValues.Proposal AndAlso (TradeState And eTradeStateValues.InNegotiation) = 0) Then Return
        If TradeState = eTradeStateValues.Proposal Then TradeState = eTradeStateValues.InNegotiation

        'All data is specific for the player passed in by lPlayerID which is the player ID reported by the SOCKET
        Dim oTP As TradePlayer = Nothing
        Dim yPlayer As Byte = 1
        If oPlayer1TP Is Nothing = False AndAlso oPlayer1TP.PlayerID = lPlayerID Then
            oTP = oPlayer1TP
        ElseIf oPlayer2TP Is Nothing = False AndAlso oPlayer2TP.PlayerID = lPlayerID Then
            oTP = oPlayer2TP
            yPlayer = 2
        End If
        If oTP Is Nothing Then Return

        If oPlayer1TP.oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso oPlayer1TP.oPlayer.AccountStatus <> AccountStatusType.eActiveAccount Then Return
        If oPlayer2TP.oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso oPlayer2TP.oPlayer.AccountStatus <> AccountStatusType.eActiveAccount Then Return

        'Ok, the submit trade msg has the following data:
        'TradeGUID (6) - used to find this object so... not needed now
        Dim lPos As Int32 = 8
        'TradeState (1)
        Dim yState As Byte = yData(lPos) : lPos += 1
        'DestID (4)
        'Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'SourceID (4)
        'Dim lSourceID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'TradePostID (4)
        Dim lTPID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'NoteLen (4)
        Dim lNoteLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Notes...
        Dim yNotes(lNoteLen - 1) As Byte
        If lNoteLen <> 0 Then
            Array.Copy(yData, lPos, yNotes, 0, lNoteLen)
        End If
        lPos += lNoteLen

        'ItemCnt (4)
		Dim lItemCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lItemCnt = Math.Max(lItemCnt, 0)
        'Items (ItemCnt * 10)
        Dim lItemIDs(lItemCnt - 1) As Int32
        Dim iTypeIDs(lItemCnt - 1) As Int16
		Dim blQtys(lItemCnt - 1) As Int64
		Dim lExtIDs(lItemCnt - 1) As Int32
        Dim bChanged(lItemCnt - 1) As Boolean

        For X As Int32 = 0 To lItemCnt - 1
            lItemIDs(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iTypeIDs(X) = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			blQtys(X) = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
			lExtIDs(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            bChanged(X) = True
        Next X

        With oTP
            'Copy my notes over
            .Notes = yNotes
            'Set my Dest
            '.DestID = lDestID
            'set my source
            '.SourceID = lSourceID
            .lTradePostID = lTPID
        End With

        'Check if the player is rejecting the trade
        If yState = 255 Then
            'Player is requesting that the trade be rejected
            TradeState = 255
        Else
            'Now, we need to compare everything...
            With oTP
                Dim bChangedAgreement As Boolean = False

                'First, if the ItemUB's don't match
                If lItemIDs.GetUpperBound(0) <> .lItemUB Then
                    bChangedAgreement = True
                Else
                    'Ok, UB's match, but are all the items the same?
                    For X As Int32 = 0 To .lItemUB
                        If .lItemIdx(X) <> -1 Then
                            For Y As Int32 = 0 To lItemCnt - 1
                                If lItemIDs(Y) = .lItemIdx(X) AndAlso iTypeIDs(Y) = .oItems(X).ObjTypeID Then
                                    If .oItems(X).Quantity = blQtys(Y) Then bChanged(Y) = False
                                    Exit For
                                End If
                            Next Y
                        End If
                    Next X

                    For X As Int32 = 0 To lItemCnt - 1
                        If bChanged(X) = True Then
                            bChangedAgreement = True
                            Exit For
                        End If
                    Next X
                End If

                If bChangedAgreement = True Then
                    If (TradeState And eTradeStateValues.Player1Accepted) <> 0 Then TradeState -= eTradeStateValues.Player1Accepted
                    If (TradeState And eTradeStateValues.Player2Accepted) <> 0 Then TradeState -= eTradeStateValues.Player2Accepted

                    .lItemUB = -1
                    For X As Int32 = 0 To lItemIDs.GetUpperBound(0)
						.AddItemQuantity(lItemIDs(X), iTypeIDs(X), blQtys(X), lExtIDs(X), -1)
                    Next X
                End If
            End With

            'Ok, now, if the player is saying that they will immediately accept the new terms
            If yPlayer = 1 Then
                If (yState And eTradeStateValues.Player1Accepted) <> 0 Then
                    TradeState = TradeState Or eTradeStateValues.Player1Accepted
                End If
            Else
                If (yState And eTradeStateValues.Player2Accepted) <> 0 Then
                    TradeState = TradeState Or eTradeStateValues.Player2Accepted
                End If
            End If

            'Now, at this point, we should be able to make an intelligent decision
            If (TradeState And eTradeStateValues.Player1Accepted) <> 0 AndAlso (TradeState And eTradeStateValues.Player2Accepted) <> 0 Then
                'We are ready to execute this trade...
                If (TradeState And eTradeStateValues.InNegotiation) <> 0 Then TradeState -= eTradeStateValues.InNegotiation
                TradeState = TradeState Or eTradeStateValues.InVerification

                If BeginTradeExecution() = True Then
                    If (TradeState And eTradeStateValues.InVerification) <> 0 Then TradeState -= eTradeStateValues.InVerification

                    'Create a Queue Item for 1/2 of the time
                    If Me.CyclesRemaining > Me.TradeCycles \ 2 Then
						AddToQueue(glCurrentCycle + (Me.CyclesRemaining - (Me.TradeCycles \ 2)), QueueItemType.eTradeEventHalfTime, Me.ObjectID, Me.ObjTypeID, -1, -1, 0, 0, 0, 0)
                    End If
                    'Create a Queue Item for the full time
					AddToQueue(glCurrentCycle + Me.CyclesRemaining, QueueItemType.eTradeEventFinal, Me.ObjectID, Me.ObjTypeID, -1, -1, 0, 0, 0, 0)

                    'Set our status to In Progress
                    TradeState = TradeState Or eTradeStateValues.InProgress
                Else
                    TradeState = eTradeStateValues.InNegotiation
                End If
            End If
        End If

        'And now update both players
		Dim yMsg() As Byte = goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand)
		oPlayer1TP.oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewTrades)
		oPlayer2TP.oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewTrades)
	End Sub

	Private Function BeginTradeExecution() As Boolean
		Me.FailureReason = eFailureReason.NoFailureReason

        Me.FailureReason = oPlayer1TP.VerifyDetails(True, oPlayer2TP.TradePost)
		If Me.FailureReason <> eFailureReason.NoFailureReason Then Return False
		Me.FailureReason = oPlayer2TP.VerifyDetails(False, oPlayer1TP.TradePost)
		If Me.FailureReason <> eFailureReason.NoFailureReason Then Return False

		'If all is well, pull the items from the two sources... make their Parent this trade object
		If oPlayer1TP.ExecuteTrade(Me) = False Then
			Me.FailureReason = eFailureReason.Player1MissingItems
			oPlayer1TP.RefundItems()
			Return False
		End If
		If oPlayer2TP.ExecuteTrade(Me) = False Then
			Me.FailureReason = eFailureReason.Player2MissingItems
			oPlayer2TP.RefundItems()
			oPlayer1TP.RefundItems()
			Return False
		End If

		'Calculate the distances from P1.Source to P2.Dest and P2.Source to P1.Dest
		'Based on the distances, set TradeCycles and CyclesRemaining
        Dim lP1toP2 As Int32 = GalacticTradeSystem.GetTradeRouteDeliveryTime(oPlayer1TP.TradePost, oPlayer2TP.TradePost, True)
        Dim lP2toP1 As Int32 = GalacticTradeSystem.GetTradeRouteDeliveryTime(oPlayer2TP.TradePost, oPlayer2TP.TradePost, True)

		Dim lCycles As Int32 = Math.Max(lP1toP2, lP2toP1) * 30
		Me.TradeCycles = lCycles
        Me.TradeStarted = GetDateAsNumber(Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime) '  glCurrentCycle
        Me.CyclesRemaining = lCycles

        'Tax Rate is total of credits transferred * .3 or 10m whichever is higher
        Dim blTotalCredits As Int64 = Math.Max(oPlayer1TP.GetTotalCreditsTransferred(), oPlayer2TP.GetTotalCreditsTransferred())
        blTotalCredits \= 10
        blTotalCredits *= 3
        If blTotalCredits < 10000000 Then blTotalCredits = 10000000
        oPlayer1TP.oPlayer.blCredits -= blTotalCredits
        oPlayer2TP.oPlayer.blCredits -= blTotalCredits

		Return True
	End Function

	Public Sub HandleHalfTime()
		'TODO: Here is where any espionage/sabotage events would occur
	End Sub

	Public Sub FinalizeTrade()
		'Easy...
		oPlayer1TP.FinalizeTrade(oPlayer2TP)
		oPlayer2TP.FinalizeTrade(oPlayer1TP)

		TradeState = eTradeStateValues.TradeCompleted Or eTradeStateValues.Player1Accepted Or eTradeStateValues.Player2Accepted
		DataChanged()
		Dim yMsg() As Byte = goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand)
		oPlayer1TP.oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewTrades)
		oPlayer2TP.oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewTrades)

		CreateTradeHistoryObjects()
	End Sub

    Private Sub CreateTradeHistoryObjects()
        Dim oNewP1 As New TradeHistory()
        Dim oNewP2 As New TradeHistory()

        Dim lTime As Int32 = CInt(Math.Ceiling(TradeCycles / 30.0F))

        'Ok, first, Player1
        With oNewP1
            .lOtherPlayerID = oPlayer2TP.PlayerID
            .lPlayerID = oPlayer1TP.PlayerID
            .lTransactionDate = GetDateAsNumber(Now)
            .yTradeEventType = TradeHistory.TradeEventType.eDirectTrade Or TradeHistory.TradeEventType.eBuyer
            .yTradeResult = TradeHistory.TradeResult.eCompletedSuccessfully
            .lDeliveryTime = lTime
            For X As Int32 = 0 To oPlayer1TP.lItemUB
                .AddTradeItem(StringToBytes(GetTradedItemName(oPlayer1TP.oItems(X))), oPlayer1TP.oItems(X).Quantity, TradeHistory.TradeHistoryItemType.eBuyerSold)
            Next X
            For X As Int32 = 0 To oPlayer2TP.lItemUB
                .AddTradeItem(StringToBytes(GetTradedItemName(oPlayer2TP.oItems(X))), oPlayer2TP.oItems(X).Quantity, TradeHistory.TradeHistoryItemType.eBuyerPurchased)
            Next X
        End With
        'Ok, now player2
        With oNewP2
            .lOtherPlayerID = oPlayer1TP.PlayerID
            .lPlayerID = oPlayer2TP.PlayerID
            .lTransactionDate = oNewP1.lTransactionDate
            .yTradeEventType = TradeHistory.TradeEventType.eDirectTrade Or TradeHistory.TradeEventType.eSeller
            .yTradeResult = TradeHistory.TradeResult.eCompletedSuccessfully
            .lDeliveryTime = lTime
            For X As Int32 = 0 To oPlayer1TP.lItemUB
                .AddTradeItem(StringToBytes(GetTradedItemName(oPlayer1TP.oItems(X))), oPlayer1TP.oItems(X).Quantity, TradeHistory.TradeHistoryItemType.eSellerPurchased)
            Next X
            For X As Int32 = 0 To oPlayer2TP.lItemUB
                .AddTradeItem(StringToBytes(GetTradedItemName(oPlayer2TP.oItems(X))), oPlayer2TP.oItems(X).Quantity, TradeHistory.TradeHistoryItemType.eSellerSold)
            Next X
        End With

        oNewP1.SaveObject()
        oNewP2.SaveObject()

        oPlayer1TP.oPlayer.AddTradeHistory(oNewP1)
        oPlayer2TP.oPlayer.AddTradeHistory(oNewP2)
    End Sub

    Private Function GetTradedItemName(ByRef oItem As TradePlayerItem) As String
        Dim lObjID As Int32 = oItem.ObjectID
        Dim iTypeID As Int16 = oItem.ObjTypeID
        Dim lExt1 As Int32 = oItem.lExtendedID
        Dim lExt2 As Int32 = oItem.lExtended2ID

        Select Case Math.Abs(iTypeID)
            '================== THESE ITEMS REQUIRE THE PARENT COLONY ===============
            Case ObjectType.eColonists
                Return "Colonists"
            Case ObjectType.eEnlisted
                Return "Enlisted"
            Case ObjectType.eOfficers
                Return "Officers"
            Case ObjectType.eCredits
                Return "Credits"
            Case ObjectType.eAgent
                Dim oAgent As Agent = GetEpicaAgent(lObjID)
                If oAgent Is Nothing = False Then Return BytesToString(oAgent.AgentName) Else Return "Unknown Agent"
            Case ObjectType.eColony
                Dim oColony As Colony = GetEpicaColony(lObjID)
                If oColony Is Nothing = False Then Return BytesToString(oColony.ColonyName) Else Return "Unknown Colony"
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
                If iTypeID < 0 Then
                    'oOtherPlayer.TradePost.AddComponentCacheToCargo(oItems(lIdx).lExtended2ID, Math.Abs(iTypeID), CInt(blQty), oItems(lIdx).lExtendedID) Is Nothing Then
                    Dim sResult As String = "Unknown Components"
                    Dim oOwner As Player = GetEpicaPlayer(lExt1)
                    If oOwner Is Nothing = False Then
                        Dim oComponent As Epica_Tech = oOwner.GetTech(lExt2, Math.Abs(iTypeID))
                        If oComponent Is Nothing = False Then sResult = BytesToString(oComponent.GetTechName())
                    End If
                    Return sResult
                End If
            Case ObjectType.eUnit
                'Units are delivered to the target colony's Parent Environment
                Dim sResult As String = "Unknown Unit"
                Dim oUnit As Unit = GetEpicaUnit(lObjID)
                If oUnit Is Nothing = False Then
                    sResult = BytesToString(oUnit.EntityName)
                End If
                Return sResult
            Case ObjectType.eFacility
                Dim sResult As String = "Unknown Facility"
                Dim oFac As Facility = GetEpicaFacility(lObjID)
                If oFac Is Nothing = False Then
                    sResult = BytesToString(oFac.EntityName)
                End If
                Return sResult
            Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
                Return "Intelligence"
            Case Else
                'Either mineral cache or ammunition
                If iTypeID = ObjectType.eAmmunition Then
                    'TODO: Uh...
                Else
                    Dim oMineral As Mineral = GetEpicaMineral(lExt1)
                    If oMineral Is Nothing = False Then
                        Return BytesToString(oMineral.MineralName)
                    Else : Return "Unknown Mineral"
                    End If
                End If
        End Select
        Return "Unknown"
    End Function
End Class