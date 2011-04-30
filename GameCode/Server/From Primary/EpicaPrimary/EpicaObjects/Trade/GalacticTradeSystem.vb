Option Strict On

'Manages all trades and such on the galactic scale (the market)
Public Class GalacticTradeSystem
    Public Enum MarketListType As Byte
        eNotUsed = 0
        eMinerals = 1
        eAlloys = 2
        eArmorComponent = 3
        eBeamPulseComponent = 4
        eBeamSolidComponent = 5
        eEngineComponent = 6
        eMissileComponent = 7
        eProjectileComponent = 8
        eRadarComponent = 9
        eShieldComponent = 10
        eSpecialComponent = 11
        eArmorDesign = 12
        eBeamPulseDesign = 13
        eBeamSolidDesign = 14
        eEngineDesign = 15
        eMissileDesign = 16
        eProjectileDesign = 17
        eRadarDesign = 18
        eShieldDesign = 19
        eLandBasedUnit = 20
        eAerialUnit = 21
        eSpaceBasedUnit = 22
        eNavalUnit = 23
        eSpaceStation = 24
        eAgent = 25
        eColonists = 26
        eEnlisted = 27
        eOfficers = 28
        ePlayerStats = 29 '- scores as of a certain time
        ePlayerComponentDesigns = 30 '- a design schematic
        eColonyLocation = 31 '- bookmarks of colony locations
        eForeignRelationsList = 32 '- the entire diplomacy screen for a player
        ePlayerSenateVoteHistory = 33 '- list of vote history the player took in senate votes of past?
        eBudgetDetails = 34 '- as of a certain time

        eBuyOrderSpecificMineral = 35       'specific mineral

        eBombComponent = 36
        eMineWeaponComponent = 37

        eSellOrderShift = 128           'bitwise item
    End Enum

    Private moSellOrders() As SellOrder
    Private mlSellOrderPlanetID() As Int32
    Private mlSellOrderSystemID() As Int32
	Private mySellOrderType() As MarketListType
	Private mlSellOrderUB As Int32 = -1

    Private moBuyOrders() As BuyOrder
    Private mlBuyOrderPlanetID() As Int32
    Private mlBuyOrderSystemID() As Int32
    Private myBuyOrderType() As MarketListType
    Private mlBuyOrderUB As Int32 = -1

    Private moTradeDelivery() As TradeDelivery
    Private myTradeDeliveryUsed() As Byte
    Private mlTradeDeliveryUB As Int32 = -1

    Private moDirectTrade() As DirectTrade
    Private mlDirectTradeIdx() As Int32
    Private mlDirectTradeUB As Int32 = -1

    Public Sub CheckBuyOrderDeadlines()
        Dim lNow As Int32 = GetDateAsNumber(Now)
        For X As Int32 = 0 To mlBuyOrderUB
            If myBuyOrderType(X) <> MarketListType.eNotUsed AndAlso moBuyOrders(X).lDeadline < lNow Then
                'Ok, got a buy order that has reached its deadline... 
                With moBuyOrders(X)
                    'Ok, first, is there an accepted player?
                    If .lAcceptedByID <> -1 Then
                        'Yes, it was accepted... has it been fulfilled?
                        If .yBuyOrderState = 255 Then
                            'Yes, nothing to do except remove it... the TradePostBuySlotsUed and OtherJobs should be updated already
                        Else
                            'No, so, buy order placer is rewarded the escrow and the Payment Amt
                            .TradePost.Owner.blCredits += (.blPaymentAmt + .blEscrow)
                            .TradePost.lTradePostBuySlotsUsed -= 1
                            If .TradePost.ParentColony Is Nothing = False Then .TradePost.ParentColony.UpdateAllValues(-1)
                            
                            Try
                                Dim oTH As New TradeHistory()
                                oTH.lPlayerID = .TradePost.Owner.ObjectID
                                oTH.lOtherPlayerID = oTH.lPlayerID
                                oTH.lTransactionDate = GetDateAsNumber(Now)
                                oTH.blTradeAmt = (.blPaymentAmt + .blEscrow)
                                oTH.yTradeEventType = TradeHistory.TradeEventType.eBuyOrder Or TradeHistory.TradeEventType.eBuyer
                                oTH.yTradeResult = TradeHistory.TradeResult.eBuyOrderAcceptorFailed
                                oTH.lDeliveryTime = 0
                                oTH.SaveObject()
                                .TradePost.Owner.AddTradeHistory(oTH)
                                oTH = Nothing
                            Catch
                            End Try
                        End If
                    Else
                        'No, so buy order placer is rewarded the payment amt
                        .TradePost.Owner.blCredits += .blPaymentAmt
                        .TradePost.lTradePostBuySlotsUsed -= 1
                        If .TradePost.ParentColony Is Nothing = False Then .TradePost.ParentColony.UpdateAllValues(-1)
                        Try
                            Dim oTH As New TradeHistory()
                            oTH.lPlayerID = .TradePost.Owner.ObjectID
                            oTH.lOtherPlayerID = oTH.lPlayerID
                            oTH.lTransactionDate = GetDateAsNumber(Now)
                            oTH.blTradeAmt = .blPaymentAmt
                            oTH.yTradeEventType = TradeHistory.TradeEventType.eBuyOrder Or TradeHistory.TradeEventType.eBuyer
                            oTH.yTradeResult = TradeHistory.TradeResult.eNoBuyOrderAcceptor
                            oTH.lDeliveryTime = 0
                            oTH.SaveObject()
                            .TradePost.Owner.AddTradeHistory(oTH)
                            oTH = Nothing
                        Catch
                        End Try
                    End If

                    .yBuyOrderState = 255

                    'Do a final save
                    .SaveObject()

                    'finally, the buy order has reached its deadline, so remove it
                    If .TradePost Is Nothing = False Then .TradePost.lTradePostBuySlotsUsed -= 1
                    myBuyOrderType(X) = MarketListType.eNotUsed
                End With
                
            End If
        Next X
    End Sub

    Public Sub LoadSellOrder(ByRef oOrder As SellOrder)
        'Ok, if we are here, place oNew within the oSellOrder list
        Dim lPlanetID As Int32 = -1
        Dim lSystemID As Int32 = -1
        With CType(oOrder.TradePost.ParentObject, Epica_GUID)
            If .ObjTypeID = ObjectType.ePlanet Then
                lPlanetID = .ObjectID
                lSystemID = CType(oOrder.TradePost.ParentObject, Planet).ParentSystem.ObjectID
            ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                lPlanetID = -1
                lSystemID = .ObjectID
            ElseIf .ObjTypeID = ObjectType.eFacility Then
                With CType(CType(oOrder.TradePost.ParentObject, Facility).ParentObject, Epica_GUID)
                    lSystemID = .ObjectID
                End With
            End If
        End With

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlSellOrderUB
            If mySellOrderType(X) = MarketListType.eNotUsed Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            Dim lTmpUB As Int32 = mlSellOrderUB + 1
            ReDim Preserve mySellOrderType(lTmpUB)
            ReDim Preserve moSellOrders(lTmpUB)
            ReDim Preserve mlSellOrderPlanetID(lTmpUB)
			ReDim Preserve mlSellOrderSystemID(lTmpUB)
			mlSellOrderUB += 1
            lIdx = lTmpUB
        End If

        moSellOrders(lIdx) = oOrder
        mlSellOrderPlanetID(lIdx) = lPlanetID
        mlSellOrderSystemID(lIdx) = lSystemID
		Dim yType As Byte = GetMarketType(oOrder.lTradeID, oOrder.iTradeTypeID, oOrder.TradePost, oOrder.iExtTypeID)
        If yType = 0 Then Return
		mySellOrderType(lIdx) = CType(yType Or MarketListType.eSellOrderShift, MarketListType)

        'Increment our sell slots used
		oOrder.TradePost.lTradePostSellSlotsUsed += 1
		If oOrder.TradePost.ParentColony Is Nothing = False Then
			oOrder.TradePost.ParentColony.OtherJobs += oOrder.TradePost.EntityDef.WorkerFactor
			oOrder.TradePost.ParentColony.NumberOfJobs += oOrder.TradePost.EntityDef.WorkerFactor
		End If
	End Sub

    Public Sub LoadTradeDelivery(ByRef oNew As TradeDelivery)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlTradeDeliveryUB
            If myTradeDeliveryUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            ReDim Preserve moTradeDelivery(mlTradeDeliveryUB + 1)
            ReDim Preserve myTradeDeliveryUsed(mlTradeDeliveryUB + 1)
            mlTradeDeliveryUB += 1
            lIdx = mlTradeDeliveryUB
        End If
        moTradeDelivery(lIdx) = oNew
        myTradeDeliveryUsed(lIdx) = 255

        'ok, get the seconds diff
        Dim lSecondsDiff As Int32 = CInt(oNew.dtDelivery.Subtract(oNew.dtStartedOn).TotalSeconds)
        'Check our half time
        Dim dtTemp As Date = oNew.dtStartedOn.AddSeconds(lSecondsDiff \ 2)
        If dtTemp > Now Then
            Dim lDiff As Int32 = 0
            Try
                lDiff = CInt(dtTemp.Subtract(Now).TotalSeconds * 30.0F)
                AddToQueue(glCurrentCycle + lDiff, QueueItemType.eTradeEventHalfTime, lIdx, ObjectType.eTrade, 1, -1, 0, 0, 0, 0)
            Catch
            End Try
        End If
        'check our final time
        If oNew.dtDelivery > Now Then
            Try
                Dim lDiff As Int32 = CInt(oNew.dtDelivery.Subtract(Now).TotalSeconds * 30.0F)
                AddToQueue(glCurrentCycle + lDiff, QueueItemType.eTradeEventFinal, lIdx, ObjectType.eTrade, 1, -1, 0, 0, 0, 0)
            Catch
            End Try
        Else : AddToQueue(glCurrentCycle, QueueItemType.eTradeEventFinal, lIdx, ObjectType.eTrade, 1, -1, 0, 0, 0, 0)
        End If
    End Sub

    Public Sub LoadBuyOrder(ByRef oOrder As BuyOrder)
        'Ok, if we are here, place oNew within the oBuyOrder list
        Dim lPlanetID As Int32 = -1
        Dim lSystemID As Int32 = -1
        With CType(oOrder.TradePost.ParentObject, Epica_GUID)
            If .ObjTypeID = ObjectType.ePlanet Then
                lPlanetID = .ObjectID
                lSystemID = CType(oOrder.TradePost.ParentObject, Planet).ParentSystem.ObjectID
            ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                lPlanetID = -1
                lSystemID = .ObjectID
            ElseIf .ObjTypeID = ObjectType.eFacility Then
                With CType(CType(oOrder.TradePost.ParentObject, Facility).ParentObject, Epica_GUID)
                    lSystemID = .ObjectID
                End With
            End If
        End With

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlBuyOrderUB
            If myBuyOrderType(X) = MarketListType.eNotUsed Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            Dim lTmpUB As Int32 = mlBuyOrderUB + 1
            ReDim Preserve myBuyOrderType(lTmpUB)
            ReDim Preserve moBuyOrders(lTmpUB)
            ReDim Preserve mlBuyOrderPlanetID(lTmpUB)
            ReDim Preserve mlBuyOrderSystemID(lTmpUB)
            mlBuyOrderUB += 1
            lIdx = lTmpUB
        End If

        moBuyOrders(lIdx) = oOrder
        mlBuyOrderPlanetID(lIdx) = lPlanetID
        mlBuyOrderSystemID(lIdx) = lSystemID
        Dim yType As Byte = oOrder.yBuyOrderType
        If yType = 0 Then Return
        myBuyOrderType(lIdx) = CType(yType, MarketListType)

        'Increment our buy slots used
        oOrder.TradePost.lTradePostBuySlotsUsed += 1
        oOrder.TradePost.ParentColony.OtherJobs += oOrder.TradePost.EntityDef.WorkerFactor
        oOrder.TradePost.ParentColony.NumberOfJobs += oOrder.TradePost.EntityDef.WorkerFactor
    End Sub

    Public Function HandleSubmitSellOrder(ByRef yData() As Byte) As Boolean
        Dim bResult As Boolean = False
        Try
            Dim oNew As New SellOrder()
            Dim lPos As Int32 = 2           '2 for msgcode
            With oNew
                .TradePostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lTradeID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iTradeTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .blQuantity = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
				.blPrice = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
				.iExtTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                If .TradePost Is Nothing Then Return False
                If .TradePost.yProductionType <> ProductionType.eTradePost Then Return False
            End With

            'Ok, attempt to find our trade object already in the array
            Dim bPreExist As Boolean = False
            Dim lTempPreExistID As Int32 = oNew.lTradeID 
            If oNew.iTradeTypeID < 0 Then
                Dim bFound As Boolean = False
                Dim oCache As ComponentCache = GetEpicaComponentCache(oNew.lTradeID)
                If oCache Is Nothing = False Then
                    If oCache.ComponentTypeID = Math.Abs(oNew.iTradeTypeID) Then
                        lTempPreExistID = oCache.ComponentID
                        bFound = True
                    End If
                End If 
            End If
            For X As Int32 = 0 To mlSellOrderUB
                If mySellOrderType(X) <> MarketListType.eNotUsed AndAlso moSellOrders(X).lTradeID = lTempPreExistID AndAlso moSellOrders(X).TradePostID = oNew.TradePostID AndAlso moSellOrders(X).iTradeTypeID = oNew.iTradeTypeID Then
                    'Ok, already found...
                    bPreExist = True

                    If oNew.blQuantity = -1 OrElse oNew.blPrice = -1 Then
                        If moSellOrders(X).iTradeTypeID = ObjectType.eUnit Then
                            Dim oUnit As Unit = GetEpicaUnit(moSellOrders(X).lTradeID)
                            If oUnit Is Nothing = False AndAlso oUnit.Owner.ObjectID = moSellOrders(X).TradePost.ObjectID Then
                                'Ok, do a dismantle of the unit
                                moSellOrders(X).TradePost.DismantleChildUnit(oUnit)
                                DestroyEntity(CType(oUnit, Epica_Entity), False, -1, False, "CancelSellOrder")
                            End If
                        End If
                        moSellOrders(X).DeleteMe()

                        mySellOrderType(X) = MarketListType.eNotUsed
                        oNew.TradePost.lTradePostSellSlotsUsed -= 1
                        Return True
                    ElseIf moSellOrders(X).iTradeTypeID = ObjectType.eUnit Then
                        If moSellOrders(X).TradePost.Owner.blCredits < moSellOrders(X).blPrice \ 4 Then
                            Return False
                        End If
                    End If

                    moSellOrders(X).blQuantity = oNew.blQuantity
                    moSellOrders(X).blPrice = oNew.blPrice

                    Exit For
                End If
            Next X

            If bPreExist = False AndAlso oNew.TradePost.Owner Is Nothing = False AndAlso oNew.TradePost.Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID AndAlso oNew.TradePost.lTradePostSellSlotsUsed >= oNew.TradePost.Owner.oSpecials.ySellOrderSlots Then
                'TODO: Indicate maximum sell orders is the reason that the sell order did not pass
                Return False
            End If

			If bPreExist = False AndAlso (oNew.blQuantity < 0 OrElse oNew.blPrice < 0) Then Return False

            'Now, we need to confirm that the object(s) exist in the tradepost's cargo/hangar
            Select Case oNew.iTradeTypeID
                Case ObjectType.eMineralCache
                    If oNew.GetQuantityInStock() < oNew.blQuantity Then Return False
                    Dim oCache As MineralCache = GetEpicaMineralCache(oNew.lTradeID)
                    If oCache Is Nothing Then Return False
                    oCache.oMineral.MineralName.CopyTo(oNew.yItemName, 0)
                Case ObjectType.eUnit
                    If oNew.GetQuantityInStock() < oNew.blQuantity Then Return False
                    Dim oUnit As Unit = GetEpicaUnit(oNew.lTradeID)
                    If oUnit Is Nothing Then Return False
                    oUnit.EntityName.CopyTo(oNew.yItemName, 0)
                    'remove it from any unitgroup it may be
                    If oUnit.lFleetID > -1 Then
                        Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(oUnit.lFleetID)
                        If oUnitGroup Is Nothing = False Then oUnitGroup.RemoveUnit(oUnit.ObjectID, True, False)
                        oUnit.lFleetID = -1
                    End If
                    oUnit.bUnitInSellOrder = True
                Case ObjectType.ePlayerIntel
                    'ok, verify the player has this intel
                    If oNew.TradePost Is Nothing = False AndAlso oNew.TradePost.Owner Is Nothing = False Then
                        Dim oPlayer As Player = oNew.TradePost.Owner
                        For X As Int32 = 0 To oPlayer.mlPlayerIntelUB
                            If oPlayer.mlPlayerIntelIdx(X) = oNew.lTradeID Then
                                Dim oTmpPlayer As Player = GetEpicaPlayer(oNew.lTradeID)
                                If oTmpPlayer Is Nothing = False Then
                                    oNew.yItemName = oTmpPlayer.PlayerName
                                End If
                                Exit For
                            End If
                        Next X
                    Else : Return False
                    End If
                Case ObjectType.ePlayerItemIntel
                    If oNew.TradePost Is Nothing = False AndAlso oNew.TradePost.Owner Is Nothing = False Then
                        Dim oPlayer As Player = oNew.TradePost.Owner
                        Dim bFound As Boolean = False
                        For X As Int32 = 0 To oPlayer.mlItemIntelUB
                            If oPlayer.moItemIntel(X) Is Nothing = False Then
                                If oPlayer.moItemIntel(X).lItemID = oNew.lTradeID AndAlso oPlayer.moItemIntel(X).iItemTypeID = oNew.iExtTypeID Then

                                    Dim oItem As PlayerItemIntel = oPlayer.moItemIntel(X)
                                    If oItem Is Nothing = False Then
                                        Dim oObj As Object = GetEpicaObject(oItem.lItemID, oItem.iItemTypeID)
                                        If oObj Is Nothing = False Then
                                            If oItem.iItemTypeID = ObjectType.eColony Then
                                                oNew.yItemName = StringToBytes(CType(oObj, Colony).Owner.sPlayerNameProper)
                                            Else
                                                oNew.yItemName = GetEpicaObjectName(oItem.iItemTypeID, oObj)
                                            End If
                                        End If
                                    End If

                                    bFound = True
                                    Exit For
                                End If
                            End If
                        Next X
                        If bFound = False Then Return False
                    End If
                Case ObjectType.ePlayerTechKnowledge
                    If oNew.TradePost Is Nothing = False AndAlso oNew.TradePost.Owner Is Nothing = False Then
                        Dim oPlayer As Player = oNew.TradePost.Owner
                        If oPlayer.HasTechKnowledge(oNew.lTradeID, oNew.iExtTypeID, PlayerTechKnowledge.KnowledgeType.eSettingsLevel1) = False Then Return False
                        oNew.yItemName = oPlayer.GetPlayerTechKnowledgeTech(oNew.lTradeID, oNew.iExtTypeID, PlayerTechKnowledge.KnowledgeType.eNameOnly).GetTechName()
                    End If
                Case Is < 0
                    Dim oColony As Colony = oNew.TradePost.ParentColony
                    If oColony Is Nothing = False Then
                        Dim lQty As Int32 = 0
                        Dim bItemSet As Boolean = False
                        'For X As Int32 = 0 To oColony.ChildrenUB
                        '    If oColony.lChildrenIdx(X) > -1 Then
                        '        Dim oFac As Facility = oColony.oChildren(X)
                        '        If oFac Is Nothing = False Then
                        '            If (oFac.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then

                        '                For Y As Int32 = 0 To oFac.lCargoUB
                        '                    If oFac.lCargoIdx(Y) <> -1 AndAlso oFac.oCargoContents(Y).ObjTypeID = ObjectType.eComponentCache Then
                        '                        Dim oCache As ComponentCache = CType(oFac.oCargoContents(Y), ComponentCache)
                        '                        If oCache.ObjectID = oNew.lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(oNew.iTradeTypeID) Then
                        '                            lQty += oCache.Quantity
                        '                            If bItemSet = False Then
                        '                                oCache.GetComponent.GetTechName().CopyTo(oNew.yItemName, 0)
                        '                                oNew.lTradeID = oCache.ComponentID
                        '                                bItemSet = True
                        '                            End If
                        '                            Exit For
                        '                        End If
                        '                    End If
                        '                Next Y
                        '            End If
                        '        End If
                        '    End If
                        'Next X

                        For X As Int32 = 0 To oColony.mlComponentCacheUB
                            If oColony.mlComponentCacheIdx(X) > -1 Then
                                If oColony.mlComponentCacheID(X) = glComponentCacheIdx(oColony.mlComponentCacheIdx(X)) Then
                                    Dim oCache As ComponentCache = goComponentCache(oColony.mlComponentCacheIdx(X))
                                    If oCache Is Nothing = False Then
                                        If oCache.ObjectID = oNew.lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(oNew.iTradeTypeID) Then
                                            lQty += oCache.Quantity
                                            If bItemSet = False Then
                                                oCache.GetComponent.GetTechName().CopyTo(oNew.yItemName, 0)
                                                oNew.lTradeID = oCache.ComponentID
                                                bItemSet = True
                                            End If
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next X

                        If lQty < oNew.blQuantity Then Return False
                        If bItemSet = False Then Return False
                    End If
                    
                Case Else
                    'TODO: What else?
            End Select

            'Ok, if we are here, place oNew within the oSellOrder list
            Dim lPlanetID As Int32 = -1
            Dim lSystemID As Int32 = -1
            With CType(oNew.TradePost.ParentObject, Epica_GUID)
                If .ObjTypeID = ObjectType.ePlanet Then
                    lPlanetID = .ObjectID
                    lSystemID = CType(oNew.TradePost.ParentObject, Planet).ParentSystem.ObjectID
                ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                    lPlanetID = -1
                    lSystemID = .ObjectID
                ElseIf .ObjTypeID = ObjectType.eFacility Then
                    With CType(CType(oNew.TradePost.ParentObject, Facility).ParentObject, Epica_GUID)
                        lSystemID = .ObjectID
                    End With
                End If
            End With

            If bPreExist = False Then
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To mlSellOrderUB
                    If mySellOrderType(X) = MarketListType.eNotUsed Then
                        lIdx = X
                        Exit For
                    End If
                Next X
                If lIdx = -1 Then
                    Dim lTmpUB As Int32 = mlSellOrderUB + 1
                    ReDim Preserve mySellOrderType(lTmpUB)
                    ReDim Preserve moSellOrders(lTmpUB)
                    ReDim Preserve mlSellOrderPlanetID(lTmpUB)
                    ReDim Preserve mlSellOrderSystemID(lTmpUB)
                    mlSellOrderUB += 1
                    lIdx = lTmpUB
                End If

                moSellOrders(lIdx) = oNew
                mlSellOrderPlanetID(lIdx) = lPlanetID
                mlSellOrderSystemID(lIdx) = lSystemID
				Dim yType As Byte = GetMarketType(oNew.lTradeID, oNew.iTradeTypeID, oNew.TradePost, oNew.iExtTypeID)
                If yType = 0 Then Return False
                mySellOrderType(lIdx) = CType(yType Or MarketListType.eSellOrderShift, MarketListType)

                'Increment our sell slots used
                oNew.TradePost.lTradePostSellSlotsUsed += 1
                If oNew.TradePost.ParentColony Is Nothing = False Then oNew.TradePost.ParentColony.UpdateAllValues(-1)
                oNew.TradePost.Owner.TestCustomTitlePermissions_Trade()
            End If

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "HandleSubmitSellOrder: " & ex.Message)
            bResult = False
        End Try
        Return bResult
    End Function

    Public Sub HandleSubmitBuyOrder(ByRef yData() As Byte, ByVal lPlayerID As Int32)
        Dim oPlayer As Player = Nothing
        Dim lTradePost As Int32 = -1
        Try
            oplayer = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing Then Return

            Dim lPos As Int32 = 2           'for msgcode
            lTradePost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yItemType As Byte = yData(lPos) : lPos += 1
            Dim lDays As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lHours As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim blEscrow As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            Dim blPayment As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            Dim blQty As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

            Dim lCnt As Int32 = 0

            If yItemType = MarketListType.eBuyOrderSpecificMineral Then
                lCnt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Else
                lCnt = yData(lPos) : lPos += 1
            End If

            'Ok, create our buy order
            Dim oNew As New BuyOrder()
            With oNew
                .blEscrow = blEscrow
                .blPaymentAmt = blPayment
                .blQuantity = blQty
                .BuyOrderID = -1
                .iTradeTypeID = -1      'not sure what this is for
                .lAcceptedByID = -1
                .lAcceptedOn = -1
                Dim dtDeadline As Date = Now.AddDays(lDays + (lHours / 24.0F))
                .lDeadline = GetDateAsNumber(dtDeadline)
                .lSpecificID = -1
                .TradePostID = lTradePost
                .yBuyOrderState = 0
                .yBuyOrderType = yItemType

                If yItemType = MarketListType.eBuyOrderSpecificMineral Then
                    .AddBOProp(BuyOrder.elBuyOrderPropID.eSpecificMineralID, lCnt, BuyOrder.BuyOrderCompareTypes.eEqualTo)
                Else
                    For X As Int32 = 0 To lCnt - 1
                        Dim lProp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Dim lVal As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Dim yCompare As Byte = yData(lPos) : lPos += 1

                        .AddBOProp(lProp, lVal, yCompare)
                    Next X
                End If

                If .TradePost Is Nothing OrElse .TradePost.Owner.ObjectID <> oPlayer.ObjectID OrElse .TradePost.yProductionType <> ProductionType.eTradePost Then
                    Return
                End If

                If .TradePost.lTradePostBuySlotsUsed + 1 > .TradePost.Owner.oSpecials.yBuyOrderSlots Then
                    Dim yResp(6) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eSubmitBuyOrder).CopyTo(yResp, 0)
					System.BitConverter.GetBytes(lTradePost).CopyTo(yResp, 2)
					yResp(6) = 1		'1 indicates max slots exceeded
					oPlayer.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
					Return
				End If

				If .ValidateData() = False Then
					Dim yResp(6) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eSubmitBuyOrder).CopyTo(yResp, 0)
					System.BitConverter.GetBytes(lTradePost).CopyTo(yResp, 2)
					yResp(6) = 3		'3 indicates invlaid properties
					oPlayer.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
					Return
				End If

				If oPlayer.blCredits < .blPaymentAmt Then
					Dim yResp(6) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eSubmitBuyOrder).CopyTo(yResp, 0)
					System.BitConverter.GetBytes(lTradePost).CopyTo(yResp, 2)
					yResp(6) = 4		'4 indicates not enough credits
					oPlayer.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
					Return
				End If
			End With

			'Ok, if we are here, place oNew within the oBuyOrder list
			Dim lPlanetID As Int32 = -1
			Dim lSystemID As Int32 = -1
			With CType(oNew.TradePost.ParentObject, Epica_GUID)
				If .ObjTypeID = ObjectType.ePlanet Then
					lPlanetID = .ObjectID
					lSystemID = CType(oNew.TradePost.ParentObject, Planet).ParentSystem.ObjectID
				ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
					lPlanetID = -1
					lSystemID = .ObjectID
				ElseIf .ObjTypeID = ObjectType.eFacility Then
					With CType(CType(oNew.TradePost.ParentObject, Facility).ParentObject, Epica_GUID)
						lSystemID = .ObjectID
					End With
				End If
			End With

			Dim lIdx As Int32 = -1
			For X As Int32 = 0 To mlBuyOrderUB
				If myBuyOrderType(X) = MarketListType.eNotUsed Then
					lIdx = X
					Exit For
				End If
			Next X
			If lIdx = -1 Then
				Dim lTmpUB As Int32 = mlBuyOrderUB + 1
				ReDim Preserve myBuyOrderType(lTmpUB)
				ReDim Preserve moBuyOrders(lTmpUB)
				ReDim Preserve mlBuyOrderPlanetID(lTmpUB)
				ReDim Preserve mlBuyOrderSystemID(lTmpUB)
				mlBuyOrderUB += 1
				lIdx = lTmpUB
			End If

			moBuyOrders(lIdx) = oNew
			mlBuyOrderPlanetID(lIdx) = lPlanetID
			mlBuyOrderSystemID(lIdx) = lSystemID
			myBuyOrderType(lIdx) = CType(oNew.yBuyOrderType, MarketListType)

			'Increment our Buy slots used
			oNew.TradePost.lTradePostBuySlotsUsed += 1
            If oNew.TradePost.ParentColony Is Nothing = False Then oNew.TradePost.ParentColony.UpdateAllValues(-1)
			'Reduce the credits of the player
			oPlayer.blCredits -= oNew.blPaymentAmt

			Try
				Dim oTH As New TradeHistory()
				With oTH
					.blTradeAmt = -oNew.blPaymentAmt
					.lOtherPlayerID = oPlayer.ObjectID
					.lPlayerID = oPlayer.ObjectID
					.lTransactionDate = GetDateAsNumber(Now)
					.yTradeEventType = TradeHistory.TradeEventType.eBuyOrder Or TradeHistory.TradeEventType.eBuyer
                    .yTradeResult = TradeHistory.TradeResult.eBuyOrderPlaced
                    .lDeliveryTime = 0
					.SaveObject()
				End With
				oPlayer.AddTradeHistory(oTH)
				oTH = Nothing
			Catch ex As Exception
				LogEvent(LogEventType.CriticalError, "Unable to create TradeHistory in HandleSubmitBuyOrder: " & ex.Message)
			End Try

			If True = True Then
				Dim yResp(6) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eSubmitBuyOrder).CopyTo(yResp, 0)
				System.BitConverter.GetBytes(lTradePost).CopyTo(yResp, 2)
				yResp(6) = 0		'0 indicates success
				oPlayer.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
			End If
		Catch ex As Exception
			LogEvent(LogEventType.Informational, "HandleSubmitBuyOrder: " & ex.Message)
			Dim yResp(6) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eSubmitBuyOrder).CopyTo(yResp, 0)
			System.BitConverter.GetBytes(lTradePost).CopyTo(yResp, 2)
			yResp(6) = 2		'2 indicates something bad
			oPlayer.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
		End Try

	End Sub

	Public Sub HandleAcceptBuyOrder(ByRef yData() As Byte, ByVal lPlayerID As Int32)
		Dim lPos As Int32 = 2		'for msgcode
		Dim lTradePostID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4		'tradepost doing the acceptance
		Dim lBO_TradePostID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4		'tradepost with the buyorder
		Dim lBuyOrderID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oTradePost As Facility = GetEpicaFacility(lTradePostID)
		If oTradePost Is Nothing Then Return
		If oTradePost.yProductionType <> ProductionType.eTradePost Then Return
		If oTradePost.Owner.ObjectID <> lPlayerID Then Return

		'Figure out what is in range of this trade post based on the player's range value
		Dim yRng As Byte = oTradePost.Owner.oSpecials.yTradeBoardRange
		If yRng = 0 Then Return '0 indicates no range
		'Ok, now, make it easy for our lookups
		Dim lTradePostPlanetID As Int32 = -1
		Dim lTradePostSystemID As Int32 = -1

		With CType(oTradePost.ParentObject, Epica_GUID)
			If .ObjTypeID = ObjectType.ePlanet Then
				lTradePostPlanetID = .ObjectID
				lTradePostSystemID = CType(oTradePost.ParentObject, Planet).ParentSystem.ObjectID
			ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
				lTradePostPlanetID = -1
				lTradePostSystemID = .ObjectID
			ElseIf .ObjTypeID = ObjectType.eFacility Then
				With CType(CType(oTradePost.ParentObject, Facility).ParentObject, Epica_GUID)
					lTradePostSystemID = .ObjectID
				End With
			End If
		End With

		'find our buy order
		For X As Int32 = 0 To mlBuyOrderUB
			If myBuyOrderType(X) <> MarketListType.eNotUsed AndAlso moBuyOrders(X).BuyOrderID = lBuyOrderID Then
				'First, verify Escrow
				If oTradePost.Owner.blCredits < moBuyOrders(X).blEscrow Then
					'Indicate that the owner did not have enough credits
					Dim yResp(6) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eAcceptBuyOrder).CopyTo(yResp, 0)
					System.BitConverter.GetBytes(lBuyOrderID).CopyTo(yResp, 2)
					yResp(6) = 1			'insufficient credits
					oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
					oTradePost.Owner.SendPlayerMessage(oTradePost.ParentColony.GetLowResourcesMsg(ProductionType.eCredits, 0, 0, -1, -1), False, AliasingRights.eViewTreasury)
					Return
				End If

				'Next, verify range
				If IsBuyOrderInTradePostRange(X, yRng, lTradePostPlanetID, lTradePostSystemID) = False Then
					'Indicate that the buy order is not in range
					Dim yResp(6) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eAcceptBuyOrder).CopyTo(yResp, 0)
					System.BitConverter.GetBytes(lBuyOrderID).CopyTo(yResp, 2)
					yResp(6) = 2			'out of range
					oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
					Return
				End If

				If moBuyOrders(X).AcceptMe(oTradePost.Owner.ObjectID) = True Then
					oTradePost.Owner.blCredits -= moBuyOrders(X).blEscrow

					'Now, make a trade history event for the credits
					Dim oTH As New TradeHistory()
					With oTH
						.lTransactionDate = GetDateAsNumber(Now)
						.blTradeAmt = -moBuyOrders(X).blEscrow
						.lOtherPlayerID = moBuyOrders(X).TradePost.Owner.ObjectID
						.lPlayerID = oTradePost.Owner.ObjectID
						.yTradeEventType = TradeHistory.TradeEventType.eBuyOrder Or TradeHistory.TradeEventType.eSeller
                        .yTradeResult = TradeHistory.TradeResult.eBuyOrderEscrow
                        .lDeliveryTime = 0
						.SaveObject()
					End With
					oTradePost.Owner.AddTradeHistory(oTH)
					oTH = Nothing

                    moBuyOrders(X).TradePost.Owner.CreateAndSendPlayerAlert(PlayerAlertType.eBuyOrderAccepted, moBuyOrders(X).BuyOrderID, -1, moBuyOrders(X).BuyOrderID, -1, oTradePost.Owner.ObjectID, "", Int32.MinValue, Int32.MinValue, "")

					'Tell the acceptor that the buy order was accepted
					Dim yResp(6) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eAcceptBuyOrder).CopyTo(yResp, 0)
					System.BitConverter.GetBytes(lBuyOrderID).CopyTo(yResp, 2)
					yResp(6) = 0			'all good
					oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
				Else
					'indicate that the buy order has already been accepted
					Dim yResp(6) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eAcceptBuyOrder).CopyTo(yResp, 0)
					System.BitConverter.GetBytes(lBuyOrderID).CopyTo(yResp, 2)
					yResp(6) = 3			'no longer available
					oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
				End If

				Exit For
			End If
		Next X

	End Sub

	Public Sub HandlePurchaseSellOrder(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		Dim lTradePostID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4		'tradepost doing the purchase
		Dim lSO_TradePostId As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4		'tradepost of the sell order
		Dim lSO_TradeID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iSO_TradeTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim blQty As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

		Dim oTradePost As Facility = GetEpicaFacility(lTradePostID)
		If oTradePost Is Nothing Then Return
		If oTradePost.yProductionType <> ProductionType.eTradePost Then Return

		'Figure out what is in range of this trade post based on the player's range values
		Dim yRng As Byte = oTradePost.Owner.oSpecials.yTradeBoardRange
        If yRng = 0 Then Return '0 indicates no range
        yRng = 3

		'Ok, now, make it easy for our lookups
		Dim lTradePostPlanetID As Int32 = -1
		Dim lTradePostSystemID As Int32 = -1

		With CType(oTradePost.ParentObject, Epica_GUID)
			If .ObjTypeID = ObjectType.ePlanet Then
				lTradePostPlanetID = .ObjectID
				lTradePostSystemID = CType(oTradePost.ParentObject, Planet).ParentSystem.ObjectID
			ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
				lTradePostPlanetID = -1
				lTradePostSystemID = .ObjectID
			ElseIf .ObjTypeID = ObjectType.eFacility Then
				With CType(CType(oTradePost.ParentObject, Facility).ParentObject, Epica_GUID)
					lTradePostSystemID = .ObjectID
				End With
			End If
		End With

		'Now, find our sell order
		For X As Int32 = 0 To mlSellOrderUB
			If mySellOrderType(X) <> MarketListType.eNotUsed Then
				With moSellOrders(X)
					If .TradePostID = lSO_TradePostId AndAlso .lTradeID = lSO_TradeID AndAlso .iTradeTypeID = iSO_TradeTypeID Then

						'is the quantity still valid?
                        Dim blQtyInStock As Int64 = .GetQuantityInStock()
                        If blQtyInStock < .blQuantity Then .blQuantity = blQtyInStock
                        If blQtyInStock < blQty Then
                            'alert the player that the purchase was unsuccessful
                            Dim yResp(2) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.ePurchaseSellOrder).CopyTo(yResp, 0)
                            yResp(2) = 1            'Quantity unavailable
                            oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
                            Return
                        End If

						'found it... does the tradepost owner have enough credits?
						Dim blBuyCost As Int64 = .blPrice * blQty
                        Dim bFree As Boolean = oTradePost Is Nothing = False AndAlso .TradePost Is Nothing = False AndAlso .TradePost.Owner Is Nothing = False AndAlso oTradePost.Owner.TradesAreFree(.TradePost.Owner.ObjectID)
						If bFree = False Then
							blBuyCost = CLng(blBuyCost * (1.0F + (CSng(oTradePost.Owner.oSpecials.yTradeCosts) / 100.0F)))
						End If

                        If oTradePost.Owner.blCredits >= blBuyCost Then 'AndAlso (iSO_TradeTypeID <> ObjectType.eUnit OrElse oTradePost.Owner.blWarpoints > 1) Then

                            'Ok, verify the range to the sell order
                            If .TradePost Is Nothing = False AndAlso IsSellOrderInTradePostRange(X, yRng, lTradePostPlanetID, lTradePostSystemID) = False Then Return

                            If .PurchaseMe(blQty) = True Then
                                'Alright, purchased, reduce the source's credits
                                oTradePost.Owner.blCredits -= blBuyCost
                                'increase the seller's credits
                                If .TradePost Is Nothing = False AndAlso .TradePost.Owner Is Nothing = False Then .TradePost.Owner.blCredits += (.blPrice * blQty)

                                'Now, create the trade route... let's get the time it will take to deliver it
                                Dim lTotalTime As Int32 = GetTradeRouteDeliveryTime(.TradePost, oTradePost, False)
                                'Create the trade history journal entries
                                TradeHistory.CreateAndSaveSellOrderPurchase(moSellOrders(X), oTradePost.Owner.ObjectID, blQty, blBuyCost, lTotalTime \ 30)

                                'Now, reduce the item from the tradepost, but first, let's get any important details about it
                                Dim lID1 As Int32 = -1
                                Dim iID2 As Int16 = -1
                                Dim lID3 As Int32 = -1

                                Dim yName() As Byte = StringToBytes("Sell Order")

                                If iSO_TradeTypeID = ObjectType.eUnit Then
                                    Dim oUnit As Unit = GetEpicaUnit(lSO_TradeID)
                                    If oUnit Is Nothing = False Then
                                        oUnit.SaveObject()  'force save the object now
                                        lID1 = oUnit.ObjectID
                                        iID2 = oUnit.ObjTypeID
                                        yName = oUnit.EntityName
                                        oUnit.ParentObject = Nothing
                                        .TradePost.Owner.lMilitaryScore -= oUnit.EntityDef.CombatRating
                                        For Y As Int32 = 0 To .TradePost.lHangarUB
                                            If .TradePost.lHangarIdx(Y) = oUnit.ObjectID Then
                                                If .TradePost.oHangarContents(Y).ObjTypeID = oUnit.ObjTypeID Then
                                                    .TradePost.lHangarIdx(Y) = -1
                                                    .TradePost.oHangarContents(Y) = Nothing
                                                    Exit For
                                                End If
                                            End If
                                        Next Y
                                        If .blQuantity <> 0 Then .blQuantity = 0
                                    End If
                                ElseIf iSO_TradeTypeID < 0 Then
                                    If .TradePost.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                                        Dim oFac As Facility = .TradePost
                                        If oFac Is Nothing = False AndAlso (oFac.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                                            For Y As Int32 = 0 To oFac.lCargoUB
                                                If oFac.lCargoIdx(Y) <> -1 AndAlso oFac.oCargoContents(Y).ObjTypeID = ObjectType.eComponentCache Then
                                                    Dim oCache As ComponentCache = CType(oFac.oCargoContents(Y), ComponentCache)

                                                    If oCache.ComponentID = .lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(iSO_TradeTypeID) Then
                                                        lID1 = oCache.ComponentID
                                                        iID2 = -oCache.ComponentTypeID
                                                        lID3 = oCache.ComponentOwnerID

                                                        If oCache.GetComponent Is Nothing = False Then yName = oCache.GetComponent.GetTechName

                                                        oCache.Quantity -= CInt(blQty)

                                                        If .blQuantity > oCache.Quantity Then .blQuantity = oCache.Quantity
                                                    End If
                                                End If
                                            Next Y
                                        End If
                                    Else
                                        Dim oColony As Colony = .TradePost.ParentColony
                                        Dim blActQty As Int64 = blQty
                                        If oColony Is Nothing = False Then
                                            For Z As Int32 = 0 To oColony.ChildrenUB
                                                If oColony.lChildrenIdx(Z) > -1 Then
                                                    Dim oFac As Facility = oColony.oChildren(Z)
                                                    If oFac Is Nothing = False AndAlso oFac.ParentColony Is Nothing = False Then

                                                        For Y As Int32 = 0 To oFac.ParentColony.mlComponentCacheUB
                                                            If oFac.ParentColony.mlComponentCacheIdx(Y) > -1 Then
                                                                If oFac.ParentColony.mlComponentCacheID(Y) = glComponentCacheIdx(oFac.ParentColony.mlComponentCacheIdx(Y)) Then
                                                                    Dim oCache As ComponentCache = goComponentCache(oFac.ParentColony.mlComponentCacheIdx(Y))
                                                                    If oCache Is Nothing = False Then
                                                                        If oCache.ComponentID = .lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(iSO_TradeTypeID) Then
                                                                            lID1 = oCache.ComponentID
                                                                            iID2 = -oCache.ComponentTypeID
                                                                            lID3 = oCache.ComponentOwnerID

                                                                            If oCache.GetComponent Is Nothing = False Then yName = oCache.GetComponent.GetTechName

                                                                            Dim lTemp As Int32 = Math.Min(oCache.Quantity, CInt(blActQty))
                                                                            oCache.Quantity -= lTemp 'CInt(blQty)
                                                                            blActQty -= lTemp
                                                                            If .blQuantity > oCache.Quantity Then .blQuantity = oCache.Quantity
                                                                            If blActQty < 1 Then Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                            If blQty < 1 Then Exit For
                                                        Next Y
                                                        'For Y As Int32 = 0 To oFac.lCargoUB
                                                        '    If oFac.lCargoIdx(Y) <> -1 AndAlso oFac.oCargoContents(Y).ObjTypeID = ObjectType.eComponentCache Then
                                                        '        Dim oCache As ComponentCache = CType(oFac.oCargoContents(Y), ComponentCache)

                                                        '        If oCache.ComponentID = .lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(iSO_TradeTypeID) Then
                                                        '            lID1 = oCache.ComponentID
                                                        '            iID2 = -oCache.ComponentTypeID
                                                        '            lID3 = oCache.ComponentOwnerID

                                                        '            If oCache.GetComponent Is Nothing = False Then yName = oCache.GetComponent.GetTechName

                                                        '            Dim lTemp As Int32 = Math.Min(oCache.Quantity, CInt(blActQty))
                                                        '            oCache.Quantity -= lTemp 'CInt(blQty)
                                                        '            blActQty -= lTemp
                                                        '            If .blQuantity > oCache.Quantity Then .blQuantity = oCache.Quantity
                                                        '            If blActQty < 1 Then Exit For
                                                        '        End If
                                                        '    End If
                                                        'Next Y

                                                        If blActQty < 1 Then Exit For
                                                    End If
                                                End If
                                            Next Z
                                        End If
                                    End If
                                ElseIf iSO_TradeTypeID = ObjectType.eMineralCache Then
                                    'Ok, mineral caches
                                    Dim oMineralCache As MineralCache = GetEpicaMineralCache(lSO_TradeID)
                                    If oMineralCache Is Nothing = False Then
                                        lID1 = oMineralCache.oMineral.ObjectID
                                        iID2 = ObjectType.eMineralCache
                                        yName = oMineralCache.oMineral.MineralName
                                        oMineralCache.Quantity -= CInt(blQty)
                                        If .blQuantity > oMineralCache.Quantity Then .blQuantity = oMineralCache.Quantity
                                    End If
                                ElseIf iSO_TradeTypeID = ObjectType.ePlayerIntel OrElse iSO_TradeTypeID = ObjectType.ePlayerItemIntel OrElse iSO_TradeTypeID = ObjectType.ePlayerTechKnowledge Then
                                    lID1 = lSO_TradeID
                                    iID2 = iSO_TradeTypeID
                                    lID3 = .iExtTypeID
                                Else
                                    'TODO: what else?
                                    'for Component Designs...
                                    ' store the ComponentID in the ObjectID
                                    ' store the ComponentTypeID as the typeid
                                    ' store the componentownerid as the extendedid
                                End If


                                'Now, define oTD (tradedelivery) and initialize it
                                Dim lTDID As Int32 = CreateTradeDelivery(lID1, iID2, lID3, blQty, oTradePost.ObjectID, .TradePost.ObjectID, lTotalTime)
                                If lTDID <> -1 Then
                                    lTotalTime *= 30            'now in cycles
                                    Dim lEndCycle As Int32 = glCurrentCycle + lTotalTime
                                    Dim lHalfCycle As Int32 = glCurrentCycle + (lTotalTime \ 2)
                                    AddToQueue(lHalfCycle, QueueItemType.eTradeEventHalfTime, lTDID, ObjectType.eTrade, 1, -1, 0, 0, 0, 0)
                                    AddToQueue(lEndCycle, QueueItemType.eTradeEventFinal, lTDID, ObjectType.eTrade, 1, -1, 0, 0, 0, 0)

                                    'alert the player that the purchase was successful
                                    Dim yResp(22) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.ePurchaseSellOrder).CopyTo(yResp, 0)
                                    yResp(2) = 0            'successful purchase
                                    yName.CopyTo(yResp, 3)
                                    oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
                                    .TradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
                                Else
                                    'Alert the player that the purchase was unsuccessful
                                    Dim yResp(2) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.ePurchaseSellOrder).CopyTo(yResp, 0)
                                    yResp(2) = 255              'unknown reason
                                    oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
                                End If

                                If .blQuantity < 1 Then
                                    .TradePost.lTradePostSellSlotsUsed -= 1
                                    If .TradePost.ParentColony Is Nothing = False Then .TradePost.ParentColony.UpdateAllValues(-1)
                                    mySellOrderType(X) = MarketListType.eNotUsed
                                    moSellOrders(X).DeleteMe()
                                End If
                            Else
                                'alert the player that the purchase was unsuccessful
                                Dim yResp(2) As Byte
                                System.BitConverter.GetBytes(GlobalMessageCode.ePurchaseSellOrder).CopyTo(yResp, 0)
                                yResp(2) = 1            'Quantity unavailable
                                oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
                            End If
                        Else
                            oTradePost.Owner.SendPlayerMessage(oTradePost.ParentColony.GetLowResourcesMsg(ProductionType.eCredits, 0, 0, -1, -1), False, AliasingRights.eViewTreasury Or AliasingRights.eViewTrades)
                        End If

						Exit For
					End If
				End With
			End If
		Next X

	End Sub

	Public Sub HandleGetGTCList(ByRef yData() As Byte, ByVal lPlayerID As Int32)
		Dim yType As Byte = yData(2)
		Dim lTradePost As Int32 = System.BitConverter.ToInt32(yData, 3)

		Dim oFac As Facility = GetEpicaFacility(lTradePost)

		If oFac Is Nothing Then Return
		If oFac.Owner.ObjectID <> lPlayerID Then Return
		If oFac.yProductionType <> ProductionType.eTradePost Then Return

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

        Dim lExcludeList() As Int32 = Nothing
        Dim lExcludeListUB As Int32 = -1
        For X As Int32 = 0 To oPlayer.PlayerRelUB
            Dim oRel As PlayerRel = oPlayer.GetPlayerRelByIndex(X)
            If oRel Is Nothing = False AndAlso oRel.oThisPlayer Is Nothing = False Then
                If oRel.WithThisScore <= elRelTypes.eWar Then
                    lExcludeListUB += 1
                    ReDim Preserve lExcludeList(lExcludeListUB)
                    lExcludeList(lExcludeListUB) = oRel.oThisPlayer.ObjectID
                ElseIf oRel.oThisPlayer.GetPlayerRelScore(lPlayerID) <= elRelTypes.eWar Then
                    lExcludeListUB += 1
                    ReDim Preserve lExcludeList(lExcludeListUB)
                    lExcludeList(lExcludeListUB) = oRel.oThisPlayer.ObjectID
                End If
            End If
        Next

		'Figure out what is in range of this trade post based on the player's range values
		Dim yRng As Byte = oPlayer.oSpecials.yTradeBoardRange
        If yRng = 0 Then Return '0 indicates no range
        yRng = 3

		'Ok, now, make it easy for our lookups
		Dim lTradePostPlanetID As Int32 = -1
		Dim lTradePostSystemID As Int32 = -1

		With CType(oFac.ParentObject, Epica_GUID)
			If .ObjTypeID = ObjectType.ePlanet Then
				lTradePostPlanetID = .ObjectID
				lTradePostSystemID = CType(oFac.ParentObject, Planet).ParentSystem.ObjectID
			ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
				lTradePostPlanetID = -1
				lTradePostSystemID = .ObjectID
			ElseIf .ObjTypeID = ObjectType.eFacility Then
				With CType(CType(oFac.ParentObject, Facility).ParentObject, Epica_GUID)
					lTradePostSystemID = .ObjectID
				End With
			End If
		End With

		'Now, determine if we are dealing with a sell order lookup or a buy order lookup
		Dim bSellOrder As Boolean = (yType And MarketListType.eSellOrderShift) <> 0
		If bSellOrder = True Then
			'Ok, now, go through the sell order list
			If yType <> MarketListType.eSellOrderShift Then
				For X As Int32 = 0 To mlSellOrderUB
                    If mySellOrderType(X) = yType AndAlso (yRng = 3 OrElse IsSellOrderInTradePostRange(X, yRng, lTradePostPlanetID, lTradePostSystemID) = True) Then
                        If moSellOrders(X).TradePost Is Nothing = False Then
                            If moSellOrders(X).TradePost.Owner.yPlayerPhase <> oPlayer.yPlayerPhase Then Continue For
                            If oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso oPlayer.AccountStatus <> AccountStatusType.eActiveAccount Then Exit For
                            If moSellOrders(X).TradePost.ServerIndex > -1 Then
                                If glFacilityIdx(moSellOrders(X).TradePost.ServerIndex) <> moSellOrders(X).TradePost.ObjectID Then
                                    moSellOrders(X).DeleteMe()
                                    mySellOrderType(X) = MarketListType.eNotUsed
                                    Continue For
                                End If
                            End If

                            'check if the player is excluded
                            Dim bExclude As Boolean = False
                            For Y As Int32 = 0 To lExcludeListUB
                                If lExcludeList(Y) = moSellOrders(X).TradePost.Owner.ObjectID Then
                                    bExclude = True
                                    Exit For
                                End If
                            Next Y
                            If bExclude = True Then Continue For
                        End If
                        If moSellOrders(X).blQuantity > 0 Then
                            'Ok, send it
                            Dim yMsg(7 + SellOrder.l_MESSAGE_LENGTH) As Byte
                            Dim lPos As Int32 = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eGetGTCList).CopyTo(yMsg, lPos) : lPos += 2
                            System.BitConverter.GetBytes(lTradePost).CopyTo(yMsg, lPos) : lPos += 4
                            yMsg(lPos) = yType : lPos += 1

                            Dim yRouteType As Byte = 0          '0 = Planetary
                            If lTradePostPlanetID <> mlSellOrderPlanetID(X) Then
                                If lTradePostSystemID <> mlSellOrderSystemID(X) Then
                                    yRouteType = 2      'system to system
                                Else : yRouteType = 1   'intra system
                                End If
                            End If
                            yMsg(lPos) = yRouteType : lPos += 1

                            moSellOrders(X).GetObjAsString.CopyTo(yMsg, lPos)

                            'Now, send it
                            oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewTrades)
                        Else
                            Dim oTmp As SellOrder = moSellOrders(X)
                            mySellOrderType(X) = MarketListType.eNotUsed
                            moSellOrders(X) = Nothing
                            If oTmp.TradePost Is Nothing = False Then oTmp.TradePost.lTradePostSellSlotsUsed -= 1
                            oTmp.DeleteMe()
                        End If
                    End If
				Next X
			End If
		ElseIf yType <> 0 Then
			'Ok, return the buy orders
			For X As Int32 = 0 To mlBuyOrderUB
				If myBuyOrderType(X) = yType AndAlso (yRng = 3 OrElse IsBuyOrderInTradePostRange(X, yRng, lTradePostPlanetID, lTradePostSystemID) = True) Then
					If moBuyOrders(X).lAcceptedByID = -1 OrElse moBuyOrders(X).lAcceptedByID = oPlayer.ObjectID OrElse moBuyOrders(X).TradePost.Owner.ObjectID = oPlayer.ObjectID Then
                        'check if the player is excluded
                        Dim bExclude As Boolean = False
                        For Y As Int32 = 0 To lExcludeListUB
                            If lExcludeList(Y) = moBuyOrders(X).TradePost.Owner.ObjectID Then
                                bExclude = True
                                Exit For
                            End If
                        Next Y
                        If bExclude = True Then Continue For

                        'Ok, send it
						Dim yMsg(7 + BuyOrder.ml_OBJECT_MSG_LEN) As Byte
						Dim lPos As Int32 = 0
						System.BitConverter.GetBytes(GlobalMessageCode.eGetGTCList).CopyTo(yMsg, lPos) : lPos += 2
						System.BitConverter.GetBytes(lTradePost).CopyTo(yMsg, lPos) : lPos += 4
						yMsg(lPos) = yType : lPos += 1

						Dim yRouteType As Byte = 0			'0 = Planetary
						If lTradePostPlanetID <> mlBuyOrderPlanetID(X) Then
							If lTradePostSystemID <> mlBuyOrderSystemID(X) Then
								yRouteType = 2		'system to system
							Else : yRouteType = 1	'intra system
							End If
						End If
						yMsg(lPos) = yRouteType : lPos += 1

						moBuyOrders(X).GetObjAsString().CopyTo(yMsg, lPos)
						'Now, send it
						oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewTrades)
					End If
				End If
			Next X
		End If

		Dim yFinalMsg(6) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eGetGTCList).CopyTo(yFinalMsg, 0)
		System.BitConverter.GetBytes(lTradePost).CopyTo(yFinalMsg, 2)
		yFinalMsg(6) = 255
		oPlayer.SendPlayerMessage(yFinalMsg, False, AliasingRights.eViewTrades)
	End Sub

	Private Function IsSellOrderInTradePostRange(ByVal lSellOrderIndex As Int32, ByVal yRng As Byte, ByVal lTradePostPlanetID As Int32, ByVal lTradePostSystemID As Int32) As Boolean
        yRng = Math.Max(moSellOrders(lSellOrderIndex).TradePost.Owner.oSpecials.yTradeBoardRange, yRng)
        yRng = 3
		If yRng = 1 Then
			Return mlSellOrderPlanetID(lSellOrderIndex) = lTradePostPlanetID
		ElseIf yRng = 2 Then
			Return mlSellOrderSystemID(lSellOrderIndex) = lTradePostSystemID
		ElseIf yRng = 3 Then
			Return True
		Else
			Return False
		End If
		'yRng = 1 means intra planetary - can trade on same planet, so any sell orders that envirID = Fac.ParentID, EnvirTypeID = planet
		'yRng = 2 means intra system - can trade with anyone within same system... so any sell orders that SystemID = Fac.SystemID
		'yRng = 3 means all
	End Function

	Private Function IsBuyOrderInTradePostRange(ByVal lBuyOrderIndex As Int32, ByVal yRng As Byte, ByVal lTradePostPlanetID As Int32, ByVal lTradePostSystemID As Int32) As Boolean
		If yRng = 1 Then
			Return mlBuyOrderPlanetID(lBuyOrderIndex) = lTradePostPlanetID
		ElseIf yRng = 2 Then
			Return mlBuyOrderSystemID(lBuyOrderIndex) = lTradePostSystemID
		ElseIf yRng = 3 Then
			Return True
		Else
			Return False
		End If
	End Function

	Private Shared Function GetMarketType(ByVal lTradeID As Int32, ByVal iTradeTypeID As Int16, ByRef oTradePost As Facility, ByVal iExtTypeID As Int16) As Byte
		Select Case iTradeTypeID
			Case ObjectType.eMineralCache
                Dim oMineral As Mineral = GetEpicaMineral(lTradeID)
                If oMineral Is Nothing = False Then
                    If oMineral.lAlloyTechID <> -1 Then Return MarketListType.eAlloys Else Return MarketListType.eMinerals
                Else
                    Dim oMinCache As MineralCache = GetEpicaMineralCache(lTradeID)
                    If oMinCache Is Nothing = False Then
                        oMineral = oMinCache.oMineral
                        If oMineral Is Nothing = False Then
                            If oMineral.lAlloyTechID <> -1 Then Return MarketListType.eAlloys Else Return MarketListType.eMinerals
                        End If
                    End If
                    For X As Int32 = 0 To oTradePost.lCargoUB
                        If oTradePost.lCargoIdx(X) = lTradeID AndAlso oTradePost.oCargoContents(X).ObjTypeID = iTradeTypeID Then
                            oMineral = CType(oTradePost.oCargoContents(X), MineralCache).oMineral
                            If oMineral Is Nothing = False Then
                                If oMineral.lAlloyTechID <> -1 Then Return MarketListType.eAlloys Else Return MarketListType.eMinerals
                            End If
                            Exit For
                        End If
                    Next X
                End If
			Case ObjectType.ePlayerIntel
				Return MarketListType.ePlayerStats
			Case ObjectType.ePlayerItemIntel
				'TODO: Add more later, for now, just return the only one relevant
				Return MarketListType.eColonyLocation
			Case ObjectType.ePlayerTechKnowledge
				Return MarketListType.ePlayerComponentDesigns
			Case ObjectType.eComponentCache, Is < 0
				Dim iCompTypeID As Int16 = Math.Abs(iTradeTypeID)
				If iTradeTypeID < 0 Then iTradeTypeID = ObjectType.eComponentCache

                If oTradePost Is Nothing = False AndAlso oTradePost.ParentColony Is Nothing = False Then
                    For lFac As Int32 = 0 To oTradePost.ParentColony.ChildrenUB
                        If oTradePost.ParentColony.lChildrenIdx(lFac) > -1 Then
                            Dim oFac As Facility = oTradePost.ParentColony.oChildren(lFac)
                            If oFac Is Nothing = False AndAlso oFac.ParentColony Is Nothing = False Then
                                With oFac.ParentColony
                                    'Next lChild
                                    For X As Int32 = 0 To .mlComponentCacheUB
                                        If .mlComponentCacheIdx(X) > -1 Then
                                            If .mlComponentCacheID(X) = glComponentCacheIdx(.mlComponentCacheIdx(X)) Then
                                                Dim oCache As ComponentCache = goComponentCache(.mlComponentCacheIdx(X))
                                                If oCache Is Nothing = False Then
                                                    If oCache.ComponentID = lTradeID AndAlso oCache.ComponentTypeID = iCompTypeID Then
                                                        Select Case oCache.ComponentTypeID
                                                            Case ObjectType.eArmorTech
                                                                Return MarketListType.eArmorComponent
                                                            Case ObjectType.eEngineTech
                                                                Return MarketListType.eEngineComponent
                                                            Case ObjectType.eRadarTech
                                                                Return MarketListType.eRadarComponent
                                                            Case ObjectType.eShieldTech
                                                                Return MarketListType.eShieldComponent
                                                            Case ObjectType.eWeaponTech
                                                                Dim oOwner As Player = GetEpicaPlayer(oCache.ComponentOwnerID)
                                                                If oOwner Is Nothing Then oOwner = goInitialPlayer
                                                                If oOwner Is Nothing = False Then
                                                                    Dim oTech As BaseWeaponTech = CType(oOwner.GetTech(oCache.ComponentID, ObjectType.eWeaponTech), BaseWeaponTech)
                                                                    If oTech Is Nothing = False Then
                                                                        Select Case oTech.WeaponClassTypeID
                                                                            Case WeaponClassType.eBomb
                                                                                Return MarketListType.eBombComponent
                                                                            Case WeaponClassType.eEnergyBeam
                                                                                Return MarketListType.eBeamSolidComponent
                                                                            Case WeaponClassType.eEnergyPulse
                                                                                Return MarketListType.eBeamPulseComponent
                                                                            Case WeaponClassType.eMine
                                                                                Return MarketListType.eMineWeaponComponent
                                                                            Case WeaponClassType.eMissile
                                                                                Return MarketListType.eMissileComponent
                                                                            Case WeaponClassType.eProjectile
                                                                                Return MarketListType.eProjectileComponent
                                                                        End Select
                                                                    End If
                                                                End If
                                                        End Select
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next X
                                End With
                                'For X As Int32 = 0 To oFac.lCargoUB
                                '    If oFac.lCargoIdx(X) <> -1 AndAlso oFac.oCargoContents(X).ObjTypeID = iTradeTypeID Then
                                '        With CType(oFac.oCargoContents(X), ComponentCache)
                                '            If .ComponentID = lTradeID AndAlso iCompTypeID = .ComponentTypeID Then
                                '                Select Case .ComponentTypeID
                                '                    Case ObjectType.eArmorTech
                                '                        Return MarketListType.eArmorComponent
                                '                    Case ObjectType.eEngineTech
                                '                        Return MarketListType.eEngineComponent
                                '                    Case ObjectType.eRadarTech
                                '                        Return MarketListType.eRadarComponent
                                '                    Case ObjectType.eShieldTech
                                '                        Return MarketListType.eShieldComponent
                                '                    Case ObjectType.eWeaponTech
                                '                        Dim oOwner As Player = GetEpicaPlayer(.ComponentOwnerID)
                                '                        If oOwner Is Nothing Then oOwner = goInitialPlayer
                                '                        If oOwner Is Nothing = False Then
                                '                            Dim oTech As BaseWeaponTech = CType(oOwner.GetTech(.ComponentID, ObjectType.eWeaponTech), BaseWeaponTech)
                                '                            If oTech Is Nothing = False Then
                                '                                Select Case oTech.WeaponClassTypeID
                                '                                    Case WeaponClassType.eBomb
                                '                                        Return MarketListType.eBombComponent
                                '                                    Case WeaponClassType.eEnergyBeam
                                '                                        Return MarketListType.eBeamSolidComponent
                                '                                    Case WeaponClassType.eEnergyPulse
                                '                                        Return MarketListType.eBeamPulseComponent
                                '                                    Case WeaponClassType.eMine
                                '                                        Return MarketListType.eMineWeaponComponent
                                '                                    Case WeaponClassType.eMissile
                                '                                        Return MarketListType.eMissileComponent
                                '                                    Case WeaponClassType.eProjectile
                                '                                        Return MarketListType.eProjectileComponent
                                '                                End Select
                                '                            End If
                                '                        End If
                                '                End Select

                                '                Exit For
                                '            End If
                                '        End With

                                '    End If
                                'Next X
                            End If
                        End If
                    Next lFac

                End If
			Case ObjectType.eUnit
				Dim oUnit As Unit = GetEpicaUnit(lTradeID)
				If oUnit Is Nothing = False Then
					If (oUnit.EntityDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
						Return MarketListType.eLandBasedUnit
					ElseIf (oUnit.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
						Return MarketListType.eSpaceBasedUnit
					ElseIf (oUnit.EntityDef.yChassisType And ChassisType.eAtmospheric) <> 0 Then
						Return MarketListType.eAerialUnit
					Else : Return MarketListType.eNavalUnit
					End If
				End If
		End Select

		Return 0
	End Function

	Public Function CreateTradeDelivery(ByVal lID1 As Int32, ByVal iID2 As Int16, ByVal lID3 As Int32, ByVal blQty As Int64, ByVal lToTradePostID As Int32, ByVal lFromTradePostID As Int32, ByVal lSecondsForDelivery As Int32) As Int32
		'Now, create a new trade delivery
		Dim oNew As New TradeDelivery()
		With oNew
			.lID1 = lID1
			.iID2 = iID2
			.lID3 = lID3
			.blQty = blQty
			.lTradePostID = lToTradePostID
			.lSourceTradePostID = lFromTradePostID

			.dtDelivery = Now.AddSeconds(lSecondsForDelivery)
			.dtStartedOn = Now
		End With

		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To mlTradeDeliveryUB
			If myTradeDeliveryUsed(X) = 0 Then
				lIdx = X
				Exit For
			End If
		Next X
		If lIdx = -1 Then
			ReDim Preserve moTradeDelivery(mlTradeDeliveryUB + 1)
			ReDim Preserve myTradeDeliveryUsed(mlTradeDeliveryUB + 1)
			mlTradeDeliveryUB += 1
			lIdx = mlTradeDeliveryUB
		End If
		moTradeDelivery(lIdx) = oNew
		myTradeDeliveryUsed(lIdx) = 255

		Return lIdx
	End Function

    Public Shared Function GetTradeRouteDeliveryTime(ByRef oFromTradePost As Facility, ByRef oToTradePost As Facility, ByVal bDirectTrade As Boolean) As Int32
        If oFromTradePost Is Nothing OrElse oToTradePost Is Nothing Then Return 0

        'Ok, get from's Parent
        Dim oFromParent As Epica_GUID = CType(oFromTradePost, Epica_GUID)
        If oFromParent Is Nothing Then Return 0
        If oFromParent.ObjTypeID = ObjectType.eFacility Then oFromParent = CType(CType(oFromParent, Facility).ParentObject, Epica_GUID)
        'Get the To's parent
        Dim oToParent As Epica_GUID = CType(oToTradePost, Epica_GUID)
        If oToParent Is Nothing Then Return 0
        If oToParent.ObjTypeID = ObjectType.eFacility Then oToParent = CType(CType(oToParent, Facility).ParentObject, Epica_GUID)

        Dim lFromPlanetID As Int32 = -1
        Dim lFromSystemID As Int32 = -1
        Dim lToPlanetID As Int32 = -1
        Dim lToSystemID As Int32 = -1


        If oFromParent.ObjTypeID = ObjectType.ePlanet Then
            lFromPlanetID = oFromParent.ObjectID
            lFromSystemID = CType(oFromParent, Planet).ParentSystem.ObjectID
        ElseIf oFromParent.ObjTypeID = ObjectType.eSolarSystem Then
            lFromSystemID = oFromParent.ObjectID
            lFromPlanetID = -1
        Else : Return 0
        End If
        If oToParent.ObjTypeID = ObjectType.ePlanet Then
            lToPlanetID = oToParent.ObjectID
            lToSystemID = CType(oToParent, Planet).ParentSystem.ObjectID
        ElseIf oToParent.ObjTypeID = ObjectType.eSolarSystem Then
            lToSystemID = oToParent.ObjectID
            lToPlanetID = -1
        Else : Return 0
        End If

        Dim lSpeed As Int32 = Math.Max(oToTradePost.Owner.oSpecials.yTradeGTCSpeed, oFromTradePost.Owner.oSpecials.yTradeGTCSpeed)
        '0 = 200 (default), 1 = 400, 2 = 600, 3 = 800
        'If lSpeed = 3 Then
        lSpeed = (lSpeed * 200) + 200
        'Else
        '	lSpeed = (lSpeed * 200) + 100
        'End If
        lSpeed *= 30        'cycles per second
        If bDirectTrade = True Then lSpeed \= 4

        Dim lTotalTime As Int32 = 900

        Try
            'Ok, now, determine where we go from here
            If lFromPlanetID = lToPlanetID AndAlso lFromPlanetID <> -1 AndAlso lToPlanetID <> -1 Then
                'Ok, same planet
                Dim fDist As Single = Distance(oFromTradePost.LocX, oFromTradePost.LocZ, oToTradePost.LocX, oToTradePost.LocZ)
                lTotalTime = CInt(Math.Ceiling(fDist / lSpeed))
            ElseIf lFromSystemID = lToSystemID Then
                'Ok, same system... ok, does the from or to location start on a planet?
                Dim lAdditionalDistance As Int32 = 0
                Dim lFX As Int32 = oFromTradePost.LocX
                Dim lFZ As Int32 = oFromTradePost.LocZ
                Dim lTX As Int32 = oToTradePost.LocX
                Dim lTZ As Int32 = oToTradePost.LocZ
                If lFromPlanetID <> -1 Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lFromPlanetID)
                    If oPlanet Is Nothing = False Then
                        lFX = oPlanet.LocX
                        lFZ = oPlanet.LocZ
                        lAdditionalDistance += Math.Abs(oPlanet.LocY)
                    Else : lAdditionalDistance += 1000
                    End If
                End If
                If lToPlanetID <> -1 Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lToPlanetID)
                    If oPlanet Is Nothing = False Then
                        lTX = oPlanet.LocX
                        lTZ = oPlanet.LocZ
                        lAdditionalDistance += Math.Abs(oPlanet.LocY)
                    Else : lAdditionalDistance += 1000
                    End If
                End If

                'Ok, now, lFX, lFZ, lTX, lTZ have our locations
                Dim fDist As Single = Distance(lFX, lFZ, lTX, lTZ) + lAdditionalDistance
                lTotalTime = CInt(Math.Ceiling(fDist / lSpeed))
            Else
                'Ok, same galaxy... let's get the System to System distance
                Dim oSystemFrom As SolarSystem = GetEpicaSystem(lFromSystemID)
                Dim oSystemTo As SolarSystem = GetEpicaSystem(lToSystemID)
                If oSystemFrom Is Nothing OrElse oSystemTo Is Nothing Then Return 0 '???

                'Ok, this distance is different
                Dim fDX As Single = oSystemFrom.LocX - oSystemTo.LocX
                Dim fDY As Single = oSystemFrom.LocY - oSystemTo.LocY
                Dim fDZ As Single = oSystemFrom.LocZ - oSystemTo.LocZ
                fDX *= fDX
                fDY *= fDY
                fDZ *= fDZ

                Dim dDist As Double = Math.Sqrt(fDX + fDY + fDZ)
                'Now, that is in system sectors... so, we'll cheat and add 1 more sector for the inter-system movements
                dDist += 1.0F
                dDist *= 10000000
                'Now, if either location (or both) is a planet, add distance
                If lFromPlanetID <> -1 Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lFromPlanetID)
                    If oPlanet Is Nothing = False Then
                        dDist += Math.Abs(oPlanet.LocY)
                    Else : dDist += 1000.0F
                    End If
                End If
                If lToPlanetID <> -1 Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lToPlanetID)
                    If oPlanet Is Nothing = False Then
                        dDist += Math.Abs(oPlanet.LocY)
                    Else : dDist += 1000.0F
                    End If
                End If
                lTotalTime = CInt(dDist / lSpeed)
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "GetTradeRouteDeliveryTime: " & ex.Message)
        End Try

        Return lTotalTime
    End Function

	Public Sub TradeHalfTime(ByVal lTradeIdx As Int32)
		'TODO: Here is where any espionage/sabotage events would occur
	End Sub

	Public Sub TradeFinalTime(ByVal lTradeIdx As Int32)
		If mlTradeDeliveryUB >= lTradeIdx AndAlso myTradeDeliveryUsed(lTradeIdx) <> 0 Then
			'If moTradeDelivery(lTradeIdx).HandleItemDelivery() = True Then
			moTradeDelivery(lTradeIdx).HandleItemDelivery()
			myTradeDeliveryUsed(lTradeIdx) = 0
			moTradeDelivery(lTradeIdx).DeleteMe()
			'End If
		End If
	End Sub

	Public Sub AddNewDirectTrade(ByRef oTrade As DirectTrade)
		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To mlDirectTradeUB
			If mlDirectTradeIdx(X) = oTrade.ObjectID Then
				Return
			ElseIf lIdx = -1 AndAlso mlDirectTradeIdx(X) = -1 Then
				lIdx = X
			End If
		Next X
		If lIdx = -1 Then
			ReDim Preserve moDirectTrade(mlDirectTradeUB + 1)
			ReDim Preserve mlDirectTradeIdx(mlDirectTradeUB + 1)
			mlDirectTradeUB += 1
			lIdx = mlDirectTradeUB
		End If
		mlDirectTradeIdx(lIdx) = oTrade.ObjectID
		moDirectTrade(lIdx) = oTrade
	End Sub

	Public Function GetDirectTrade(ByVal lTradeID As Int32) As DirectTrade
		For X As Int32 = 0 To mlDirectTradeUB
			If mlDirectTradeIdx(X) = lTradeID Then Return moDirectTrade(X)
		Next X
		Return Nothing
	End Function

	Public Function SaveGTC() As Boolean
		'Save Sell Orders
		For X As Int32 = 0 To mlSellOrderUB
			If mySellOrderType(X) <> MarketListType.eNotUsed Then moSellOrders(X).SaveObject()
		Next X

		'Save Buy Orders... Ok, we clear out the buy order tables
		For X As Int32 = 0 To mlBuyOrderUB
			If myBuyOrderType(X) <> MarketListType.eNotUsed Then moBuyOrders(X).SaveObject()
		Next X

		'Save Trade Deliveries
		For X As Int32 = 0 To mlTradeDeliveryUB
			If myTradeDeliveryUsed(X) <> 0 Then moTradeDelivery(X).SaveObject()
		Next (X)

		'Save Direct Trades
		For X As Int32 = 0 To mlDirectTradeUB
			If mlDirectTradeIdx(X) <> -1 Then
				moDirectTrade(X).SaveObject()
			End If
		Next X
	End Function

	Public Function HandleGetOrderSpecifics(ByRef yData() As Byte) As Byte()
		Dim lTP_ID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim yOrderType As Byte = yData(6)
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 7)
		Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 11)
		Dim lFromTP_ID As Int32 = System.BitConverter.ToInt32(yData, 13)

		Dim oFromTP As Facility = GetEpicaFacility(lFromTP_ID)
		If oFromTP Is Nothing Then Return Nothing

		If yOrderType = 0 Then
			'SELL ORDERS
			For X As Int32 = 0 To mlSellOrderUB
				If mySellOrderType(X) <> MarketListType.eNotUsed Then
					With moSellOrders(X)
						If .TradePostID = lTP_ID AndAlso .lTradeID = lObjID AndAlso .iTradeTypeID = iTypeID Then

							Dim yResp(20) As Byte
							System.BitConverter.GetBytes(GlobalMessageCode.eGetOrderSpecifics).CopyTo(yResp, 0)
							System.BitConverter.GetBytes(lTP_ID).CopyTo(yResp, 2)
							yResp(6) = yOrderType
							System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 7)
							System.BitConverter.GetBytes(iTypeID).CopyTo(yResp, 11)

                            Dim lTime As Int32 = GetTradeRouteDeliveryTime(oFromTP, .TradePost, False)
							System.BitConverter.GetBytes(lTime).CopyTo(yResp, 13)
							System.BitConverter.GetBytes(.TradePost.Owner.ObjectID).CopyTo(yResp, 17)

							Return yResp
						End If
					End With
				End If
			Next X
		Else
			'BUY ORDERS
			For X As Int32 = 0 To mlBuyOrderUB
				If myBuyOrderType(X) <> MarketListType.eNotUsed Then
					With moBuyOrders(X)
						If .TradePostID = lTP_ID AndAlso .BuyOrderID = lObjID Then
                            Dim lTime As Int32 = GetTradeRouteDeliveryTime(oFromTP, .TradePost, False)
							Return .GetObjDetailMsg(lTP_ID, lTime)
						End If
					End With
				End If
			Next X
		End If

		Return Nothing
	End Function

	Public Function GetSellOrderItem(ByVal lTP_ID As Int32, ByVal lTradeID As Int32, ByVal iTradeTypeID As Int16) As SellOrder
		For X As Int32 = 0 To mlSellOrderUB
			If mySellOrderType(X) <> MarketListType.eNotUsed Then
				With moSellOrders(X)
					If .TradePostID = lTP_ID AndAlso .lTradeID = lTradeID AndAlso .iTradeTypeID = iTradeTypeID Then Return moSellOrders(X)
				End With
			End If
		Next X
		Return Nothing
	End Function

	Public Function GetBuyOrderItem(ByVal lBuyOrderID As Int32) As BuyOrder
		For X As Int32 = 0 To mlBuyOrderUB
			If myBuyOrderType(X) <> MarketListType.eNotUsed AndAlso moBuyOrders(X).BuyOrderID = lBuyOrderID Then Return moBuyOrders(X)
		Next X
		Return Nothing
	End Function

	Public Function GetBuyOrderAcceptedEmailBody(ByVal lBuyOrderID As Int32) As String
		For X As Int32 = 0 To mlBuyOrderUB
			If myBuyOrderType(X) <> MarketListType.eNotUsed Then
				Dim oSB As New System.Text.StringBuilder()

				With moBuyOrders(X)
					oSB.AppendLine("Buy Order Details")
					oSB.AppendLine()
					oSB.AppendLine("Tradepost: " & BytesToString(.TradePost.EntityName) & " (" & BytesToString(.TradePost.ParentColony.ColonyName) & ")")
					oSB.AppendLine("Deadline: " & GetDateFromNumber(.lDeadline).ToShortDateString)
					oSB.AppendLine("Escrow: " & .blEscrow.ToString("#,##0"))
					oSB.AppendLine("Payment: " & .blPaymentAmt.ToString("#,##0"))
				End With

				Return oSB.ToString
			End If
		Next X
		Return ""
	End Function

	Public Sub HandleDeliverBuyOrder(ByRef yData() As Byte, ByVal lPlayerID As Int32)
		Dim lPos As Int32 = 2
		Dim lBO_TPID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4		'where the buy order is from
		Dim lBuyOrderID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4		'the buy order
		Dim lTradePostID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4	'the tradepost fulfilling the buy order
		Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iCacheTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		Dim oTradePost As Facility = GetEpicaFacility(lTradePostID)
		If oTradePost Is Nothing Then Return
		If oTradePost.Owner.ObjectID <> lPlayerID Then Return
		If oTradePost.yProductionType <> ProductionType.eTradePost Then Return

		Dim oBuyOrder As BuyOrder = Nothing
		Dim lBuyOrderIdx As Int32 = -1
		For X As Int32 = 0 To mlBuyOrderUB
			If myBuyOrderType(X) <> MarketListType.eNotUsed AndAlso moBuyOrders(X).BuyOrderID = lBuyOrderID Then
				oBuyOrder = moBuyOrders(X)
				lBuyOrderIdx = X
				Exit For
			End If
		Next X
		If oBuyOrder Is Nothing Then Return

		If oBuyOrder.yBuyOrderState = 255 Then
			'Indicate unable to deliver because the deadline is exceeded
			Dim yResp(2) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eDeliverBuyOrder).CopyTo(yResp, 0)
			yResp(2) = 1		'deadline exceeded
			oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
			Return
		Else
            Dim lTime As Int32 = GalacticTradeSystem.GetTradeRouteDeliveryTime(oBuyOrder.TradePost, oTradePost, False)
			Dim dtTemp As Date = GetDateFromNumber(oBuyOrder.lDeadline)
			dtTemp = dtTemp.AddSeconds(-Math.Abs(lTime))
			Dim lAdjDeadline As Int32 = GetDateAsNumber(dtTemp)
			If lAdjDeadline < GetDateAsNumber(Now) Then
				Dim yResp(2) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eDeliverBuyOrder).CopyTo(yResp, 0)
				yResp(2) = 1		'deadline exceeded
				oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
				Return
			End If
		End If

		Dim bFound As Boolean = False
		For X As Int32 = 0 To oTradePost.lCargoUB
			If oTradePost.lCargoIdx(X) <> -1 Then
				If oTradePost.oCargoContents(X).ObjectID = lCacheID AndAlso oTradePost.oCargoContents(X).ObjTypeID = iCacheTypeID Then
					If oBuyOrder.ValidateCache(oTradePost.oCargoContents(X)) = True Then

						'Get the time for the delivery to pass into our CreateTradeDelivery
                        Dim lTime As Int32 = GetTradeRouteDeliveryTime(oTradePost, oBuyOrder.TradePost, False)
						'And our ID variables
						Dim lID1, lID3 As Int32
						Dim iID2 As Int16
						Dim yItemName() As Byte = Nothing

						'Ok, the cache validated, so let's pull the contents
						If oTradePost.oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
							With CType(oTradePost.oCargoContents(X), MineralCache)
								'Store the mineral now...
								lID1 = .oMineral.ObjectID
								iID2 = ObjectType.eMineralCache
								lID3 = -1
								yItemName = .oMineral.MineralName

								'Reduce the quantity...
								.Quantity = CInt(.Quantity - oBuyOrder.blQuantity)
							End With
						ElseIf oTradePost.oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
							With CType(oTradePost.oCargoContents(X), ComponentCache)
								'store the component now
								lID1 = .ComponentID
								iID2 = -.ComponentTypeID
								lID3 = .ComponentOwnerID
								yItemName = .GetComponent.GetTechName()

								'Reduce the quantity...
								.Quantity = CInt(.Quantity - oBuyOrder.blQuantity)
							End With
						Else
							'NOTE: add more here as needed, for now, cause an error
							LogEvent(LogEventType.PossibleCheat, "Entered Create Trade Delivery portion of HandleDeliverBuyOrder and Contents referenced are not Minerals or Components.")
							Return
						End If

						Dim lTDID As Int32 = CreateTradeDelivery(lID1, iID2, lID3, oBuyOrder.blQuantity, oBuyOrder.TradePostID, oTradePost.ObjectID, lTime)
						If lTDID <> -1 Then
							Dim lEndCycle As Int32 = glCurrentCycle + (lTime \ 30)
							Dim lHalfCycle As Int32 = glCurrentCycle + (lTime \ 60)
							AddToQueue(lHalfCycle, QueueItemType.eTradeEventHalfTime, lTDID, ObjectType.eTrade, 1, -1, 0, 0, 0, 0)
							AddToQueue(lEndCycle, QueueItemType.eTradeEventFinal, lTDID, ObjectType.eTrade, 1, -1, 0, 0, 0, 0)

							'alert the player that the delivery was successful
							Dim yResp(2) As Byte
							System.BitConverter.GetBytes(GlobalMessageCode.eDeliverBuyOrder).CopyTo(yResp, 0)
							yResp(2) = 0			'successful
							oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
						Else
							'Alert the player that the delivery was unsuccessful
							Dim yResp(2) As Byte
							System.BitConverter.GetBytes(GlobalMessageCode.eDeliverBuyOrder).CopyTo(yResp, 0)
							yResp(2) = 255			'unknown error
							oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
						End If

						'Now, if we are here, the delivering party gets the credits
						oTradePost.Owner.blCredits += oBuyOrder.blPaymentAmt + oBuyOrder.blEscrow
						'The delivery is on the way so the buyer gets their items... the buy order is finished
						oBuyOrder.yBuyOrderState = 255
						myBuyOrderType(lBuyOrderIdx) = MarketListType.eNotUsed
						oBuyOrder.yBuyOrderType = MarketListType.eNotUsed

						'Create our TradeHistorys
                        TradeHistory.CreateAndSaveBuyOrderDelivery(oBuyOrder, oTradePost.Owner.ObjectID, yItemName, lTime \ 30)

						'TODO: Alert players that the buy order is fulfilled (if needed)

						'Ok, go ahead and remove the buy order slot usage
						oBuyOrder.TradePost.lTradePostBuySlotsUsed -= 1I
                        If oBuyOrder.TradePost.ParentColony Is Nothing = False Then oBuyOrder.TradePost.ParentColony.UpdateAllValues(-1)

						bFound = True
					End If
					Exit For
				End If
			End If
		Next X

		If bFound = False Then
			Dim yResp(2) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eDeliverBuyOrder).CopyTo(yResp, 0)
			yResp(2) = 2			'unacceptable cargo
			oTradePost.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewTrades)
		End If

	End Sub

	Public Function HandleGetTradeDeliveries(ByVal lForPlayerID As Int32) As Byte()
		Dim yMsg() As Byte = Nothing

		Dim lCnt As Int32 = 0

		For X As Int32 = 0 To mlTradeDeliveryUB
			If myTradeDeliveryUsed(X) <> 0 Then
                If moTradeDelivery(X).oTargetTradePost Is Nothing = False AndAlso moTradeDelivery(X).oTargetTradePost.Owner.ObjectID = lForPlayerID Then lCnt += 1
			End If
		Next X

		'Ok, dim our msg
		Try
			ReDim yMsg(7 + (lCnt * TradeDelivery.ml_TRADE_DELIVERY_ITEM_MSG_LEN))
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eGetTradeDeliveries).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lForPlayerID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(CShort(lCnt)).CopyTo(yMsg, lPos) : lPos += 2
			For X As Int32 = 0 To mlTradeDeliveryUB
				If myTradeDeliveryUsed(X) <> 0 Then
                    If moTradeDelivery(X).oTargetTradePost Is Nothing = False AndAlso moTradeDelivery(X).oTargetTradePost.Owner.ObjectID = lForPlayerID Then
                        moTradeDelivery(X).GetTradeDeliveryPackageItem.CopyTo(yMsg, lPos) : lPos += TradeDelivery.ml_TRADE_DELIVERY_ITEM_MSG_LEN
                    End If
				End If
			Next X
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "HandleGetTradeDeliveries: " & ex.Message)
			yMsg = Nothing
		End Try
		Return yMsg
	End Function

    Public Sub SendPlayerDirectTrades(ByVal lPlayerID As Int32, ByRef oSocket As NetSock)
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        'For X As Int32 = 0 To mlDirectTradeUB
        '	If mlDirectTradeIdx(X) <> -1 Then
        '		If moDirectTrade(X).oPlayer1TP.PlayerID = lPlayerID OrElse moDirectTrade(X).oPlayer2TP.PlayerID = lPlayerID Then
        '			Dim yMsg() As Byte = goMsgSys.GetAddObjectMessage(moDirectTrade(X), GlobalMessageCode.eAddObjectCommand)
        '			If yMsg Is Nothing = False AndAlso yMsg.Length > 2 Then oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewTrades)
        '		End If
        '	End If
        'Next X

        Dim yCache(200000) As Byte
        Dim yFinal() As Byte
        Dim lPos As Int32
        Dim lSingleMsgLen As Int32
        Dim yTemp() As Byte

        lPos = 0
        lSingleMsgLen = -1
        For Y As Int32 = 0 To mlDirectTradeUB
            If mlDirectTradeIdx(Y) <> -1 AndAlso (moDirectTrade(Y).oPlayer1TP.PlayerID = lPlayerID OrElse moDirectTrade(Y).oPlayer2TP.PlayerID = lPlayerID) AndAlso moDirectTrade(Y).TradeState <> DirectTrade.eTradeStateValues.TradeRejected Then
                yTemp = goMsgSys.GetAddObjectMessage(moDirectTrade(Y), GlobalMessageCode.eAddObjectCommand)
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
            If oSocket Is Nothing = False AndAlso oSocket.IsConnected = True Then
                Try
                    oSocket.SendLenAppendedData(yFinal)
                Catch
                End Try
            End If
            'oplayer.SendLenAppendedData(yFinal)
        End If
    End Sub

	'returns the number of items being traded matching these values
	Public Function GetItemsBeingTraded(ByVal lTradepostID As Int32, ByVal lItemID As Int32, ByVal iTypeID As Int16) As Int64
		For X As Int32 = 0 To mlSellOrderUB
			If mySellOrderType(X) <> MarketListType.eNotUsed Then
				If moSellOrders(X).TradePostID = lTradepostID Then
					If moSellOrders(X).lTradeID = lItemID AndAlso moSellOrders(X).iTradeTypeID = iTypeID Then
						Return moSellOrders(X).blQuantity
					End If
				End If
			End If
		Next X
		Return 0
	End Function

    'Public Function GetTradepostTradeCargo(ByVal lTradepostID As Int32) As Int64
    '	Dim blResult As Int64 = 0
    '	For X As Int32 = 0 To mlSellOrderUB
    '		If mySellOrderType(X) <> MarketListType.eNotUsed Then
    '			If moSellOrders(X).TradePostID = lTradepostID Then
    '				Dim iTmpTypeID As Int16 = moSellOrders(X).iTradeTypeID
    '				If iTmpTypeID = ObjectType.eUnit Then
    '					Dim blTemp As Int64 = moSellOrders(X).GetQuantityInStock
    '					If blTemp <> 0 Then
    '						Dim oUnit As Unit = GetEpicaUnit(moSellOrders(X).lTradeID)
    '						If oUnit Is Nothing = False AndAlso oUnit.EntityDef Is Nothing = False Then
    '							blResult += oUnit.EntityDef.HullSize
    '						End If
    '					End If
    '				ElseIf iTmpTypeID <> ObjectType.ePlayerIntel AndAlso iTmpTypeID <> ObjectType.ePlayerItemIntel AndAlso iTmpTypeID <> ObjectType.ePlayerTechKnowledge Then
    '					Dim blTemp As Int64 = moSellOrders(X).GetQuantityInStock
    '					moSellOrders(X).blQuantity = Math.Min(blTemp, moSellOrders(X).blQuantity)
    '					blResult += moSellOrders(X).blQuantity
    '				End If
    '			End If
    '		End If
    '	Next X
    '	Return blResult
    'End Function

	Public Function GetNumberOfIntelItemsBeingSold(ByVal lTradepostID As Int32) As Int32
		Dim lCnt As Int32 = 0
		For X As Int32 = 0 To mlSellOrderUB
			If mySellOrderType(X) <> MarketListType.eNotUsed Then
				If moSellOrders(X).TradePostID = lTradepostID Then
					Dim iTmpTypeID As Int16 = moSellOrders(X).iTradeTypeID
					If iTmpTypeID = ObjectType.ePlayerIntel OrElse iTmpTypeID = ObjectType.ePlayerItemIntel OrElse iTmpTypeID = ObjectType.ePlayerTechKnowledge Then
						lCnt += 1
					End If
				End If
			End If
		Next X
		Return lCnt
	End Function
	Public Sub FillWithIntelItemsBeingsold(ByRef yData() As Byte, ByRef lPos As Int32, ByVal lTradePostID As Int32)
		Try
			For X As Int32 = 0 To mlSellOrderUB
				If mySellOrderType(X) <> MarketListType.eNotUsed Then
					If moSellOrders(X).TradePostID = lTradePostID Then
						Dim iTmpTypeID As Int16 = moSellOrders(X).iTradeTypeID
						If iTmpTypeID = ObjectType.ePlayerIntel OrElse iTmpTypeID = ObjectType.ePlayerItemIntel OrElse iTmpTypeID = ObjectType.ePlayerTechKnowledge Then

							With moSellOrders(X)
								System.BitConverter.GetBytes(.lTradeID).CopyTo(yData, lPos) : lPos += 4
								System.BitConverter.GetBytes(.iTradeTypeID).CopyTo(yData, lPos) : lPos += 2
								System.BitConverter.GetBytes(.iExtTypeID).CopyTo(yData, lPos) : lPos += 2
								System.BitConverter.GetBytes(.blPrice).CopyTo(yData, lPos) : lPos += 8
							End With

						End If
					End If
				End If
			Next X
		Catch
		End Try
	End Sub

    Public Sub ClearAllTradeDeliveries(ByVal lTP_ID As Int32)
        Dim lCurUB As Int32 = -1
        If myTradeDeliveryUsed Is Nothing = False Then lCurUB = Math.Min(mlTradeDeliveryUB, myTradeDeliveryUsed.GetUpperBound(0))
        Try
            For X As Int32 = 0 To lCurUB
                If myTradeDeliveryUsed(X) <> 0 Then
                    Dim oTD As TradeDelivery = moTradeDelivery(X)
                    If oTD Is Nothing = False Then
                        Dim oFac As Facility = oTD.oTargetTradePost
                        If oFac Is Nothing = False AndAlso oFac.ObjectID = lTP_ID Then
                            oTD.ClearTargetTradepost()
                        End If
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub

    Public Sub CancelSellOrder(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal lTP_ID As Int32)
        Dim lCurUB As Int32 = -1
        If mySellOrderType Is Nothing = False Then lCurUB = Math.Min(mlSellOrderUB, mySellOrderType.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mySellOrderType(X) <> MarketListType.eNotUsed Then
                Dim oSellOrder As SellOrder = moSellOrders(X)
                If oSellOrder Is Nothing = False Then
                    If oSellOrder.TradePostID = lTP_ID AndAlso oSellOrder.lTradeID = lObjectID AndAlso oSellOrder.iTradeTypeID = iObjTypeID Then
                        'Ok, cancel me
                        mySellOrderType(X) = MarketListType.eNotUsed
                        moSellOrders(X) = Nothing
                        If oSellOrder.TradePost Is Nothing = False Then oSellOrder.TradePost.lTradePostSellSlotsUsed -= 1
                        oSellOrder.DeleteMe()
                    End If
                End If
            End If
        Next X
    End Sub

    Public Sub CancelPlayersOrders(ByVal lPlayerID As Int32)
        Dim lCurUB As Int32 = -1
        Try
            If mySellOrderType Is Nothing = False Then lCurUB = Math.Min(mlSellOrderUB, mySellOrderType.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If mySellOrderType(X) <> MarketListType.eNotUsed Then
                    Dim oSellOrder As SellOrder = moSellOrders(X)
                    If oSellOrder Is Nothing = False Then
                        If oSellOrder.TradePost Is Nothing = False AndAlso oSellOrder.TradePost.Owner.ObjectID = lPlayerID Then
                            mySellOrderType(X) = MarketListType.eNotUsed
                            moSellOrders(X) = Nothing
                            If oSellOrder.TradePost Is Nothing = False Then oSellOrder.TradePost.lTradePostSellSlotsUsed -= 1
                            oSellOrder.DeleteMe()
                        End If
                    End If
                End If
            Next X
        Catch
        End Try

        Try
            lCurUB = -1
            If myBuyOrderType Is Nothing = False Then lCurUB = Math.Min(mlBuyOrderUB, myBuyOrderType.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If myBuyOrderType(X) <> MarketListType.eNotUsed Then
                    Dim oBuyOrder As BuyOrder = moBuyOrders(X)
                    If oBuyOrder Is Nothing = False Then
                        If oBuyOrder.TradePost Is Nothing = False AndAlso oBuyOrder.TradePost.Owner.ObjectID = lPlayerID Then
                            myBuyOrderType(X) = MarketListType.eNotUsed
                            moBuyOrders(X) = Nothing
                            If oBuyOrder.TradePost Is Nothing = False Then oBuyOrder.TradePost.lTradePostBuySlotsUsed -= 1
                            oBuyOrder.DeleteMe()
                        End If
                    End If
                End If
            Next X
        Catch
        End Try

        lCurUB = -1
        If myTradeDeliveryUsed Is Nothing = False Then lCurUB = Math.Min(mlTradeDeliveryUB, myTradeDeliveryUsed.GetUpperBound(0))
        Try
            For X As Int32 = 0 To lCurUB
                If myTradeDeliveryUsed(X) <> 0 Then
                    Dim oTD As TradeDelivery = moTradeDelivery(X)
                    If oTD Is Nothing = False Then
                        Dim oFac As Facility = oTD.oTargetTradePost
                        If oFac Is Nothing = False AndAlso oFac.Owner Is Nothing = False AndAlso oFac.Owner.ObjectID = lPlayerID Then
                            myTradeDeliveryUsed(X) = 0
                            oTD.ClearTargetTradepost()
                        End If
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub

    Public Function NumberOfMarketTypesForSell(ByVal lTradepostID As Int32) As Int32
        Dim lCnts(37) As Int32
        Dim lCurUB As Int32 = -1
        If mySellOrderType Is Nothing = False Then lCurUB = Math.Min(mlSellOrderUB, mySellOrderType.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mySellOrderType(X) <> MarketListType.eNotUsed Then
                Dim oSellOrder As SellOrder = moSellOrders(X)
                If oSellOrder Is Nothing = False AndAlso oSellOrder.TradePostID = lTradepostID Then
                    Dim lVal As Int32 = mySellOrderType(X)
                    If (lVal And MarketListType.eSellOrderShift) <> 0 Then lVal = lVal Xor MarketListType.eSellOrderShift
                    If lVal > lCnts.GetUpperBound(0) Then Continue For
                    lCnts(lVal) += 1
                End If
            End If
        Next X
        'Now, go back thru our counts
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lCnts.GetUpperBound(0)
            If lCnts(X) > 0 Then lCnt += 1
        Next X
        Return lCnt
    End Function

    Public Sub AddMineralCacheToNPCTradepost(ByRef oCache As MineralCache)
        Try
            If True = True Then Return
            'ok, first, get the NPCTradepost relevant to this cache
            If oCache.ParentObject Is Nothing = False Then
                Dim iTemp As Int16 = CType(oCache.ParentObject, Epica_GUID).ObjTypeID
                Dim oSys As SolarSystem = Nothing
                If iTemp = ObjectType.eSolarSystem Then
                    oSys = CType(oCache.ParentObject, SolarSystem)
                ElseIf iTemp = ObjectType.ePlanet Then
                    oSys = CType(oCache.ParentObject, Planet).ParentSystem
                Else : Return
                End If
                If oSys Is Nothing Then Return

                'Now, get the tradepost of that system
                Dim oTP As Facility = oSys.NPCTradepost
                If oTP Is Nothing Then Return

                'Now, add the cache to the cargo of the tradepost
                Dim oCargoCache As MineralCache = Nothing
                For X As Int32 = 0 To oTP.lCargoUB
                    If oTP.lCargoIdx(X) > -1 Then
                        If oTP.oCargoContents(X).ObjTypeID = oCache.ObjTypeID Then
                            With CType(oTP.oCargoContents(X), MineralCache)
                                If .oMineral.ObjectID = oCache.oMineral.ObjectID Then
                                    .Quantity += oCache.Quantity
                                    oCargoCache = CType(oTP.oCargoContents(X), MineralCache)
                                    Exit For
                                End If
                            End With
                        End If
                    End If
                Next X
                If oCargoCache Is Nothing Then
                    oCargoCache = oTP.AddMineralCacheToCargo(oCache.oMineral.ObjectID, oCache.Quantity)
                End If

                'then, determine the cost per item for sale
                Dim blCalcPricePer As Int64 = oCache.oMineral.MineralValue

                Try
                    Dim blCostEach As Int64 = 0
                    Dim lCount As Int32 = 0
                    Dim lUB As Int32 = -1
                    If glFacilityIdx Is Nothing = False Then lUB = Math.Min(glFacilityIdx.GetUpperBound(0), glFacilityUB)

                    Dim lBaseMineralID As Int32 = oCache.oMineral.ObjectID
                    For X As Int32 = 0 To lUB
                        If glFacilityIdx(X) > -1 Then
                            Dim oFac As Facility = goFacility(X)
                            If oFac Is Nothing = False AndAlso oFac.oMiningBid Is Nothing = False Then
                                If oFac.oMiningBid.oMineralCache Is Nothing = False AndAlso oFac.oMiningBid.oMineralCache.oMineral.ObjectID = lBaseMineralID Then
                                    blCostEach += oFac.oMiningBid.lMaxDaysSold
                                    lCount += 1
                                End If
                            End If
                        End If
                    Next X
                    If lCount > 0 Then blCalcPricePer += (blCostEach \ lCount)
                Catch
                End Try

                'then, add the cache onto the open market
                Dim oNew As New SellOrder
                With oNew
                    .blPrice = blCalcPricePer
                    .blQuantity = oCargoCache.Quantity
                    .iExtTypeID = CShort(oCargoCache.oMineral.ObjectID)
                    .iTradeTypeID = ObjectType.eMineralCache
                    .lTradeID = oCargoCache.ObjectID
                    .TradePostID = oTP.ObjectID
                    oCargoCache.oMineral.MineralName.CopyTo(.yItemName, 0)
                End With
                AddNPCSellOrder(oNew)

            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "AddMineralCacheToNPCTradepost: " & ex.Message)
        End Try

    End Sub
    Public Sub AddComponentCacheToNPCTradepost(ByRef oCache As ComponentCache)
        Try
            If True = True Then Return
            'ok, first, get the NPCTradepost relevant to this cache
            If oCache.ParentObject Is Nothing = False Then
                Dim iTemp As Int16 = CType(oCache.ParentObject, Epica_GUID).ObjTypeID
                Dim oSys As SolarSystem = Nothing
                If iTemp = ObjectType.eSolarSystem Then
                    oSys = CType(oCache.ParentObject, SolarSystem)
                ElseIf iTemp = ObjectType.ePlanet Then
                    oSys = CType(oCache.ParentObject, Planet).ParentSystem
                Else : Return
                End If
                If oSys Is Nothing Then Return

                'Now, get the tradepost of that system
                Dim oTP As Facility = oSys.NPCTradepost
                If oTP Is Nothing Then Return

                'Now, add the cache to the cargo of the tradepost
                Dim oCargoCache As ComponentCache = Nothing
                For X As Int32 = 0 To oTP.lCargoUB
                    If oTP.lCargoIdx(X) > -1 Then
                        If oTP.oCargoContents(X).ObjTypeID = oCache.ObjTypeID Then
                            With CType(oTP.oCargoContents(X), ComponentCache)
                                If .ComponentID = oCache.ComponentID AndAlso .ComponentTypeID = oCache.ComponentTypeID Then
                                    .Quantity += oCache.Quantity
                                    oCargoCache = CType(oTP.oCargoContents(X), ComponentCache)
                                    Exit For
                                End If
                            End With
                        End If
                    End If
                Next X
                If oCargoCache Is Nothing Then
                    oCargoCache = oTP.AddComponentCacheToCargo(oCache.ComponentID, oCache.ComponentTypeID, oCache.Quantity, oCache.ComponentOwnerID)
                End If

                'then, determine the cost per item for sale
                Dim blCalcPricePer As Int64 = 0

                'then, add the cache onto the open market
                Dim oNew As New SellOrder
                With oNew
                    .blPrice = blCalcPricePer
                    .blQuantity = oCargoCache.Quantity
                    .iExtTypeID = 0
                    .iTradeTypeID = -oCargoCache.ComponentTypeID
                    .lTradeID = oCargoCache.ComponentID
                    .TradePostID = oTP.ObjectID
                    oCargoCache.GetComponent.GetTechName().CopyTo(.yItemName, 0)
                End With
                AddNPCSellOrder(oNew)

            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "AddComponentCacheToNPCTradepost: " & ex.Message)
        End Try
    End Sub
    Private Sub AddNPCSellOrder(ByRef oNew As SellOrder)
        Try
            With oNew
                If .TradePost Is Nothing Then Return
                If .TradePost.yProductionType <> ProductionType.eTradePost Then Return
            End With

            'Ok, attempt to find our trade object already in the array
            Dim bPreExist As Boolean = False
            Dim lTempPreExistID As Int32 = oNew.lTradeID
            If oNew.iTradeTypeID < 0 Then
                With oNew.TradePost
                    Dim bFound As Boolean = False
                    For X As Int32 = 0 To .lCargoUB
                        If .lCargoIdx(X) = oNew.lTradeID AndAlso .oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
                            With CType(.oCargoContents(X), ComponentCache)
                                If .ComponentTypeID = Math.Abs(oNew.iTradeTypeID) Then
                                    lTempPreExistID = .ComponentID
                                    bFound = True
                                    Exit For
                                End If
                            End With
                        End If
                    Next X
                    If bFound = False Then Return ' False
                End With
            End If
            For X As Int32 = 0 To mlSellOrderUB
                If mySellOrderType(X) <> MarketListType.eNotUsed AndAlso moSellOrders(X).lTradeID = lTempPreExistID AndAlso moSellOrders(X).TradePostID = oNew.TradePostID AndAlso moSellOrders(X).iTradeTypeID = oNew.iTradeTypeID Then
                    'Ok, already found...
                    bPreExist = True

                    moSellOrders(X).blQuantity = oNew.blQuantity
                    moSellOrders(X).blPrice = oNew.blPrice
                    Exit For
                End If
            Next X

            If bPreExist = False AndAlso (oNew.blQuantity < 0 OrElse oNew.blPrice < 0) Then Return 'False

            'Ok, if we are here, place oNew within the oSellOrder list
            Dim lPlanetID As Int32 = -1
            Dim lSystemID As Int32 = -1
            With CType(oNew.TradePost.ParentObject, Epica_GUID)
                If .ObjTypeID = ObjectType.ePlanet Then
                    lPlanetID = .ObjectID
                    lSystemID = CType(oNew.TradePost.ParentObject, Planet).ParentSystem.ObjectID
                ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                    lPlanetID = -1
                    lSystemID = .ObjectID
                ElseIf .ObjTypeID = ObjectType.eFacility Then
                    With CType(CType(oNew.TradePost.ParentObject, Facility).ParentObject, Epica_GUID)
                        lSystemID = .ObjectID
                    End With
                End If
            End With

            If bPreExist = False Then
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To mlSellOrderUB
                    If mySellOrderType(X) = MarketListType.eNotUsed Then
                        lIdx = X
                        Exit For
                    End If
                Next X
                If lIdx = -1 Then
                    Dim lTmpUB As Int32 = mlSellOrderUB + 1
                    ReDim Preserve mySellOrderType(lTmpUB)
                    ReDim Preserve moSellOrders(lTmpUB)
                    ReDim Preserve mlSellOrderPlanetID(lTmpUB)
                    ReDim Preserve mlSellOrderSystemID(lTmpUB)
                    mlSellOrderUB += 1
                    lIdx = lTmpUB
                End If

                moSellOrders(lIdx) = oNew
                mlSellOrderPlanetID(lIdx) = lPlanetID
                mlSellOrderSystemID(lIdx) = lSystemID
                Dim yType As Byte = GetMarketType(oNew.lTradeID, oNew.iTradeTypeID, oNew.TradePost, oNew.iExtTypeID)
                If yType = 0 Then Return 'False
                mySellOrderType(lIdx) = CType(yType Or MarketListType.eSellOrderShift, MarketListType)

                'Increment our sell slots used
                'oNew.TradePost.lTradePostSellSlotsUsed += 1
                'If oNew.TradePost.ParentColony Is Nothing = False Then oNew.TradePost.ParentColony.UpdateAllValues(-1)
                'oNew.TradePost.Owner.TestCustomTitlePermissions_Trade()
            End If

            'bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "AddNPCSellOrder: " & ex.Message)
            'bResult = False
        End Try
        'Return bResult
    End Sub

    Public Sub ReRegisterTradeDeliveries(ByVal oTradepost As Facility)
        Try
            If oTradepost.ParentObject Is Nothing Then Return
            Dim lPID As Int32 = CType(oTradepost.ParentObject, Epica_GUID).ObjectID
            Dim iPTypeID As Int16 = CType(oTradepost.ParentObject, Epica_GUID).ObjTypeID
            Dim lOwnerID As Int32 = oTradepost.Owner.ObjectID

            For X As Int32 = 0 To mlTradeDeliveryUB
                If myTradeDeliveryUsed(X) > 0 Then
                    Dim oTD As TradeDelivery = moTradeDelivery(X)
                    If oTD Is Nothing = False Then
                        If oTD.oTargetTradePost Is Nothing Then
                            If oTD.lOriginalDestEnvirID = lPID AndAlso oTD.iOriginalDestEnvirTypeID = iPTypeID AndAlso oTD.lOriginalDestPlayerID = lOwnerID Then
                                oTD.lTradePostID = oTradepost.ObjectID
                            End If
                        End If
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub

    Public Sub ClearOrdersAboutIntelOfPlayer(ByVal lID As Int32)
        For X As Int32 = 0 To mlSellOrderUB
            If mySellOrderType(X) <> MarketListType.eNotUsed Then
                If moSellOrders(X) Is Nothing = False Then
                    'ePlayerIntel - tradeid
                    'ePlayerItemIntel - get oPII of the item and determine ownership
                    'ePlayerTechKnowledge - get oPTK of the item and determine ownership
                    If moSellOrders(X).iTradeTypeID = ObjectType.ePlayerIntel Then
                        If moSellOrders(X).lTradeID = lID Then
                            If moSellOrders(X).TradePost Is Nothing = False Then moSellOrders(X).TradePost.lTradePostSellSlotsUsed -= 1
                            moSellOrders(X).DeleteMe()
                            mySellOrderType(X) = MarketListType.eNotUsed
                        End If
                    ElseIf moSellOrders(X).iTradeTypeID = ObjectType.ePlayerItemIntel Then
                        If moSellOrders(X).TradePost Is Nothing = False Then
                            Dim oPII As PlayerItemIntel = moSellOrders(X).TradePost.Owner.GetPlayerItemIntel(moSellOrders(X).lTradeID, moSellOrders(X).iExtTypeID, lID)
                            If oPII Is Nothing = False Then
                                moSellOrders(X).TradePost.lTradePostSellSlotsUsed -= 1
                                moSellOrders(X).DeleteMe()
                                mySellOrderType(X) = MarketListType.eNotUsed
                            End If
                        Else
                            moSellOrders(X).DeleteMe()
                            mySellOrderType(X) = MarketListType.eNotUsed
                        End If
                    ElseIf moSellOrders(X).iTradeTypeID = ObjectType.ePlayerTechKnowledge Then
                        If moSellOrders(X).TradePost Is Nothing = False Then
                            Dim oPTK As PlayerTechKnowledge = moSellOrders(X).TradePost.Owner.GetPlayerTechKnowledge(moSellOrders(X).lTradeID, moSellOrders(X).iExtTypeID)
                            If oPTK Is Nothing = False AndAlso oPTK.oTech Is Nothing = False AndAlso oPTK.oTech.Owner Is Nothing = False AndAlso oPTK.oTech.Owner.ObjectID = lID Then
                                oPTK.yArchived = 255
                                oPTK.SaveObject()
                                moSellOrders(X).TradePost.lTradePostSellSlotsUsed -= 1
                                moSellOrders(X).DeleteMe()
                                mySellOrderType(X) = MarketListType.eNotUsed
                            End If
                        Else
                            moSellOrders(X).DeleteMe()
                            mySellOrderType(X) = MarketListType.eNotUsed
                        End If
                    End If
                End If
            End If
        Next
    End Sub

End Class