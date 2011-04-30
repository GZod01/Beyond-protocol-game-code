Option Strict On

Public Class TradePlayer
    'PK
    Public TradeID As Int32
    Public PlayerID As Int32
    'END OF PK

    Public Notes() As Byte      'could be nothing!

    Public lTradePostID As Int32 = -1

    Public oItems() As TradePlayerItem
    Public lItemIdx() As Int32
    Public lItemUB As Int32 = -1

    Private moTradePost As Facility = Nothing
    Public ReadOnly Property TradePost() As Facility
        Get
            If moTradePost Is Nothing OrElse moTradePost.ObjectID <> lTradePostID Then
                If lTradePostID <> -1 Then moTradePost = GetEpicaFacility(lTradePostID)
            End If
            Return moTradePost
        End Get
    End Property

    Public Sub AddItemQuantity(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal blQuantity As Int64, ByVal lExtendedID As Int32, ByVal lExtended2ID As Int32)
        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1

        For X As Int32 = 0 To lItemUB
            If lItemIdx(X) = lObjectID AndAlso oItems(X).ObjTypeID = iObjTypeID Then
                lIdx = X
                Exit For
            ElseIf lFirstIdx = -1 AndAlso lItemIdx(X) = -1 Then
                lFirstIdx = X
            End If
        Next X

        'Did we find it?
        If lIdx = -1 Then
            'No, did we find an empty slot?
            If lFirstIdx = -1 Then
                'No, redim
                lItemUB += 1
                ReDim Preserve oItems(lItemUB)
                ReDim Preserve lItemIdx(lItemUB)
                lIdx = lItemUB
            Else : lIdx = lFirstIdx
            End If

            With oItems(lIdx)
                .ObjectID = lObjectID
                .ObjTypeID = iObjTypeID
                '.TP_ID = Me.TP_ID
                .PlayerID = Me.PlayerID
                .TradeID = Me.TradeID
                .lExtendedID = lExtendedID
                .lExtended2ID = lExtended2ID
                .Quantity = 0
            End With
            lItemIdx(lIdx) = lObjectID
        End If

        'By now, lIDX should point to a valid item
        oItems(lIdx).Quantity += blQuantity

        If oItems(lIdx).Quantity < 1 Then lItemIdx(lIdx) = -1
    End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try

            sSQL = "UPDATE tblTradePlayer SET Notes = '"
            If Notes Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(Notes))
            sSQL &= "', TradePostID = " & lTradePostID & " WHERE TradeID = " & TradeID & " AND PlayerID = " & PlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                oComm = Nothing

                sSQL = "INSERT INTO tblTradePlayer (TradeID, PlayerID, Notes, TradePostID) VALUES (" & TradeID & ", " & PlayerID & ", '"
                If Notes Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(Notes))
                sSQL &= "', " & lTradePostID & ")"

                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
            End If
            oComm = Nothing

            'Ok, because there is no real key on the tradeplayeritem table...
            sSQL = "DELETE FROM tblTradePlayerItem WHERE TradeID = " & Me.TradeID & " AND PlayerID = " & Me.PlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            For X As Int32 = 0 To lItemUB
                If lItemIdx(X) <> -1 Then
                    oItems(X).PlayerID = Me.PlayerID
                    oItems(X).TradeID = Me.TradeID
                    oItems(X).SaveObject()
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object TradePlayer. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Private moPlayer As Player = Nothing
    Public ReadOnly Property oPlayer() As Player
        Get
            If moPlayer Is Nothing Then moPlayer = GetEpicaPlayer(PlayerID)
            Return moPlayer
        End Get
    End Property

    Public Function VerifyDetails(ByVal bPlayer1 As Boolean, ByRef oOtherTradePost As Facility) As Byte
        'Ok, confirm that the tradepost exists
        If TradePost Is Nothing OrElse TradePost.ParentColony Is Nothing Then
            If bPlayer1 = True Then
                Return DirectTrade.eFailureReason.Player1SourceNotFound
            Else : Return DirectTrade.eFailureReason.Player2SourceNotFound
            End If
        End If

        'Now, get the parent colony
        Dim oColony As Colony = TradePost.ParentColony
        If oColony Is Nothing Then
            If bPlayer1 = True Then Return DirectTrade.eFailureReason.Player1SourceNotFound Else Return DirectTrade.eFailureReason.Player2SourceNotFound
        End If

        Dim bNoGood As Boolean

        Dim lAgentCnt As Int32 = 0

        'Now, verify source has the items being offered
        For lIdx As Int32 = 0 To lItemUB
            If lItemIdx(lIdx) = -1 Then Continue For

            bNoGood = False

            Dim blQty As Int64 = oItems(lIdx).Quantity
            Dim lID As Int32 = oItems(lIdx).ObjectID
            Dim iTypeID As Int16 = oItems(lIdx).ObjTypeID

            Select Case Math.Abs(iTypeID)
                '================== THESE ITEMS REQUIRE THE PARENT COLONY ===============
                Case ObjectType.eColonists
                    bNoGood = oColony.Population < blQty
                Case ObjectType.eEnlisted
                    bNoGood = oColony.ColonyEnlisted < blQty
                Case ObjectType.eOfficers
                    bNoGood = oColony.ColonyOfficers < blQty
                Case ObjectType.eCredits
                    bNoGood = oPlayer.blCredits < blQty
                Case ObjectType.eAgent
                    Dim oAgent As Agent = GetEpicaAgent(lID)
                    If oAgent Is Nothing Then
                        bNoGood = True
                    Else
                        If (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.OnAMission Or AgentStatus.IsInfiltrated Or AgentStatus.Infiltrating Or AgentStatus.CounterAgent Or AgentStatus.ReturningHome)) <> 0 Then bNoGood = True
                    End If

                    If bNoGood = False Then lAgentCnt += 1

                Case ObjectType.eFacility
                    'Source and Dest must be within the same environment for this to work
                    Dim lPID As Int32
                    Dim lPTID As Int16
                    With CType(TradePost.ParentObject, Epica_GUID)
                        lPID = .ObjectID
                        lPTID = .ObjTypeID
                    End With
                    With CType(oOtherTradePost.ParentObject, Epica_GUID)
                        If .ObjectID <> lPID OrElse .ObjTypeID <> lPTID Then
                            Return DirectTrade.eFailureReason.FacilityTradeDifferentEnvirs
                        End If
                    End With

                    'in the colony
                    bNoGood = True
                    For X As Int32 = 0 To oColony.ChildrenUB
                        If oColony.lChildrenIdx(X) = lID Then
                            bNoGood = False
                            Exit For
                        End If
                    Next X

                    '================  THESE ITEMS COME ONLY FROM THE TRADEPOST ==================
                Case ObjectType.eColony
                    bNoGood = oColony Is Nothing
                    'TODO: Figure this out, for now make nogood true
                    bNoGood = True
                Case ObjectType.eFood
                    'TODO: Figure this out
                    bNoGood = True
                Case ObjectType.eStock
                    'TODO: Figure this out
                    bNoGood = True
                Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
                    'Ok, determine if we are giving a component cache or a component design
                    If iTypeID < 0 Then
                        'Component Cache
                        bNoGood = True
                        iTypeID = Math.Abs(iTypeID)
                        With TradePost.ParentColony
                            'For lChild As Int32 = 0 To .ChildrenUB
                            '    If .lChildrenIdx(lChild) <> -1 Then
                            '        Dim oFac As Facility = .oChildren(lChild)
                            '        If oFac Is Nothing = False AndAlso oFac.Active = True AndAlso (oFac.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                            '            For lCIdx As Int32 = 0 To oFac.lCargoUB
                            '                If oFac.lCargoIdx(lCIdx) <> -1 AndAlso oFac.oCargoContents(lCIdx).ObjTypeID = ObjectType.eComponentCache Then
                            '                    With CType(oFac.oCargoContents(lCIdx), ComponentCache)
                            '                        If .ObjectID = lID AndAlso .ComponentTypeID = iTypeID Then
                            '                            blQty -= .Quantity
                            '                        End If
                            '                    End With
                            '                End If
                            '                If blQty < 1 Then Exit For
                            '            Next lCIdx
                            '            If blQty < 1 Then Exit For
                            '        End If
                            '    End If
                            'Next lChild 
                            'Next lChild
                            For X As Int32 = 0 To .mlComponentCacheUB
                                If .mlComponentCacheIdx(X) > -1 Then
                                    If .mlComponentCacheID(X) = glComponentCacheIdx(.mlComponentCacheIdx(X)) Then
                                        Dim oCache As ComponentCache = goComponentCache(.mlComponentCacheIdx(X))
                                        If oCache Is Nothing = False Then
                                            If oCache.ObjectID = lID AndAlso oCache.ComponentTypeID = Math.Abs(iTypeID) Then
                                                blQty -= oCache.Quantity
                                            End If
                                        End If
                                    End If
                                End If
                                If blQty < 1 Then Exit For
                            Next X
                        End With
                        If blQty < 1 Then
                            bNoGood = False
                        End If
                    Else
                        'Component design
                        If iTypeID = ObjectType.eSpecialTech Then bNoGood = True
                        'TODO: Ensure they can give this technology design... for now always say no good
                        bNoGood = True
                    End If
                Case ObjectType.eUnit
                    'in a hangar
                    bNoGood = True
                    Dim oUnit As Unit = GetEpicaUnit(lID)
                    If oUnit Is Nothing = False Then
                        If CType(oUnit.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                            If CType(oUnit.ParentObject, Facility).ParentColony Is Nothing = False AndAlso CType(oUnit.ParentObject, Facility).ParentColony.ObjectID = oColony.ObjectID Then
                                bNoGood = False
                            End If
                        End If
                        'If oOtherTradePost Is Nothing = False AndAlso oOtherTradePost.Owner Is Nothing = False AndAlso oOtherTradePost.Owner.blWarpoints < oUnit.EntityDef.WarpointUpkeep Then
                        '    bNoGood = True
                        'End If
                    End If
                Case ObjectType.ePlayerIntel
                    Dim oPI As PlayerIntel = Me.oPlayer.GetOrAddPlayerIntel(lID, True)
                    bNoGood = oPI Is Nothing
                Case ObjectType.ePlayerItemIntel
                    bNoGood = True
                    With oPlayer
                        For X As Int32 = 0 To .mlItemIntelUB
                            If .moItemIntel(X) Is Nothing = False Then
                                If .moItemIntel(X).lItemID = lID AndAlso .moItemIntel(X).iItemTypeID = oItems(lIdx).lExtendedID Then
                                    If .moItemIntel(X).yIntelType > PlayerItemIntel.PlayerItemIntelType.eExistance Then
                                        bNoGood = False
                                    End If
                                    Exit For
                                End If
                            End If
                        Next X
                    End With
                Case ObjectType.ePlayerTechKnowledge
                    Dim oPTK As PlayerTechKnowledge = oPlayer.GetPlayerTechKnowledge(lID, CShort(oItems(lIdx).lExtendedID))
                    bNoGood = oPTK Is Nothing
                Case Else
                    'in a cargo bay
                    bNoGood = True

                    Dim oMinCache As MineralCache = GetEpicaMineralCache(lID)
                    If oMinCache Is Nothing = False Then

                        If oMinCache.ParentObject Is Nothing = False Then
                            Dim iTemp As Int16 = CType(oMinCache.ParentObject, Epica_GUID).ObjTypeID
                            If iTemp = ObjectType.eFacility Then
                                Dim oFac As Facility = CType(oMinCache.ParentObject, Facility)
                                If oFac.ParentColony Is Nothing = False Then
                                    If oFac.ParentColony.ObjectID = TradePost.ParentColony.ObjectID Then
                                        blQty -= oMinCache.Quantity
                                    End If
                                End If
                            ElseIf iTemp = ObjectType.eColony Then
                                Dim oTestColony As Colony = CType(oMinCache.ParentObject, Colony)
                                If oTestColony Is Nothing = False AndAlso oTestColony.ObjectID = oColony.ObjectID Then
                                    blQty -= oMinCache.Quantity
                                End If
                            End If
                        End If


                    End If
                    ''blqty -= tradepost.ParentColony.ColonyMineralQuantity(
                    'With TradePost.ParentColony
                    '    For lCIdx As Int32 = 0 To .lCargoUB
                    '        If .lCargoIdx(lCIdx) <> -1 AndAlso .oCargoContents(lCIdx).ObjTypeID = iTypeID Then
                    '            'Ok, now... 
                    '            If iTypeID = ObjectType.eMineralCache Then
                    '                If CType(.oCargoContents(lCIdx), MineralCache).ObjectID = lID Then
                    '                    blQty -= CType(.oCargoContents(lCIdx), MineralCache).Quantity
                    '                End If
                    '            End If
                    '        End If

                    '        If blQty < 1 Then Exit For
                    '    Next lCIdx
                    'End With

                    If blQty < 1 Then
                        bNoGood = False
                    End If

            End Select

            If bNoGood = True Then
                If bPlayer1 = True Then
                    Return DirectTrade.eFailureReason.Player1MissingItems
                Else : Return DirectTrade.eFailureReason.Player2MissingItems
                End If
            End If
        Next lIdx

        If lAgentCnt > 0 Then
            For X As Int32 = 0 To oOtherTradePost.Owner.mlAgentUB
                If oOtherTradePost.Owner.mlAgentIdx(X) <> -1 Then lAgentCnt += 1
            Next X
            If lAgentCnt > 50 Then
                If bPlayer1 = True Then
                    Return DirectTrade.eFailureReason.Player2AgentsFull
                Else : Return DirectTrade.eFailureReason.Player1AgentsFull
                End If
            End If
        End If


        'if we are here, all is good
        Return DirectTrade.eFailureReason.NoFailureReason

    End Function

    Public Function ExecuteTrade(ByRef oTrade As DirectTrade) As Boolean
        'Confirm the Source is found
        If TradePost Is Nothing OrElse TradePost.ParentColony Is Nothing Then Return False

        Dim oSource As Colony = TradePost.ParentColony
        If oSource Is Nothing Then Return False

        For lIdx As Int32 = 0 To lItemUB
            If lItemIdx(lIdx) <> -1 Then oItems(lIdx).blConsumed = 0
        Next lIdx

        'Now, verify source has the items being offered
        For lIdx As Int32 = 0 To lItemUB
            If lItemIdx(lIdx) = -1 Then Continue For

            Dim blQty As Int64 = oItems(lIdx).Quantity
            Dim lID As Int32 = oItems(lIdx).ObjectID
            Dim iTypeID As Int16 = oItems(lIdx).ObjTypeID

            Select Case Math.Abs(iTypeID)
                '================== THESE ITEMS REQUIRE THE PARENT COLONY ===============
                Case ObjectType.eColonists
                    If oSource.Population < blQty Then Return False
                    oSource.Population -= CInt(blQty)
                Case ObjectType.eEnlisted
                    If oSource.ColonyEnlisted < blQty Then Return False
                    oSource.ColonyEnlisted -= CInt(blQty)
                Case ObjectType.eOfficers
                    If oSource.ColonyOfficers < blQty Then Return False
                    oSource.ColonyOfficers -= CInt(blQty)
                Case ObjectType.eCredits
                    If oPlayer.blCredits < blQty Then Return False
                    oPlayer.blCredits -= blQty
                    oItems(lIdx).blConsumed = blQty
                    Continue For
                Case ObjectType.eAgent
                    'confirm the agent is still available
                    Dim bFound As Boolean = False
                    For X As Int32 = 0 To oPlayer.mlAgentUB
                        If oPlayer.mlAgentID(X) = lID AndAlso glAgentIdx(oPlayer.mlAgentIdx(X)) = lID Then
                            Dim oAgent As Agent = goAgent(oPlayer.mlAgentIdx(X))
                            If oAgent Is Nothing = False Then
                                If (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.OnAMission Or AgentStatus.IsInfiltrated Or AgentStatus.Infiltrating Or AgentStatus.CounterAgent Or AgentStatus.ReturningHome)) = 0 Then
                                    bFound = True
                                End If
                            End If
                            Exit For
                        End If
                    Next X
                    If bFound = False Then Return False
                Case ObjectType.eFacility
                    'in the colony
                    'this occurs at the end of the trade


                    '================  THESE ITEMS COME ONLY FROM THE TRADEPOST ==================
                Case ObjectType.eColony
                    'This occurs at the end of the trade
                Case ObjectType.eFood
                    'TODO: Figure this out
                Case ObjectType.eStock
                    'TODO: Figure this out
                Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
                    If iTypeID < 0 Then
                        'Giving a component cache
                        Dim bNoGood As Boolean = True
                        Dim lTemp As Int32
                        iTypeID = Math.Abs(iTypeID)

                        With TradePost.ParentColony
                            'For lChild As Int32 = 0 To .ChildrenUB
                            '    If .lChildrenIdx(lChild) <> -1 Then
                            '        Dim oFac As Facility = .oChildren(lChild)
                            '        If oFac Is Nothing = False AndAlso oFac.Active = True AndAlso (oFac.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
                            '            For lCIdx As Int32 = 0 To oFac.lCargoUB
                            '                If oFac.lCargoIdx(lCIdx) <> -1 AndAlso oFac.oCargoContents(lCIdx).ObjTypeID = ObjectType.eComponentCache Then
                            '                    With CType(oFac.oCargoContents(lCIdx), ComponentCache)
                            '                        If .ObjectID = lID AndAlso .ComponentTypeID = iTypeID Then
                            '                            lTemp = CInt(Math.Min(blQty, .Quantity))
                            '                            .Quantity -= lTemp
                            '                            blQty -= lTemp

                            '                            oItems(lIdx).lExtendedID = .ComponentOwnerID
                            '                            oItems(lIdx).lExtended2ID = .ComponentID
                            '                        End If
                            '                    End With
                            '                End If
                            '                If blQty < 1 Then Exit For
                            '            Next lCIdx
                            '            If blQty < 1 Then Exit For
                            '        End If
                            '    End If
                            'Next lChild
                            For X As Int32 = 0 To .mlComponentCacheUB
                                If .mlComponentCacheIdx(X) > -1 Then
                                    If .mlComponentCacheID(X) = glComponentCacheIdx(.mlComponentCacheIdx(X)) Then
                                        Dim oCache As ComponentCache = goComponentCache(.mlComponentCacheIdx(X))
                                        If oCache Is Nothing = False Then
                                            If oCache.ObjectID = lID AndAlso oCache.ComponentTypeID = Math.Abs(iTypeID) Then
                                                lTemp = CInt(Math.Min(blQty, oCache.Quantity))
                                                oCache.Quantity -= lTemp
                                                blQty -= lTemp

                                                oItems(lIdx).lExtendedID = oCache.ComponentOwnerID
                                                oItems(lIdx).lExtended2ID = oCache.ComponentID
                                            End If
                                        End If
                                    End If
                                End If
                                If blQty < 1 Then Exit For
                            Next X
                        End With

                        If blQty < 1 Then
                            bNoGood = False
                        End If
                    Else
                        'Giving a technology design (no special techs)... the result will occur at the end of the trade
                    End If
                Case ObjectType.eUnit
                    'in a hangar
                    Dim bFound As Boolean = False

                    Dim oUnit As Unit = GetEpicaUnit(lID)
                    If oUnit Is Nothing = False AndAlso oUnit.ParentObject Is Nothing = False Then
                        If CType(oUnit.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                            If CType(oUnit.ParentObject, Facility).ParentColony.ObjectID = TradePost.ParentColony.ObjectID Then
                                bFound = True
                                blQty = 0

                                With CType(oUnit.ParentObject, Facility)
                                    Dim lHIdx As Int32 = -1
                                    For X As Int32 = 0 To .lHangarUB
                                        If .lHangarIdx(X) = lID Then
                                            lHIdx = X
                                            Exit For
                                        End If
                                    Next X
                                    If lHIdx <> -1 Then
                                        Dim lFleetID As Int32 = CType(.oHangarContents(lHIdx), Unit).lFleetID
                                        If lFleetID <> -1 Then
                                            Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lFleetID)
                                            If oUnitGroup Is Nothing = False Then oUnitGroup.RemoveUnit(lID, True, False)
                                        End If

                                        'And remove it from my hangar
                                        .lHangarIdx(lHIdx) = -1
                                        .oHangarContents(lHIdx) = Nothing
                                    End If

                                End With

                            End If
                        End If
                    End If
                    If bFound = False Then Return False

                Case ObjectType.ePlayerIntel
                    Dim oPI As PlayerIntel = Me.oPlayer.GetOrAddPlayerIntel(lID, True)
                    If oPI Is Nothing Then Return False
                Case ObjectType.ePlayerItemIntel
                    With oPlayer
                        Dim bFound As Boolean = False
                        For X As Int32 = 0 To .mlItemIntelUB
                            If .moItemIntel(X) Is Nothing = False Then
                                If .moItemIntel(X).lItemID = lID AndAlso .moItemIntel(X).iItemTypeID = oItems(lIdx).lExtendedID Then
                                    If .moItemIntel(X).yIntelType > PlayerItemIntel.PlayerItemIntelType.eExistance Then
                                        bFound = True
                                    End If
                                End If
                            End If
                        Next X
                        If bFound = False Then Return False
                    End With
                Case ObjectType.ePlayerTechKnowledge
                    If oPlayer.HasTechKnowledge(lID, CShort(oItems(lIdx).lExtendedID), PlayerTechKnowledge.KnowledgeType.eSettingsLevel1) = False Then Return False
                Case Else
                    'in a cargo bay
                    Dim bNoGood As Boolean = True
                    'Dim lTemp As Int32

                    'With TradePost
                    '	For lCIdx As Int32 = 0 To .lCargoUB
                    '		If .lCargoIdx(lCIdx) <> -1 AndAlso .oCargoContents(lCIdx).ObjTypeID = iTypeID Then
                    '			Select Case iTypeID
                    '				Case ObjectType.eMineralCache
                    '					If CType(.oCargoContents(lCIdx), MineralCache).ObjectID = lID Then
                    '						lTemp = CInt(Math.Min(blQty, CType(.oCargoContents(lCIdx), MineralCache).Quantity))
                    '						oItems(lIdx).lExtendedID = CType(.oCargoContents(lCIdx), MineralCache).oMineral.ObjectID
                    '						CType(.oCargoContents(lCIdx), MineralCache).Quantity -= lTemp
                    '						blQty -= lTemp
                    '					End If
                    '				Case ObjectType.eAmmunition
                    '					If CType(.oCargoContents(lCIdx), AmmunitionCache).oWeaponTech.ObjectID = lID Then
                    '						lTemp = CInt(Math.Min(blQty, CType(.oCargoContents(lCIdx), AmmunitionCache).Quantity))
                    '						CType(.oCargoContents(lCIdx), AmmunitionCache).Quantity -= lTemp
                    '						blQty -= lTemp
                    '					End If
                    '			End Select
                    '		End If
                    '		If blQty < 1 Then Exit For
                    '	Next lCIdx
                    'End With
                    Dim oCache As MineralCache = GetEpicaMineralCache(lID)
                    If oCache Is Nothing = False Then oCache.Quantity -= CInt(blQty)
                    blQty = 0
                    bNoGood = False
            End Select

            oItems(lIdx).blConsumed = oItems(lIdx).Quantity - blQty
        Next lIdx

        'if we are here, all is good
        Return True
    End Function

    Public Sub FinalizeTrade(ByRef oOtherPlayer As TradePlayer)
        If oOtherPlayer.TradePost Is Nothing OrElse oOtherPlayer.TradePost.ParentColony Is Nothing Then Return
        Dim oColony As Colony = oOtherPlayer.TradePost.ParentColony

        Dim bIncludesColony As Boolean = False

        'Ok, the items for oPlayer1TP need to go to oDest2
        For lIdx As Int32 = 0 To lItemUB
            If lItemIdx(lIdx) <> -1 Then
                Dim iTypeID As Int16 = oItems(lIdx).ObjTypeID
                Dim lID As Int32 = oItems(lIdx).ObjectID
                Dim blQty As Int64 = oItems(lIdx).Quantity

                Select Case Math.Abs(iTypeID)
                    '================== THESE ITEMS REQUIRE THE PARENT COLONY ===============
                    Case ObjectType.eColonists
                        oColony.Population += CInt(blQty)
                    Case ObjectType.eEnlisted
                        oColony.ColonyEnlisted += CInt(blQty)
                    Case ObjectType.eOfficers
                        oColony.ColonyOfficers += CInt(blQty)
                    Case ObjectType.eCredits
                        oOtherPlayer.oPlayer.blCredits += blQty
                    Case ObjectType.eAgent
                        For X As Int32 = 0 To oPlayer.mlAgentUB
                            If oPlayer.mlAgentID(X) = lID AndAlso glAgentIdx(oPlayer.mlAgentIdx(X)) = lID Then
                                Dim oAgent As Agent = goAgent(oPlayer.mlAgentIdx(X))
                                If oAgent Is Nothing = False Then
                                    If (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.OnAMission)) = 0 Then
                                        Dim yResp(9) As Byte
                                        System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yResp, 0)
                                        System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yResp, 2)
                                        System.BitConverter.GetBytes(Int32.MaxValue).CopyTo(yResp, 6)
                                        oPlayer.SendPlayerMessage(yResp, False, AliasingRights.eViewAgents)
                                        oAgent.RemoveMeFromMissions()

                                        oAgent.oOwner = oOtherPlayer.oPlayer
                                        oAgent.Loyalty = CByte(Rnd() * 100)
                                        oAgent.Suspicion = 0
                                        oAgent.lTargetID = -1
                                        oAgent.iTargetTypeID = -1
                                        oAgent.InfiltrationLevel = 0
                                        oAgent.InfiltrationType = eInfiltrationType.eGeneralInfiltration
                                        If (oAgent.lAgentStatus And AgentStatus.NewRecruit) = 0 Then oAgent.lAgentStatus = AgentStatus.NormalStatus
                                        oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                        oAgent.oOwner.AddAgentLookup(oAgent.ObjectID, oAgent.ServerIndex)
                                    End If
                                End If
                                oPlayer.mlAgentIdx(X) = -1
                                oPlayer.mlAgentID(X) = -1
                                Exit For
                            End If
                        Next X

                    Case ObjectType.eColony
                        bIncludesColony = True
                    Case ObjectType.eFood
                        'TODO: Figure this out
                    Case ObjectType.eStock
                        'TODO: Figure this out


                        '================  THESE ITEMS COME ONLY FROM THE TRADEPOST ==================
                    Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
                        If iTypeID < 0 Then
                            'Giving a component cache... lID is the ComponentCACHEID not the componentid
                            ' iTypeID is the Component's type ID, lExtended2ID is the ComponentID
                            If oOtherPlayer.TradePost.AddComponentCacheToCargo(oItems(lIdx).lExtended2ID, Math.Abs(iTypeID), CInt(blQty), oItems(lIdx).lExtendedID) Is Nothing Then
                                'TODO: Some components did not make the destination
                            End If
                        Else
                            'TODO: Figure this out
                        End If
                    Case ObjectType.eUnit
                        'Units are delivered to the target colony's Parent Environment
                        Dim oUnit As Unit = GetEpicaUnit(lID)
                        If oUnit Is Nothing Then
                            'TODO: Unit did not make it
                            Continue For
                        End If

                        'TODO: Need to refund the Givers's Enlisted and Officers and deduct the receiver's Enlisted and Officers

                        'ensure the unit is removed from the previous hangar
                        If oUnit.ParentObject Is Nothing = False AndAlso CType(oUnit.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                            With CType(oUnit.ParentObject, Facility)
                                For X As Int32 = 0 To .lHangarUB
                                    If .lHangarIdx(X) = oUnit.ObjectID Then
                                        .lHangarIdx(X) = -1
                                        .oHangarContents = Nothing
                                        Exit For
                                    End If
                                Next X
                            End With
                        End If

                        Dim bAddHangarRef As Boolean = True
                        Dim iTemp As Int16
                        If (oUnit.EntityDef.yChassisType And oOtherPlayer.TradePost.EntityDef.yChassisType) <> 0 Then
                            oUnit.ParentObject = oOtherPlayer.TradePost
                        Else
                            If (oUnit.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                                iTemp = CType(oOtherPlayer.TradePost.ParentObject, Epica_GUID).ObjTypeID
                                If iTemp = ObjectType.eSolarSystem Then
                                    'parent
                                    oUnit.ParentObject = oOtherPlayer.TradePost.ParentObject
                                    bAddHangarRef = False
                                ElseIf iTemp = ObjectType.ePlanet Then
                                    'parent's parent
                                    With CType(oOtherPlayer.TradePost.ParentObject, Planet)
                                        oUnit.LocX = .LocX
                                        oUnit.LocZ = .LocZ
                                    End With
                                    oUnit.ParentObject = CType(oOtherPlayer.TradePost.ParentObject, Planet).ParentSystem
                                    bAddHangarRef = False
                                Else
                                    'parent's parent (facility)
                                    With CType(oOtherPlayer.TradePost.ParentObject, Facility)
                                        oUnit.LocX = .LocX
                                        oUnit.LocZ = .LocZ + 1000
                                    End With
                                    oUnit.ParentObject = CType(oOtherPlayer.TradePost.ParentObject, Facility).ParentObject
                                    bAddHangarRef = False
                                End If
                            Else : oUnit.ParentObject = oOtherPlayer.TradePost      'TODO: does not handle naval correctly
                            End If
                        End If

                        oUnit.SetHangarContentsOwner(oOtherPlayer.oPlayer)

                        oUnit.Owner.lMilitaryScore -= oUnit.EntityDef.CombatRating
                        oUnit.Owner = oOtherPlayer.oPlayer
                        oUnit.DataChanged()

                        iTemp = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
                        If bAddHangarRef = True Then
                            If iTemp = ObjectType.eFacility OrElse iTemp = ObjectType.eUnit Then
                                CType(oUnit.ParentObject, Epica_Entity).AddHangarRef(CType(oUnit, Epica_GUID))
                            End If
                        Else
                            'ensure we've called datachanged
                            oUnit.DataChanged()
                            Try
                                If iTemp = ObjectType.eSolarSystem Then
                                    CType(oUnit.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand))
                                ElseIf iTemp = ObjectType.ePlanet Then
                                    CType(oUnit.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand))
                                End If
                            Catch ex As Exception
                                LogEvent(LogEventType.CriticalError, "TradePlayer.FinalizeTrade: Unit traded. Not in hangar. Could not add object to domain. UnitID: " & oUnit.ObjectID & ", Parent: " & CType(oUnit.ParentObject, Epica_GUID).ObjectID & ", " & iTemp & ". Reason: " & ex.Message)
                            End Try
                        End If

                    Case ObjectType.eFacility

                        Dim oFac As Facility = GetEpicaFacility(lID)
                        If oFac Is Nothing = False Then
                            Dim oSource As Colony = Nothing
                            With oFac
                                'All contents go with it... this is the case for cargo contents but not for hangar
                                .SetHangarContentsOwner(oColony.Owner)

                                .ClearQueue()
                                .CurrentProduction = Nothing

                                'Ok, set the owner to nothing
                                .Owner = oColony.Owner
                                'Remove it from the colony
                                oSource = .ParentColony
                                .ParentColony = oColony

                                oColony.AddChildFacility(oFac)
                                oColony.UpdateAllValues(-1)
                            End With

                            If oSource Is Nothing = False Then
                                For X As Int32 = 0 To oSource.ChildrenUB
                                    If oSource.lChildrenIdx(X) = lID Then
                                        oSource.lChildrenIdx(X) = -1
                                        oSource.oChildren(X) = Nothing
                                        Exit For
                                    End If
                                Next X
                                oSource.UpdateAllValues(-1)
                            End If

                            oFac.DataChanged()
                            With CType(oFac.ParentObject, Epica_GUID)
                                If .ObjTypeID = ObjectType.ePlanet Then
                                    CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                                ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                                    CType(oFac.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                                End If
                            End With
                        Else
                            'TODO: Facility did not make the trade
                        End If
                    Case ObjectType.ePlayerIntel
                        Dim oPI As PlayerIntel = Me.oPlayer.GetOrAddPlayerIntel(lID, True)
                        If oPI Is Nothing = False Then
                            oPI = oPI.CloneForPlayer(oOtherPlayer.oPlayer)
                            If oPI Is Nothing = False Then
                                oOtherPlayer.oPlayer.SetPlayerIntel(oPI)
                                oOtherPlayer.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPI, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                            End If
                        End If
                    Case ObjectType.ePlayerItemIntel
                        With oPlayer
                            For X As Int32 = 0 To .mlItemIntelUB
                                If .moItemIntel(X) Is Nothing = False Then
                                    If .moItemIntel(X).lItemID = lID AndAlso .moItemIntel(X).iItemTypeID = oItems(lIdx).lExtendedID Then
                                        If .moItemIntel(X).yIntelType > PlayerItemIntel.PlayerItemIntelType.eExistance Then
                                            Dim oPII As PlayerItemIntel = .moItemIntel(X).CloneForPlayer(oOtherPlayer.oPlayer)
                                            If oPII Is Nothing = False Then
                                                oOtherPlayer.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                            End If
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next X
                        End With
                    Case ObjectType.ePlayerTechKnowledge
                        Dim oPTK As PlayerTechKnowledge = oPlayer.GetPlayerTechKnowledge(lID, CShort(oItems(lIdx).lExtendedID))
                        If oPTK Is Nothing = False Then
                            Dim oNewPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oOtherPlayer.oPlayer, oPTK.oTech, oPTK.yKnowledgeType, False)
                            If oNewPTK Is Nothing = False Then
                                'oOtherPlayer.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oNewPTK, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents Or AliasingRights.eViewTechDesigns Or AliasingRights.eViewResearch)
                                oNewPTK.SendMsgToPlayer()
                            End If
                        End If
                    Case Else
                        'Either mineral cache or ammunition
                        If iTypeID = ObjectType.eAmmunition Then
                            'TODO: Uh...
                        Else
                            If oOtherPlayer.TradePost.AddMineralCacheToCargo(oItems(lIdx).lExtendedID, CInt(blQty)) Is Nothing Then
                                'TODO: Some components did not make the destination
                            Else
                                oOtherPlayer.oPlayer.CheckFirstContactWithMineral(oItems(lIdx).lExtendedID)
                            End If
                        End If
                End Select
            End If
        Next lIdx


        If bIncludesColony = True Then
            If oColony Is Nothing = False Then
                'Ok, now, check the dest player
                Dim oNewCol As Colony = Nothing
                With CType(oColony.ParentObject, Epica_GUID)
                    Dim lColIdx As Int32 = oColony.Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
                    If lColIdx <> -1 Then
                        If glColonyIdx(lColIdx) <> -1 Then
                            oNewCol = goColony(lColIdx)
                        End If
                    End If
                End With

                If oNewCol Is Nothing = False Then
                    oNewCol.ConsumeColony(oColony)
                Else
                    oNewCol = oColony
                    oNewCol.Owner = oColony.Owner
                    For X As Int32 = 0 To oNewCol.ChildrenUB
                        If oNewCol.lChildrenIdx(X) <> -1 Then
                            oNewCol.oChildren(X).Owner = oColony.Owner
                            oNewCol.oChildren(X).SetHangarContentsOwner(oColony.Owner)
                        End If
                    Next X
                End If

                'Check if we need to update everyone...
                Dim iTemp As Int16 = CType(oNewCol.ParentObject, Epica_GUID).ObjTypeID
                If iTemp <> ObjectType.eFacility Then
                    Dim oSocket As NetSock = Nothing
                    If iTemp = ObjectType.ePlanet Then
                        oSocket = CType(oNewCol.ParentObject, Planet).oDomain.DomainSocket
                    ElseIf iTemp = ObjectType.eSolarSystem Then
                        oSocket = CType(oNewCol.ParentObject, SolarSystem).oDomain.DomainSocket
                    End If

                    If oSocket Is Nothing = False Then
                        For X As Int32 = 0 To oNewCol.ChildrenUB
                            If oNewCol.lChildrenIdx(X) <> -1 Then
                                oNewCol.oChildren(X).DataChanged()
                                oSocket.SendData(goMsgSys.GetAddObjectMessage(oNewCol.oChildren(X), GlobalMessageCode.eAddObjectCommand))
                            End If
                        Next X
                    End If

                End If

            End If
        End If

    End Sub

    Public Function GetTotalCreditsTransferred() As Int64
        Dim blVal As Int64 = 0
        For X As Int32 = 0 To lItemUB
            If lItemIdx(X) < 0 Then Continue For
            If oItems(X).blConsumed < 1 Then Continue For
            If Math.Abs(oItems(X).ObjTypeID) = ObjectType.eCredits Then
                blVal += oItems(X).Quantity
            End If
        Next X
        Return blVal
    End Function

    Public Sub RefundItems()
        If TradePost Is Nothing OrElse TradePost.ParentColony Is Nothing Then Return
        'Confirm the Source is found
        Dim oSource As Colony = TradePost.ParentColony
        If oSource Is Nothing Then Return

        'Now, verify source has the items being offered
        For lIdx As Int32 = 0 To lItemUB
            If lItemIdx(lIdx) = -1 Then Continue For
            If oItems(lIdx).blConsumed < 1 Then Continue For

            Dim blQty As Int64 = oItems(lIdx).blConsumed
            Dim lID As Int32 = oItems(lIdx).ObjectID
            Dim iTypeID As Int16 = oItems(lIdx).ObjTypeID

            Select Case Math.Abs(iTypeID)
                '================== THESE ITEMS REQUIRE THE PARENT COLONY ===============
                Case ObjectType.eColonists
                    oSource.Population += CInt(blQty)
                Case ObjectType.eEnlisted
                    oSource.ColonyEnlisted += CInt(blQty)
                Case ObjectType.eOfficers
                    oSource.ColonyOfficers += CInt(blQty)
                Case ObjectType.eCredits
                    oPlayer.blCredits += blQty
                Case ObjectType.eAgent
                    Dim oAgent As Agent = GetEpicaAgent(lID)
                    If oAgent Is Nothing = False Then
                        If oAgent.oOwner Is Nothing = False Then
                            If oAgent.oOwner.ObjectID <> Me.oPlayer.ObjectID Then
                                For X As Int32 = 0 To oAgent.oOwner.mlAgentUB
                                    If oAgent.oOwner.mlAgentID(X) = oAgent.ObjectID Then
                                        Dim yResp(9) As Byte
                                        System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yResp, 0)
                                        System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yResp, 2)
                                        System.BitConverter.GetBytes(Int32.MaxValue).CopyTo(yResp, 6)
                                        oAgent.oOwner.SendPlayerMessage(yResp, False, AliasingRights.eViewAgents)

                                        oAgent.oOwner.mlAgentIdx(X) = -1
                                        oAgent.oOwner.mlAgentID(X) = -1
                                        Exit For
                                    End If
                                Next X
                            End If
                        End If
                        oAgent.oOwner = Me.oPlayer
                        Me.oPlayer.AddAgentLookup(oAgent.ObjectID, oAgent.ServerIndex)
                    End If


                    '================  THESE ITEMS COME ONLY FROM THE TRADEPOST ==================
                Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
                    If iTypeID < 0 Then
                        'Giving a component cache
                        'oSource.AddObjectCaches(Math.Abs(lID), Math.Abs(iTypeID), CInt(blQty))
                        TradePost.AddComponentCacheToCargo(Math.Abs(lID), Math.Abs(iTypeID), CInt(blQty), oItems(lIdx).lExtendedID)
                    End If
                Case ObjectType.eUnit
                    'in a hangar
                    Dim oUnit As Unit = GetEpicaUnit(lID)
                    If oUnit Is Nothing = False Then
                        TradePost.AddHangarRef(CType(oUnit, Epica_Entity))
                    End If
                Case ObjectType.eFacility
                    'in the colony
                    'this occurs at the end of the trade, so nothing to do here
                Case Else
                    'in a cargo bay
                    If iTypeID = ObjectType.eAmmunition Then
                        'TODO: Uh....
                    Else
                        TradePost.AddMineralCacheToCargo(lID, CInt(blQty))
                    End If
            End Select

            oItems(lIdx).blConsumed = 0
        Next lIdx
    End Sub

End Class

Public Structure TradePlayerItem
    'Public TP_ID As Int32

    'PK
    Public TradeID As Int32
    Public PlayerID As Int32
    Public ObjectID As Int32
    Public ObjTypeID As Int16
    'END OF PK

    Public Quantity As Int64

    Public lExtendedID As Int32
    Public lExtended2ID As Int32

    'For refunding...
    Public blConsumed As Int64

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing

        Try
            sSQL = "UPDATE tblTradePlayerItem SET Quantity = " & Quantity & ", ExtendedID = " & lExtendedID & ", Extended2ID = " & lExtended2ID & _
              " WHERE TradeID = " & TradeID & " AND PlayerID = " & PlayerID & " AND ObjectID = " & ObjectID & " AND ObjTypeID = " & ObjTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                'Ok, try an insert
                sSQL = "INSERT INTO tblTradePlayerItem (TradeID, PlayerID, ObjectID, ObjTypeID, Quantity, ExtendedID, Extended2ID) VALUES (" & TradeID & _
                  ", " & PlayerID & ", " & ObjectID & ", " & ObjTypeID & ", " & Quantity & ", " & lExtendedID & ", " & lExtended2ID & ")"
                oComm.Dispose()
                oComm = Nothing
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
            End If
            bResult = True
        Catch
            bResult = False
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save TradePlayerItem(" & TradeID & ", " & PlayerID & ", " & ObjectID & ", " & ObjTypeID & "). Reason: " & Err.Description)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

        Return bResult
    End Function
End Structure
