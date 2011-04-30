Option Strict On

'Manages a TradeDelivery - When the GTC makes a delivery run
Public Class TradeDelivery
    Public lID1 As Int32                'typically the object ID
    Public iID2 As Int16                'typically the object type id
    Public lID3 As Int32                'for example, OwnerID of a componentcache's component
    Public blQty As Int64               'the quantity being purchased
    Public lTradePostID As Int32        'the tradepost receiving the items
    Public lSourceTradePostID As Int32  'the tradepost sending the items

    Public dtDelivery As Date           'real date/time that the delivery will occur
    Public dtStartedOn As Date          'real date/time that the delivery started (left source)

    Public lOriginalDestEnvirID As Int32 = -1
    Public iOriginalDestEnvirTypeID As Int16 = -1
    Public lOriginalDestPlayerID As Int32 = -1

    Private moTarget As Facility = Nothing
    Public ReadOnly Property oTargetTradePost() As Facility
        Get
            If moTarget Is Nothing Then moTarget = GetEpicaFacility(lTradePostID)
            If moTarget Is Nothing = False Then
                Dim oParent As Epica_GUID = CType(moTarget.ParentObject, Epica_GUID)
                If oParent Is Nothing = False Then
                    lOriginalDestEnvirID = oParent.ObjectID
                    iOriginalDestEnvirTypeID = oParent.ObjTypeID
                    If moTarget.Owner Is Nothing = False Then lOriginalDestPlayerID = moTarget.Owner.ObjectID
                End If
            Else
                'ok, check for a facility here
                If lOriginalDestPlayerID > -1 Then
                    Dim oPlayer As Player = GetEpicaPlayer(lOriginalDestPlayerID)
                    If oPlayer Is Nothing = False Then
                        Dim lIdx As Int32 = oPlayer.GetColonyFromParent(lOriginalDestEnvirID, iOriginalDestEnvirTypeID)
                        If lIdx > -1 Then
                            Dim oColony As Colony = goColony(lIdx)
                            If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = lOriginalDestPlayerID Then
                                For X As Int32 = 0 To oColony.ChildrenUB
                                    If oColony.lChildrenIdx(X) > -1 Then
                                        Dim oFac As Facility = oColony.oChildren(X)
                                        If oFac Is Nothing = False Then
                                            If oFac.yProductionType = ProductionType.eTradePost Then moTarget = oFac
                                            Return moTarget
                                        End If
                                    End If
                                Next X
                            End If
                        End If
                    End If
                End If
            End If
            Return moTarget
        End Get
    End Property
    Private moSource As Facility = Nothing
    Public ReadOnly Property oSourceTradePost() As Facility
        Get
            If moSource Is Nothing Then moSource = GetEpicaFacility(lSourceTradePostID)
            Return moSource
        End Get
    End Property

    Public Sub ClearTargetTradepost()
        moTarget = Nothing
    End Sub

    Public Function HandleItemDelivery() As Boolean
        Try
            If oTargetTradePost Is Nothing Then
                'ok, find a facility in the environment belonging to the player
                If lOriginalDestEnvirID > -1 AndAlso iOriginalDestEnvirTypeID > -1 AndAlso lOriginalDestPlayerID > -1 Then
                    Dim oPlayer As Player = GetEpicaPlayer(lOriginalDestPlayerID)
                    If oPlayer Is Nothing = False Then
                        Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lOriginalDestEnvirID, iOriginalDestEnvirTypeID)
                        If lColonyIdx > -1 Then
                            Dim oColony As Colony = goColony(lColonyIdx)
                            If oColony Is Nothing = False Then
                                For X As Int32 = 0 To oColony.ChildrenUB
                                    If oColony.lChildrenIdx(X) > -1 Then
                                        Dim oFac As Facility = oColony.oChildren(X)
                                        If oFac Is Nothing = False AndAlso oFac.yProductionType = ProductionType.eTradePost Then
                                            lTradePostID = oFac.ObjectID
                                            Exit For
                                        End If
                                    End If
                                Next X
                            End If
                        End If
                    End If
                End If
                If oTargetTradePost Is Nothing Then Return False
            End If
        Catch
            Return False
        End Try

        'Alrighty, source is assumed to have already been taken care of... we have standardized the way this works...
        ' So.... check iID2
        Select Case Math.Abs(iID2)
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
                If iID2 < 0 Then
                    'Giving a component cache
                    Dim lRemainder As Int32 = CInt(blQty)
                    If oTargetTradePost.ParentColony Is Nothing = False Then
                        lRemainder = oTargetTradePost.ParentColony.AddObjectCaches(lID1, Math.Abs(iID2), lRemainder, True)
                    End If
                    If lRemainder <> 0 Then
                        If oTargetTradePost.AddComponentCacheToCargo(lID1, Math.Abs(iID2), lRemainder, lID3) Is Nothing = False Then Return True
                    Else : Return True
                    End If
                Else
                    'Giving a component design (should not be possible yet)
                    Dim oPlayer As Player = GetEpicaPlayer(lID3)
                    If oPlayer Is Nothing = False Then
                        Dim oTech As Epica_Tech = oPlayer.GetTech(lID1, iID2)
                        If oTech Is Nothing = False Then
                            oTech.Owner = oTargetTradePost.Owner
                            oTargetTradePost.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewTrades)
                            oTargetTradePost.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech.GetResearchCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewTrades)
                            oTargetTradePost.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech.GetProductionCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewTrades)
                            Return True
                        End If
                    End If
                End If
            Case ObjectType.eMineralCache, ObjectType.eMineral
                'Giving a mineral cache
                Dim lRemainder As Int32 = CInt(blQty)
                If oTargetTradePost.Owner Is Nothing = False Then oTargetTradePost.Owner.CheckFirstContactWithMineral(lID1)
                If oTargetTradePost.ParentColony Is Nothing = False Then
                    lRemainder = oTargetTradePost.ParentColony.AddObjectCaches(lID1, ObjectType.eMineralCache, lRemainder, True)
                End If
                If lRemainder > 0 Then
                    If oTargetTradePost.AddMineralCacheToCargo(lID1, lRemainder) Is Nothing = False Then Return True
                Else : Return True
                End If
            Case ObjectType.eUnit
                Dim oUnit As Unit = GetEpicaUnit(lID1)
                If oUnit Is Nothing = False AndAlso oTargetTradePost Is Nothing = False Then
                    'Ok, first, let's verify that the unit can be placed inside the tradepost...
                    Dim oPObj As Object = Nothing   'used if the tradepost has no room
                    Dim bForceNewParent As Boolean = False  'if the tradepost cannot accept the entity, this is true at which point, we use the PObj...
                    Dim bStation As Boolean = False
                    If CType(oTargetTradePost.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                        If (oUnit.EntityDef.yChassisType And (ChassisType.eAtmospheric Or ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                            oPObj = oTargetTradePost.ParentObject
                        Else
                            oPObj = CType(oTargetTradePost.ParentObject, Planet).ParentSystem
                            bForceNewParent = True
                        End If
                    Else
                        bStation = True
                    End If
                    If bForceNewParent = False Then
                        If (oTargetTradePost.Hangar_Cap < oUnit.EntityDef.HullSize OrElse oTargetTradePost.EntityDef.HasHangarDoorSize(oUnit.EntityDef.HullSize) = False) AndAlso bStation = False Then
                            bForceNewParent = True
                        End If
                    End If
                    If bForceNewParent = True Then
                        oUnit.ParentObject = oPObj
                    Else
                        oUnit.ParentObject = oTargetTradePost
                        oTargetTradePost.AddHangarRef(CType(oUnit, Epica_GUID))
                    End If
                    oUnit.LocX = oTargetTradePost.LocX
                    oUnit.LocZ = oTargetTradePost.LocZ
                    oUnit.Owner = oTargetTradePost.Owner
                    oUnit.bUnitInSellOrder = False
                    If CLng(oUnit.Owner.lMilitaryScore) + CLng(oUnit.EntityDef.CombatRating) < Int32.MaxValue Then oUnit.Owner.lMilitaryScore += oUnit.EntityDef.CombatRating
                    oUnit.DataChanged()
                    oUnit.SetHangarContentsOwner(oTargetTradePost.Owner)

                    'oTargetTradePost.Owner.blWarpoints -= oUnit.EntityDef.WarpointUpkeep

                    Dim iTemp As Int16 = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
                    If iTemp = ObjectType.ePlanet Then
                        CType(oUnit.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand))
                    ElseIf iTemp = ObjectType.eSolarSystem Then
                        CType(oUnit.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand))
                    End If
                    oUnit.SaveObject()

                    Return True
                End If
            Case ObjectType.ePlayerIntel
                If Me.oSourceTradePost Is Nothing Then Return False
                Dim oPI As PlayerIntel = Me.oSourceTradePost.Owner.GetOrAddPlayerIntel(lID1, True)
                If oPI Is Nothing = False Then
                    oPI = oPI.CloneForPlayer(Me.oTargetTradePost.Owner)
                    If oPI Is Nothing = False Then
                        Me.oTargetTradePost.Owner.SetPlayerIntel(oPI)
                        Me.oTargetTradePost.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPI, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                    End If
                    Return True
                Else : Return False
                End If
            Case ObjectType.ePlayerItemIntel
                If Me.oSourceTradePost Is Nothing Then Return False
                With Me.oSourceTradePost.Owner
                    For X As Int32 = 0 To .mlItemIntelUB
                        If .moItemIntel(X) Is Nothing = False Then
                            If .moItemIntel(X).lItemID = lID1 AndAlso .moItemIntel(X).iItemTypeID = lID3 Then
                                If .moItemIntel(X).yIntelType > PlayerItemIntel.PlayerItemIntelType.eExistance Then
                                    Dim oPII As PlayerItemIntel = .moItemIntel(X).CloneForPlayer(Me.oTargetTradePost.Owner)
                                    If oPII Is Nothing = False Then
                                        Me.oTargetTradePost.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                        Return True
                                    End If
                                End If
                                Exit For
                            End If
                        End If
                    Next X
                End With
            Case ObjectType.ePlayerTechKnowledge
                If oSourceTradePost Is Nothing Then Return False
                Dim oPTK As PlayerTechKnowledge = Me.oSourceTradePost.Owner.GetPlayerTechKnowledge(lID1, CShort(lID3))
                If oPTK Is Nothing = False Then
                    Dim oNewPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(Me.oTargetTradePost.Owner, oPTK.oTech, oPTK.yKnowledgeType, False)
                    If oNewPTK Is Nothing = False Then
                        'Me.oTargetTradePost.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oNewPTK, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents Or AliasingRights.eViewTechDesigns Or AliasingRights.eViewResearch)
                        oNewPTK.SendMsgToPlayer()
                        Return True
                    End If
                End If
            Case ObjectType.eAgent
                Dim oAgent As Agent = GetEpicaAgent(lID1)
                If oAgent Is Nothing Then Return False

                Dim oPlayer As Player = oAgent.oOwner
                If oPlayer Is Nothing = False Then
                    For X As Int32 = 0 To oPlayer.mlAgentUB
                        If oPlayer.mlAgentID(X) = lID1 AndAlso glAgentIdx(oPlayer.mlAgentIdx(X)) = lID1 Then

                            If oAgent Is Nothing = False Then
                                If (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.OnAMission)) = 0 Then
                                    Dim yResp(9) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yResp, 0)
                                    System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yResp, 2)
                                    System.BitConverter.GetBytes(Int32.MaxValue).CopyTo(yResp, 6)
                                    oAgent.oOwner.SendPlayerMessage(yResp, False, AliasingRights.eViewAgents)
                                    oAgent.RemoveMeFromMissions()

                                    oAgent.oOwner = oTargetTradePost.Owner
                                    oAgent.Loyalty = CByte(Rnd() * 100)
                                    oAgent.Suspicion = 0
                                    oAgent.lTargetID = -1
                                    oAgent.iTargetTypeID = -1
                                    oAgent.InfiltrationLevel = 0
                                    oAgent.InfiltrationType = eInfiltrationType.eGeneralInfiltration
                                    If (oAgent.lAgentStatus And AgentStatus.NewRecruit) = 0 Then
                                        oAgent.dtRecruited = Now
                                        oAgent.lAgentStatus = AgentStatus.NormalStatus
                                    End If
                                    oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                    oAgent.oOwner.AddAgentLookup(oAgent.ObjectID, oAgent.ServerIndex)
                                End If
                            End If
                            oPlayer.mlAgentIdx(X) = -1
                            oPlayer.mlAgentID(X) = -1
                            Return True
                            Exit For
                        End If
                    Next X
                End If


        End Select
        Return False
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            Dim lDelivery As Int32 = GetDateAsNumber(dtDelivery)
            Dim lStartedOn As Int32 = GetDateAsNumber(dtStartedOn)

            sSQL = "DELETE FROM tblTradeDelivery WHERE ID1 = " & lID1 & " AND ID2 = " & iID2 & " AND ID3 = " & lID3 & _
              " AND Quantity = " & blQty & " AND ToTradePostID = " & lTradePostID & " AND FromTradePostID = " & _
              lSourceTradePostID & " AND StartedOn = " & lStartedOn
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            sSQL = "INSERT INTO tblTradeDelivery (ID1, ID2, ID3, Quantity, ToTradePostID, FromTradePostID, DeliveryOn, StartedOn) VALUES (" & _
              lID1 & ", " & iID2 & ", " & lID3 & ", " & blQty & ", " & lTradePostID & ", " & lSourceTradePostID & ", " & lDelivery & ", " & lStartedOn & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object TradeDelivery. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub DeleteMe() 
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            Dim lDelivery As Int32 = GetDateAsNumber(dtDelivery)
            Dim lStartedOn As Int32 = GetDateAsNumber(dtStartedOn)

            sSQL = "DELETE FROM tblTradeDelivery WHERE ID1 = " & lID1 & " AND ID2 = " & iID2 & " AND ID3 = " & lID3 & _
              " AND Quantity = " & blQty & " AND ToTradePostID = " & lTradePostID & " AND FromTradePostID = " & _
              lSourceTradePostID & " AND StartedOn = " & lStartedOn
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to delete object TradeDelivery. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try 
    End Sub

    Public Const ml_TRADE_DELIVERY_ITEM_MSG_LEN As Int32 = 38
    Public Function GetTradeDeliveryPackageItem() As Byte()
        Dim yMsg(ml_TRADE_DELIVERY_ITEM_MSG_LEN - 1) As Byte
        Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, lPos) : lPos += 4
		If oSourceTradePost Is Nothing = False Then
			System.BitConverter.GetBytes(oSourceTradePost.Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
		Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
		End If


        Dim lSeconds As Int32 = CInt(dtDelivery.Subtract(Now).TotalSeconds)
		System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4
		lSeconds = CInt(Now.Subtract(dtStartedOn).TotalSeconds)
		System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oTargetTradePost.ParentColony.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lID1).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(iID2).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lID3).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(blQty).CopyTo(yMsg, lPos) : lPos += 8
        Return yMsg
    End Function
End Class
