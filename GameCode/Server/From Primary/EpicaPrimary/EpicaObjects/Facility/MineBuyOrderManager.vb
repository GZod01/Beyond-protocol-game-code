Option Strict On

Public Class MineBuyOrderManager

    Private Structure Bid
        Public BidAmount As Int32
        Public MaxQuantity As Int32
        Public lPlayerID As Int32
        Public oPlayer As Player
        Public lColonyIdx As Int32
        Public lColonyID As Int32

        Public lPreviousQuantityReceived As Int32
    End Structure

    Public oParentFacility As Facility
    Public oMineralCache As MineralCache

    Public lMaxDaysSold As Int32 = 0
    Public lCurrentConseqDays As Int32 = 0
    Public lDaysNotSold As Int32 = 0
    Public bSomethingSold As Boolean = False
    Public lMinBidAmt As Int32

    Public lPreviousProductionRate As Int32 = 0

    Private muBids(-1) As Bid
    Private mlBidUB As Int32 = -1

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            sSQL = "DELETE FROM tblMineBuyOrder WHERE StructureID = " & oParentFacility.ObjectID & " AND CacheID = " & oMineralCache.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()

            sSQL = "DELETE FROM tblMineBuyOrderBid WHERE StructureID = " & oParentFacility.ObjectID & " AND CacheID = " & oMineralCache.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()

            sSQL = "INSERT INTO tblMineBuyOrder (StructureID, CacheID, MaxDaysSold, CurrentConseqDays, DaysNotSold, SomethingSold) VALUES ("
            sSQL &= oParentFacility.ObjectID & ", " & oMineralCache.ObjectID & ", " & Me.lMaxDaysSold & ", " & Me.lCurrentConseqDays & ", " & Me.lDaysNotSold & ", "
            If bSomethingSold = True Then sSQL &= "1)" Else sSQL &= "0)"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() < 1 Then
                Throw New Exception("No Records affected when inserting into tblMineBuyOrder! ParentFac: " & oParentFacility.ObjectID & ", CacheID = " & oMineralCache.ObjectID)
            End If
            oComm.Dispose()

            Dim lOrderNum As Int32 = 1
            For X As Int32 = 0 To mlBidUB
                With muBids(X)
                    sSQL = "INSERT INTO tblMineBuyOrderBid (StructureID, CacheID, OrderNum, BidAmount, MaxQuantity, PlayerID) VALUES (" & _
                        oParentFacility.ObjectID & ", " & oMineralCache.ObjectID & ", " & lOrderNum & ", " & .BidAmount & ", " & _
                        .MaxQuantity & ", " & .lPlayerID & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() < 1 Then
                        LogEvent(LogEventType.CriticalError, "error saving tblMineBuyOrderBid: No Records affected with insert!")
                    Else : lOrderNum += 1
                    End If
                    oComm.Dispose()
                End With
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save MineBuyOrder. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function GetSaveObjectText() As String
        Dim sSQL As String
        Dim oSB As New System.Text.StringBuilder()

        Try
            'sSQL = "DELETE FROM tblMineBuyOrder WHERE StructureID = " & oParentFacility.ObjectID & " AND CacheID = " & oMineralCache.ObjectID
            'oSB.AppendLine(sSQL)

            'sSQL = "DELETE FROM tblMineBuyOrderBid WHERE StructureID = " & oParentFacility.ObjectID & " AND CacheID = " & oMineralCache.ObjectID
            'oSB.AppendLine(sSQL)

            sSQL = "INSERT INTO tblMineBuyOrder (StructureID, CacheID, MaxDaysSold, CurrentConseqDays, DaysNotSold, SomethingSold) VALUES ("
            sSQL &= oParentFacility.ObjectID & ", " & oMineralCache.ObjectID & ", " & Me.lMaxDaysSold & ", " & Me.lCurrentConseqDays & ", " & Me.lDaysNotSold & ", "
            If bSomethingSold = True Then sSQL &= "1)" Else sSQL &= "0)"
            oSB.AppendLine(sSQL)

            Dim lOrderNum As Int32 = 1
            For X As Int32 = 0 To mlBidUB
                With muBids(X)
                    sSQL = "INSERT INTO tblMineBuyOrderBid (StructureID, CacheID, OrderNum, BidAmount, MaxQuantity, PlayerID) VALUES (" & _
                        oParentFacility.ObjectID & ", " & oMineralCache.ObjectID & ", " & lOrderNum & ", " & .BidAmount & ", " & _
                        .MaxQuantity & ", " & .lPlayerID & ")"
                    oSB.AppendLine(sSQL)
                    lOrderNum += 1
                End With
            Next X

        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save MineBuyOrder. Reason: " & Err.Description)
        End Try
        Return oSB.ToString
    End Function

    Public Function AddBid(ByVal lAmt As Int32, ByVal lMaxQty As Int32, ByRef oPlayer As Player, ByVal bSave As Boolean) As Int32
        Dim lResult As Int32 = -1

        SyncLock muBids
            'Ok, first, remove my bid...
            Dim lPlacement(mlBidUB) As Int32
            For X As Int32 = 0 To mlBidUB
                lPlacement(X) = muBids(X).lPlayerID
            Next X

            For X As Int32 = 0 To mlBidUB
                If muBids(X).lPlayerID = oPlayer.ObjectID Then
                    For Y As Int32 = X To mlBidUB - 1
                        muBids(Y) = muBids(Y + 1)
                    Next Y
                    mlBidUB -= 1
                    Exit For
                End If
            Next X

            If lAmt < 1 OrElse lMaxQty < 1 Then Return -1

            If lAmt < lMinBidAmt Then Return -1

            Dim lParentID As Int32 = -1
            Dim iParentTypeID As Int16 = -1

            With CType(oParentFacility.ParentObject, Epica_GUID)
                lParentID = .ObjectID
                iParentTypeID = .ObjTypeID
            End With

            Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lParentID, iParentTypeID)
            Dim oColony As Colony = Nothing
            If lColonyIdx > -1 AndAlso lColonyIdx <= glColonyUB Then oColony = goColony(lColonyIdx)
            If oColony Is Nothing Then Return -1

            'Now, determine where to add this one...
            mlBidUB += 1
            ReDim Preserve muBids(mlBidUB)
            muBids(mlBidUB).BidAmount = Int32.MinValue
            For X As Int32 = 0 To mlBidUB
                If muBids(X).BidAmount < lAmt Then
                    'Ok, here it is...
                    For Y As Int32 = mlBidUB To X + 1 Step -1
                        muBids(Y) = muBids(Y - 1)
                    Next Y

                    With muBids(X)
                        .lPlayerID = oPlayer.ObjectID
                        .oPlayer = oPlayer
                        .MaxQuantity = lMaxQty
                        .BidAmount = lAmt
                        .lColonyIdx = lColonyIdx
                        .lColonyID = oColony.ObjectID
                        .lPreviousQuantityReceived = 0
                    End With

                    lResult = X
                    Exit For
                End If
            Next X

            If lResult > -1 AndAlso bSave = True Then
                Dim lNewMax As Int32 = lMinBidAmt
                If mlBidUB > -1 Then lNewMax = muBids(0).BidAmount

                For X As Int32 = 0 To lPlacement.GetUpperBound(0)
                    If mlBidUB >= X AndAlso lPlacement(X) <> oPlayer.ObjectID AndAlso muBids(X).lPlayerID <> lPlacement(X) Then
                        For Y As Int32 = 0 To mlBidUB
                            If lPlacement(X) = muBids(Y).lPlayerID Then
                                If Y > X Then
                                    'MsgBox("Player Demoted: " & muBids(Y).lPlayerID & " from " & X & " to " & Y)
                                    Dim yMsg(17) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eOutBidAlert).CopyTo(yMsg, 0)
                                    oParentFacility.GetGUIDAsString.CopyTo(yMsg, 2)
                                    If oParentFacility.ParentObject Is Nothing = False Then
                                        CType(oParentFacility.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 8)
                                    End If
                                    System.BitConverter.GetBytes(lNewMax).CopyTo(yMsg, 14)
                                    muBids(Y).oPlayer.SendPlayerMessage(yMsg, True, AliasingRights.eViewMining)
                                End If
                                Exit For
                            End If
                        Next Y
                    End If
                Next X

                If oColony.Cargo_Cap < 1 Then
                    oColony.Owner.SendPlayerMessage(oColony.GetLowResourcesMsg(ProductionType.eWareHouse, -1, -1, oMineralCache.oMineral.ObjectID, ObjectType.eMineral), True, AliasingRights.eViewEmail)
                End If
            End If

        End SyncLock

        If bSave = True Then
            Me.SaveObject()
        End If

        Return lResult
    End Function
    Public Sub RemoveBid(ByRef oPlayer As Player)
        AddBid(-1, -1, oPlayer, True)
    End Sub

    ''' <summary>
    ''' Gets the distribution numbers of the mining facility if there was exactly 4 bidders. It calculates the remainders properly.
    ''' </summary>
    ''' <param name="lQty"> Total quantity produced this cycle by the mining facility </param>
    ''' <returns> An array of int32 with the actual quantities on a per placement basis (1st, 2nd, 3rd, 4th) </returns>
    ''' <remarks></remarks>
    Private Function GetDists(ByVal lQty As Int32) As Int32()
        Dim lEveryone As Int32 = lQty \ 4
        Dim lDists() As Int32 = {lEveryone, lEveryone, lEveryone, lEveryone}
        Dim lRemainder As Int32 = lQty - (lEveryone * 4) 'lQty Mod 4
        For X As Int32 = 0 To lRemainder - 1
            lDists(X) += 1
        Next X
        Return lDists
    End Function

    ''' <summary>
    ''' Does the distribution of minerals to the bidders based on their bids.
    ''' Good practice is to call this with the expected quantity to be pulled and then adjust that quantity based on the result. That result is then extracted.
    ''' </summary>
    ''' <param name="lQty"> the number of minerals extracted by the mining facility </param>
    ''' <returns> Int32 of the number of minerals NOT distributed </returns>
    ''' <remarks></remarks>
    Public Function DeliverMinerals(ByVal lQty As Int32) As Int32
        'ok, lQty is our total quantity...
        Try

            lPreviousProductionRate = lQty
            If lQty < 1 Then Return 0
            If mlBidUB < 0 Then Return lQty

            lMinBidAmt = oMineralCache.oMineral.MineralValue
            lMinBidAmt += lMaxDaysSold

            'Dim lMaxDist As Int32 = lQty \ 4
            Dim lDists() As Int32 = GetDists(lQty)

            For X As Int32 = 0 To mlBidUB
                muBids(X).lPreviousQuantityReceived = 0
            Next X

            Dim oPlanetOwner As Player = Nothing
            Dim lEnvirID As Int32 = -1
            Dim iEnvirTypeID As Int16 = -1

            Dim oPlanet As Planet = Nothing

            If oMineralCache.ParentObject Is Nothing = False Then

                With CType(oMineralCache.ParentObject, Epica_GUID)
                    lEnvirID = .ObjectID
                    iEnvirTypeID = .ObjTypeID
                End With

                If CType(oMineralCache.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                    oPlanet = CType(oMineralCache.ParentObject, Planet)

                    If CType(oMineralCache.ParentObject, Planet).OwnerID > -1 Then
                        oPlanetOwner = GetEpicaPlayer(CType(oMineralCache.ParentObject, Planet).OwnerID)
                    End If
                End If
            Else : Return lQty
            End If

            Dim lCurUB As Int32 = mlBidUB 'Math.Min(3, mlBidUB)
            Dim lMineralID As Int32 = Me.oMineralCache.oMineral.ObjectID
            Dim lBidNumber As Int32 = 0
            Dim bNeedToResort As Boolean = False
            For X As Int32 = 0 To lCurUB
                With muBids(X)

                    If .oPlayer.yIronCurtainState = eIronCurtainState.IronCurtainIsUpOnEverything Then
                        RemoveBid(.oPlayer)
                        Return lQty
                    End If

                    'ok... verify the colony still exists
                    If .lColonyIdx < 0 OrElse .lColonyIdx > glColonyUB Then
                        RemoveBid(.oPlayer)
                        Return lQty
                    ElseIf glColonyIdx(.lColonyIdx) <> .lColonyID Then
                        RemoveBid(.oPlayer)
                        Return lQty
                    End If

                    If .BidAmount < lMinBidAmt Then
                        RemoveBid(.oPlayer)
                        Return lQty
                    End If

                    Dim oColony As Colony = goColony(.lColonyIdx)
                    If oColony Is Nothing Then
                        RemoveBid(.oPlayer)
                        Return lQty
                    End If

                    Dim lMaxDist As Int32 = lDists(lBidNumber)

                    Dim lTemp As Int32 = Math.Min(.MaxQuantity, lMaxDist)
                    lTemp = Math.Max(0, Math.Min(lTemp, oColony.Cargo_Cap))

                    If lTemp < 1 Then
                        If X >= 3 Then Exit For
                        Continue For
                        'Return lQty
                    End If

                    'If planet is in corruption, cannot produce minerals
                    If oPlanet Is Nothing = False AndAlso oPlanet.PlanetInCorruption(.oPlayer.lStartedEnvirID, .oPlayer.iStartedEnvirTypeID) = True Then Continue For

                    Dim blCost As Int64 = CLng(lTemp) * CLng(.BidAmount)
                    If .oPlayer.blCredits > blCost Then

                        lBidNumber += 1

                        bSomethingSold = True

                        Dim blOwnerShare As Int64 = (blCost \ 5)

                        .oPlayer.blCredits -= blCost
                        .oPlayer.oBudget.AddMiningBidExpense(blCost, lEnvirID, iEnvirTypeID)
                        oParentFacility.Owner.blCredits += blOwnerShare     'owner gets 20% of the bid
                        oParentFacility.Owner.oBudget.AddMiningBidIncome(blOwnerShare, lEnvirID, iEnvirTypeID)

                        If oPlanetOwner Is Nothing = False Then
                            oPlanetOwner.blCredits += blOwnerShare
                            oPlanetOwner.oBudget.AddMiningBidIncome(blOwnerShare, lEnvirID, iEnvirTypeID)
                        End If

                        muBids(X).MaxQuantity -= lTemp
                        lQty -= lTemp

                        .lPreviousQuantityReceived = lTemp

                        If lTemp <> 0 Then
                            If .oPlayer.CheckFirstContactWithMineral(lMineralID) = True Then
                                If lMineralID = 157 Then lTemp += 30000
                            End If
                        End If

                        'Now, add the amount to the colony
                        oColony.AdjustColonyMineralCache(lMineralID, lTemp)

                        If muBids(X).MaxQuantity < 1 Then
                            bNeedToResort = True
                        End If

                        If lBidNumber > 3 Then
                            Exit For
                        End If
                    End If
                End With
            Next X

            If bNeedToResort = True Then
                Dim bDone As Boolean = False
                While bDone = False
                    bDone = True
                    For X As Int32 = 0 To mlBidUB
                        If muBids(X).BidAmount < 1 OrElse muBids(X).MaxQuantity < 1 Then
                            RemoveBid(muBids(X).oPlayer)
                            bDone = False
                            Continue While
                        End If
                    Next X
                End While
            End If

            If lQty <> 0 Then
                'remainder goes to the top bidder if possible
                If mlBidUB > -1 Then
                    For X As Int32 = 0 To mlBidUB
                        If muBids(X).lPreviousQuantityReceived <> 0 Then
                            With muBids(X)
                                'Ok, see if this bidder has amount remaining...
                                If .MaxQuantity > lQty Then
                                    Dim lTemp As Int32 = Math.Min(.MaxQuantity, lQty)

                                    If .lColonyIdx < 0 OrElse .lColonyIdx > glColonyUB Then
                                        RemoveBid(.oPlayer)
                                        Return lQty
                                    ElseIf glColonyIdx(.lColonyIdx) <> .lColonyID Then
                                        RemoveBid(.oPlayer)
                                        Return lQty
                                    End If

                                    Dim oColony As Colony = goColony(.lColonyIdx)
                                    If oColony Is Nothing Then
                                        RemoveBid(.oPlayer)
                                        Return lQty
                                    End If
                                    lTemp = Math.Min(lTemp, oColony.Cargo_Cap)
                                    lTemp = Math.Max(0, lTemp)
                                    If lTemp < 1 Then Return lQty
                                    lQty -= lTemp
                                    .MaxQuantity -= lTemp
                                    .lPreviousQuantityReceived += lTemp
                                    oColony.AdjustColonyMineralCache(lMineralID, lTemp)
                                    If .MaxQuantity < 1 Then
                                        RemoveBid(.oPlayer)
                                    End If

                                End If
                            End With
                        End If
                    Next X

                End If
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "DeliverMinerals: " & ex.Message)
        End Try

        Return lQty
    End Function

    Public Function GetBidAsMsg(ByVal lForPlayerID As Int32) As Byte()

        Dim yMsg(77) As Byte
        Dim lPos As Int32 = 0

        lMinBidAmt = oMineralCache.oMineral.MineralValue
        lMinBidAmt += lMaxDaysSold

        System.BitConverter.GetBytes(GlobalMessageCode.eMineralBid).CopyTo(yMsg, lPos) : lPos += 2

        System.BitConverter.GetBytes(oParentFacility.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oMineralCache.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMinBidAmt).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPreviousProductionRate).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oMineralCache.Quantity).CopyTo(yMsg, lPos) : lPos += 4

        Dim lPlayerBid As Int32 = 0
        Dim lPlayerQty As Int32 = oMineralCache.Quantity

        Dim lBidsAdded As Int32 = 0
        For X As Int32 = 0 To mlBidUB
            If muBids(X).lPlayerID = lForPlayerID Then
                lPlayerBid = muBids(X).BidAmount
                lPlayerQty = muBids(X).MaxQuantity
            End If
            lBidsAdded += 1
            If lBidsAdded < 5 Then
                System.BitConverter.GetBytes(muBids(X).BidAmount).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(muBids(X).lPreviousQuantityReceived).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(muBids(X).lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
            End If
        Next X
        If lBidsAdded < 5 Then
            lBidsAdded = 4 - lBidsAdded
            lPos += (lBidsAdded * 12)
        End If
        System.BitConverter.GetBytes(lPlayerBid).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(Math.Min(lPlayerQty, oMineralCache.Quantity)).CopyTo(yMsg, lPos) : lPos += 4

        Return yMsg
    End Function

    Private Enum eyMineralBidStatus As Byte
        eNoBid = 0
        eBiddingFirstPlace = 1
        eBiddingSecondPlace = 2
        eBiddingThirdPlace = 3
        eBiddingFourthPlace = 4
        eBiddingNotInTopFour = 5
        eBiddingNotMetMinimum = 6
    End Enum
    Public Function GetPlayerBidStatus(ByVal lPlayerID As Int32) As Byte
        Try
            For X As Int32 = 0 To mlBidUB
                If muBids(X).lPlayerID = lPlayerID Then
                    If muBids(X).BidAmount < lMinBidAmt Then Return eyMineralBidStatus.eBiddingNotMetMinimum

                    Select Case X
                        Case 0
                            Return eyMineralBidStatus.eBiddingFirstPlace
                        Case 1
                            Return eyMineralBidStatus.eBiddingSecondPlace
                        Case 2
                            Return eyMineralBidStatus.eBiddingThirdPlace
                        Case 3
                            Return eyMineralBidStatus.eBiddingFourthPlace
                        Case Else
                            Return eyMineralBidStatus.eBiddingNotInTopFour
                    End Select
                End If
            Next X
        Catch
        End Try
        Return eyMineralBidStatus.eNoBid
    End Function

    Public Function GetPlayerBidSlot(ByVal lPlayerID As Int32) As Int32
        Try
            For X As Int32 = 0 To mlBidUB
                If muBids(X).lPlayerID = lPlayerID Then
                    If muBids(X).BidAmount < lMinBidAmt Then Return -1

                    Select Case X
                        Case 0
                            Return 1
                        Case 1
                            Return 2
                        Case 2
                            Return 3
                        Case 3
                            Return 4
                        Case Else
                            Return 10
                    End Select
                End If
            Next X
        Catch
        End Try
        Return 0
    End Function

    Public Sub ReassessMineralBids()
        If bSomethingSold = False Then
            lCurrentConseqDays = 0
            lDaysNotSold += 1
            lMaxDaysSold = Math.Max(0, lMaxDaysSold - lDaysNotSold)
        Else
            lCurrentConseqDays += 1
            lDaysNotSold = 0
            lMaxDaysSold = Math.Max(lMaxDaysSold, lCurrentConseqDays)
        End If
        bSomethingSold = False

        Try
            lMinBidAmt = oMineralCache.oMineral.MineralValue
            lMinBidAmt += lMaxDaysSold
            For X As Int32 = 0 To mlBidUB
                If muBids(X).BidAmount < lMinBidAmt Then
                    Dim yMsg(17) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eOutBidAlert).CopyTo(yMsg, 0)
                    oParentFacility.GetGUIDAsString.CopyTo(yMsg, 2)
                    If oParentFacility.ParentObject Is Nothing = False Then
                        CType(oParentFacility.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 8)
                    End If
                    System.BitConverter.GetBytes(lMinBidAmt).CopyTo(yMsg, 14)
                    muBids(X).oPlayer.SendPlayerMessage(yMsg, True, AliasingRights.eViewMining)
                End If 
            Next X
        Catch
        End Try

        SaveObject()
    End Sub

    Public Sub SendBiddersDeathAlert(ByVal lKillerPlayerID As Int32, ByVal bDismantle As Boolean)
        Try
            Dim sMsg As String = ""
            If bDismantle = True Then
                Dim oPlayer As Player = GetEpicaPlayer(lKillerPlayerID)
                Dim sPlayer As String = ""
                If oPlayer Is Nothing = False Then sPlayer = oPlayer.sPlayerNameProper
                oPlayer = Nothing

                If sPlayer <> "" Then sMsg = sPlayer & " has dismantled a mine that you were bidding at." Else sMsg = "A mine that you were bidding at has been dismantled."
            Else
                sMsg = "A mining facility that you were bidding at has been destroyed due to conflict in the area."
            End If

            sMsg &= " The mining facility was mining " & BytesToString(Me.oMineralCache.oMineral.MineralName) & "."

            Dim uWP(0) As PlayerComm.WPAttachment
            With uWP(0)
                .AttachNumber = 0
                With CType(oParentFacility.ParentObject, Epica_GUID)
                    uWP(0).EnvirID = .ObjectID
                    uWP(0).EnvirTypeID = .ObjTypeID
                End With
                .LocX = oParentFacility.LocX
                .LocZ = oParentFacility.LocZ
                .yWPNameBytes = StringToBytes("Location")
            End With

            For X As Int32 = 0 To mlBidUB
                If (muBids(X).oPlayer.iInternalEmailSettings And eEmailSettings.eMineralOutbid) <> 0 Then
                    Dim oPC As PlayerComm = muBids(X).oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sMsg, "Bidding Interrupted", Me.oParentFacility.Owner.ObjectID, GetDateAsNumber(Now), False, muBids(X).oPlayer.sPlayerNameProper, uWP)
                    If oPC Is Nothing = False Then muBids(X).oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If

            Next X

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "SendBiddersDeathAlert: " & ex.Message)
        End Try
    End Sub

    Public Function GetFirstPlaceBid() As Int32
        Try
            If mlBidUB > -1 Then
                Return muBids(0).BidAmount
            End If
        Catch
        End Try
    End Function
    Public Function GetFirstPlaceBidderID() As Int32
        Try
            If mlBidUB > -1 Then
                Return muBids(0).lPlayerID
            End If
        Catch
        End Try
        Return -1
    End Function
End Class
