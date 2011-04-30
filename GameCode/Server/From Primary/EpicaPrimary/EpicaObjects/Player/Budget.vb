Option Strict On

Public Enum eyBudgetEnvirFlag As Byte
    UsedThisCycle = 1
    HasCCInEnvir = 2
    ConflictInEnvir = 4
End Enum

Public Class Budget

    Private Const mlBudgetUpdateTime As Int32 = 17280
    Public Const ml_NoCCMultiplier As Int32 = 5
    Private Const mfNonAirMult As Single = ((10000 * 1) / mlBudgetUpdateTime) * 5
    Private Const mfAirCostMult As Single = ((10000 * 10) / mlBudgetUpdateTime) * 5

    Private Shared mlExpLevelLookup() As Int32

    Private Structure BudgetEnvir
        Public lEnvirID As Int32
        Public iEnvirTypeID As Int16

        Public lHWSupplyLine As Int32

        Public lJumpToEnvirID As Int32

        Public lColonyID As Int32
        Public GovScore As Byte

        'Revenue
        Public TaxIncome As Int32
        Public yTaxRate As Byte

        'Expense - Colony based
        Public PopUpkeep As Int32
        Public ResearchCost As Int32
        Public FactoryCost As Int32
        Public SpaceportCost As Int32
        Public OtherFacCost As Int32
        Public Unemployment As Int32
        Public TurretCost As Int32
        Public ExcessStorage As Int32
        'Mining Bid details
        Public blMiningBidIncome As Int64
        Public blMiningBidExpense As Int64

        'Expense - Unit Based
        Public NonAirCost As Int32
        Public AirCost As Int32

        'Public bUsedThisCycle As Boolean
        'Public bHasCC As Boolean
        Public yFlags As Byte

        'Count holders
        Public lAirUnits As Int32
        Public lNonAirUnits As Int32
        Public lDockedAirUnits As Int32
        Public lDockedNonAirUnits As Int32
        Public lCPUsed As Int32
        Public lPrevCPUsed As Int32
        Public lFacilities As Int32

        Public Function GetTotalRevenue() As Int64
            Dim blTemp As Int64 = 0
            blTemp += TaxIncome

            'If lTradeValue Is Nothing = False Then
            '    For X As Int32 = 0 To lTradeValue.GetUpperBound(0)
            '        blTemp += lTradeValue(X)
            '    Next X
            'End If
            Return blTemp
        End Function

        Public Function GetTotalExpense() As Int64
            Return CLng(PopUpkeep) + CLng(ResearchCost) + CLng(FactoryCost) + CLng(SpaceportCost) + CLng(OtherFacCost) + CLng(Unemployment) + CLng(NonAirCost) + CLng(AirCost) + CLng(TurretCost) + CLng(lDockedAirUnits) + CLng(lDockedNonAirUnits) + CLng(ExcessStorage) + CLng(lHWSupplyLine)
        End Function
    End Structure

    Private muItems() As BudgetEnvir
    Public mlItemUB As Int32 = -1

    Private mlAgentMaintCost As Int32 = 0
    Private Const ml_AGENT_FILTER_STATUS As AgentStatus = AgentStatus.Dismissed Or AgentStatus.IsDead Or AgentStatus.HasBeenCaptured Or AgentStatus.NewRecruit
    Public oPlayer As Player        'pointer back to owning player

    Public lTradeValue() As Int32
    Public lTradePlayerID() As Int32
    Private mblTotalTradeIncome As Int64 = 0

    Public lPirateIncome As Int32 = 0
    Public lTotalTradeIncome As Int32 = 0
    Public Function IsEnvironmentPositiveIncome(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Boolean
        For X As Int32 = 0 To mlItemUB
            If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iEnvirTypeID AndAlso (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                Return muItems(X).GetTotalRevenue > muItems(X).GetTotalExpense AndAlso muItems(X).lColonyID > 0
            End If
        Next X

        Return False
    End Function

    Public Sub Reset()
        mlAgentMaintCost = 0
        mblTotalTradeIncome = 0

        ReDim lTradePlayerID(-1)
        ReDim lTradeValue(-1)

        For X As Int32 = 0 To mlItemUB
            With muItems(X)
                '.bUsedThisCycle = False
                '.bHasCC = False
                .yFlags = eyBudgetEnvirFlag.ConflictInEnvir And .yFlags

                '.lTradePlayerID = Nothing
                '.lTradeValue = Nothing
                .lAirUnits = 0
                .lNonAirUnits = 0
                .lPrevCPUsed = .lCPUsed
                .lCPUsed = 0
                .lFacilities = 0
                .TaxIncome = 0
                .lDockedAirUnits = 0
                .lDockedNonAirUnits = 0

                .PopUpkeep = 0
                .lHWSupplyLine = 0
                .ResearchCost = 0
                .FactoryCost = 0
                .SpaceportCost = 0
                .OtherFacCost = 0
                .Unemployment = 0
                .TurretCost = 0
                .ExcessStorage = 0

                .blMiningBidIncome = 0
                .blMiningBidExpense = 0
            End With
        Next X
    End Sub

    Public Sub SetColonyDetails(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lColonyID As Int32, _
      ByVal lTaxIncome As Int32, ByVal lPopUpkeep As Int32, ByVal lResCost As Int32, ByVal lFacCost As Int32, _
      ByVal lSpaceportCost As Int32, ByVal lOtherCost As Int32, ByVal lUnemployedCost As Int32, ByVal bHasCC As Boolean, _
      ByVal yTaxRate As Byte, ByVal lJumpToEnvirID As Int32, ByVal lTurretCost As Int32, ByVal lExcessCargo As Int32, _
      ByVal lTotalFacilities As Int32, ByVal yPlanetPercent As Byte, ByVal lHWSupplyLine As Int32)

        Dim lIdx As Int32 = GetOrAddEnvirItem(lEnvirID, iEnvirTypeID)
        If lIdx = -1 Then Return

        With muItems(lIdx)
            .lHWSupplyLine = lHWSupplyLine
            .lColonyID = lColonyID
            .lFacilities = lTotalFacilities
            .TaxIncome = lTaxIncome
            .PopUpkeep = lPopUpkeep
            .ResearchCost = lResCost
            .FactoryCost = lFacCost
            .SpaceportCost = lSpaceportCost
            .OtherFacCost = lOtherCost
            .Unemployment = lUnemployedCost
            .yTaxRate = yTaxRate
            .lJumpToEnvirID = lJumpToEnvirID
            .TurretCost = lTurretCost
            .ExcessStorage = lExcessCargo
            .GovScore = yPlanetPercent

            '.bHasCC = bHasCC
            If bHasCC = True Then .yFlags = .yFlags Or eyBudgetEnvirFlag.HasCCInEnvir
        End With

    End Sub

    Public Sub FinalizeUnitCosts()
        Dim blTotal As Int64 = 0

        Dim fAirMultBase As Single = oPlayer.oSpecials.fAerialUpkeepMult * mfAirCostMult
        Dim fNonAirMultBase As Single = oPlayer.oSpecials.fNonAerialUpkeepMult * mfNonAirMult

        Dim blCashflow As Int64 = 0

        If oPlayer.yCurrentDoctrine = eDoctrineOfLeadershipSetting.eFlexibility Then
            fAirMultBase *= 0.98F
            fNonAirMultBase *= 0.98F
        End If

        Dim lMaxFac As Int32 = Math.Max(60, CInt(Me.oPlayer.oSpecials.iCPLimit) \ 5I)
        Dim blMaxExp As Int64 = 0
        Dim lMaxExpID As Int32 = -1
        Dim iMaxExpTypeID As Int16 = -1

        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                With muItems(X)
                    Dim fAirMult As Single = fAirMultBase
                    Dim fNonAirMult As Single = fNonAirMultBase

                    If .lFacilities > lMaxFac Then
                        Dim fCPFacMult As Single = 1.0F
                        fCPFacMult = 1.0F + ((.lFacilities - lMaxFac) / 5.0F)
                        fAirMult *= fCPFacMult
                        fNonAirMult *= fCPFacMult
                    End If

                    If (.yFlags And eyBudgetEnvirFlag.HasCCInEnvir) <> 0 Then
                        .AirCost = CInt(.lAirUnits * fAirMult)
                        .NonAirCost = CInt(.lNonAirUnits * fNonAirMult)
                    Else
                        If .lAirUnits = 0 AndAlso .lNonAirUnits = 1 Then Continue For
                        .AirCost = CInt(.lAirUnits * fAirMult * ml_NoCCMultiplier)
                        .NonAirCost = CInt(.lNonAirUnits * fNonAirMult * ml_NoCCMultiplier)
                    End If

                    .lDockedAirUnits = CInt(.lDockedAirUnits * fAirMult * 0.4F)
                    .lDockedNonAirUnits = CInt(.lDockedNonAirUnits * fNonAirMult * 0.4F)
                    blTotal += .lDockedAirUnits
                    blTotal += .lDockedNonAirUnits

                    If .lCPUsed > oPlayer.oSpecials.iCPLimit Then
                        Dim blTemp As Int64 = 0
                        Try
                            blTemp = CLng(Math.Pow((.AirCost + .NonAirCost), Math.Min(CSng(.lCPUsed / oPlayer.oSpecials.iCPLimit), 3)))
                        Catch
                            blTemp = Int32.MaxValue
                        End Try
                        Dim lTotalVal As Int32 = .AirCost + .NonAirCost

                        blTotal += blTemp

                        'Need to adjust the aircost and non air cost accordingly
                        Dim blTempValue As Int64 = CLng(CSng(.AirCost / lTotalVal) * blTemp)
                        If blTempValue > Int32.MaxValue Then blTempValue = Int32.MaxValue
                        If blTempValue < Int32.MinValue Then blTempValue = Int32.MinValue
                        .AirCost = CInt(blTempValue)
                        blTempValue = CLng(CSng(.NonAirCost / lTotalVal) * blTemp)
                        If blTempValue > Int32.MaxValue Then blTempValue = Int32.MaxValue
                        If blTempValue < Int32.MinValue Then blTempValue = Int32.MinValue
                        .NonAirCost = CInt(blTempValue)
                    Else : blTotal += .AirCost + .NonAirCost
                    End If

                    Dim blTempExp As Int64 = .GetTotalExpense
                    If blTempExp > blMaxExp Then
                        blMaxExp = blTempExp
                        lMaxExpID = .lEnvirID
                        iMaxExpTypeID = .iEnvirTypeID
                    End If

                    blCashflow += (.GetTotalRevenue - blTempExp)

                    If .lCPUsed <> .lPrevCPUsed AndAlso (.iEnvirTypeID = ObjectType.ePlanet OrElse .iEnvirTypeID = ObjectType.eSolarSystem) Then
                        'Ok, update rgn
                        Dim yMsg(15) As Byte
                        Dim lPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateCommandPoints).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(Me.oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(.lCPUsed).CopyTo(yMsg, lPos) : lPos += 4

                        Dim oDomain As NetSock = Nothing
                        If .iEnvirTypeID = ObjectType.ePlanet Then
                            Dim oPlanet As Planet = GetEpicaPlanet(.lEnvirID)
                            If oPlanet Is Nothing = False AndAlso oPlanet.oDomain Is Nothing = False Then
                                oDomain = oPlanet.oDomain.DomainSocket
                            End If
                        ElseIf .iEnvirTypeID = ObjectType.eSolarSystem Then
                            Dim oSystem As SolarSystem = GetEpicaSystem(.lEnvirID)
                            If oSystem Is Nothing = False AndAlso oSystem.oDomain Is Nothing = False Then
                                oDomain = oSystem.oDomain.DomainSocket
                            End If
                        End If
                        If oDomain Is Nothing = False Then oDomain.SendData(yMsg)

                    End If
                End With
            End If
        Next X

        If lTradePlayerID Is Nothing = False AndAlso lTradePlayerID.Length > 20 Then
            'ok, gotta pick the top 20
            Dim lKeepers(19) As Int32

            'ok, gotta pick the top 20
            Dim lCurrMin As Int32 = 0
            Dim lCutCnt As Int32 = 0
            For X As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                If lTradeValue(X) > lCurrMin OrElse lCutCnt < 20 Then
                    If lCutCnt < 20 Then
                        lCutCnt += 1
                        lKeepers(lCutCnt - 1) = X
                        If lCutCnt = 20 Then
                            lCurrMin = Int32.MaxValue
                            For Y As Int32 = 0 To 19
                                lCurrMin = Math.Min(lCurrMin, lTradeValue(lKeepers(Y)))
                            Next
                        End If
                    Else
                        Dim lNewMin As Int32 = Int32.MaxValue
                        Dim bSet As Boolean = False
                        For Y As Int32 = 0 To 19
                            If bSet = False AndAlso lTradeValue(lKeepers(Y)) = lCurrMin Then
                                lKeepers(Y) = X
                                bSet = True
                            End If
                            lNewMin = Math.Min(lTradeValue(lKeepers(Y)), lNewMin)
                        Next Y
                        lCurrMin = lNewMin
                    End If
                End If
            Next X 
            mblTotalTradeIncome = 0
            For X As Int32 = 0 To 19
                mblTotalTradeIncome += lTradeValue(lKeepers(X))
            Next X
        End If
        blCashflow += mblTotalTradeIncome
        blTotal -= mblTotalTradeIncome

        mlAgentMaintCost = 0
        For X As Int32 = 0 To oPlayer.mlAgentUB
            If oPlayer.mlAgentIdx(X) <> -1 Then
                If glAgentIdx(oPlayer.mlAgentIdx(X)) = oPlayer.mlAgentID(X) Then
                    Dim oAgent As Agent = goAgent(oPlayer.mlAgentIdx(X))
                    If oAgent Is Nothing = False Then
                        If (oAgent.lAgentStatus And ml_AGENT_FILTER_STATUS) = 0 Then
                            mlAgentMaintCost += oAgent.lMaintCost
                        End If
                    End If
                Else : oPlayer.mlAgentIdx(X) = -1
                End If
            End If
        Next X
        blTotal += mlAgentMaintCost
        blCashflow -= mlAgentMaintCost

        'Now adjust the credits of the player
        If oPlayer.bInFullLockDown = False Then oPlayer.blCredits -= blTotal
        If oPlayer.blCredits < -100000000 Then
            oPlayer.blCredits = -100000000
        Else
            If blCashflow < -3000000000 Then
                LogEvent(LogEventType.ExtensiveLogging, "CashFlowReduction: " & Me.oPlayer.ObjectID & " @ " & blCashflow.ToString)
                LogEvent(LogEventType.ExtensiveLogging, "  MaxExpense: " & lMaxExpID.ToString & ", " & iMaxExpTypeID.ToString & " @ " & blMaxExp.ToString)
            End If
        End If

        If oPlayer.blCredits < 0 Then
            If gb_IS_TEST_SERVER = True Then
                oPlayer.blCredits = Int32.MaxValue
            Else
                'Ok, penalize the player for being negative credits
            End If
        End If
        

        If blCashflow < 0 Then
            oPlayer.bInNegativeCashflow = True
        Else : oPlayer.bInNegativeCashflow = False
        End If
    End Sub

    Private Function GetOrAddEnvirItem(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Int32
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlItemUB
            With muItems(X)
                If .lEnvirID = lEnvirID AndAlso .iEnvirTypeID = iEnvirTypeID Then
                    .yFlags = .yFlags Or eyBudgetEnvirFlag.UsedThisCycle
                    lIdx = X
                    Exit For
                End If
            End With
        Next X

        If lIdx = -1 Then
            mlItemUB += 1
            ReDim Preserve muItems(mlItemUB)
            lIdx = mlItemUB

            With muItems(lIdx)
                .lEnvirID = lEnvirID
                .iEnvirTypeID = iEnvirTypeID
                .yFlags = .yFlags Or eyBudgetEnvirFlag.UsedThisCycle


            End With
        End If
        Return lIdx
    End Function

    Private Const md_TRADE_INCOME_POP_MULT As Single = 1D / 3000000D
    Public Sub AddTraderIncome(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lPlayerID As Int32, ByVal lIncome As Int32)
        Dim lIdx As Int32 = -1

        If oPlayer.yIronCurtainState = eIronCurtainState.IronCurtainIsUpOnEverything Then Return

        If oPlayer.blCurrentPopulation < 3000000 Then
            lIncome = CInt(lIncome * Math.Min(1D, oPlayer.blCurrentPopulation * md_TRADE_INCOME_POP_MULT))
        End If

        'Process our espionage effects here...
        Dim fIncomeMult As Single = 1.0F
        For X As Int32 = 0 To mlAgentEffectUB
            If EffectValid(X) = True Then
                If moAgentEffects(X).yType = AgentEffectType.eTradeIncome Then
                    If moAgentEffects(X).bAmountAsPerc = True Then
                        fIncomeMult *= (moAgentEffects(X).lAmount / 100.0F)
                    Else : lIncome += moAgentEffects(X).lAmount
                    End If
                End If
            End If
        Next X
        lIncome = CInt(lIncome * fIncomeMult)
        If lIncome < 1 Then lIncome = 0

        Dim lTmpIdx As Int32 = -1
        If lTradePlayerID Is Nothing OrElse lTradeValue Is Nothing Then
            ReDim lTradeValue(0)
            ReDim lTradePlayerID(0)
            lTmpIdx = 0
        Else
            lTmpIdx = -1
            For X As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                If lTradePlayerID(X) = lPlayerID Then
                    lTmpIdx = X
                    Exit For
                End If
            Next X

            If lTmpIdx = -1 Then
                ReDim Preserve lTradePlayerID(lTradePlayerID.GetUpperBound(0) + 1)
                lTmpIdx = lTradePlayerID.GetUpperBound(0)
                ReDim Preserve lTradeValue(lTmpIdx)
                lTradeValue(lTmpIdx) = 0
            End If
        End If
        If lTmpIdx <> -1 Then
            lTradePlayerID(lTmpIdx) = lPlayerID
            lTradeValue(lTmpIdx) += lIncome
            mblTotalTradeIncome += lIncome
        End If
    End Sub

    Public Function GetObjAsString() As Byte()
        If oPlayer Is Nothing Then Return Nothing

        Dim lLen As Int32 = 26
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                lLen += 94 '85
            End If
        Next X

        Dim lTPCnt As Int32 = 0

        If lTradePlayerID Is Nothing = False Then
            lTPCnt = (lTradePlayerID.GetUpperBound(0) + 1)

            If lTPCnt > 30 Then
                lTPCnt = 30
            End If

            lLen += (lTPCnt * 8)
        End If

        Dim yMsg(lLen - 1) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(ObjectType.eBudget).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlAgentMaintCost).CopyTo(yMsg, lPos) : lPos += 4
        If Me.oPlayer.lIronCurtainPlanet = -1 Then Me.oPlayer.lIronCurtainPlanet = Me.oPlayer.lStartedEnvirID
        System.BitConverter.GetBytes(Me.oPlayer.lIronCurtainPlanet).CopyTo(yMsg, lPos) : lPos += 4

        Dim blActualDBPop As Int64 = oPlayer.blMaxPopulation - oPlayer.blDBPopulation
        If blActualDBPop < 0 Then blActualDBPop = 0
        If blActualDBPop > 20000000 Then
            System.BitConverter.GetBytes(Int32.MaxValue).CopyTo(yMsg, lPos) : lPos += 4
        Else
            Dim lTmpVal As Int32 = CInt(blActualDBPop * 100)
            System.BitConverter.GetBytes(lTmpVal).CopyTo(yMsg, lPos) : lPos += 4
        End If

        Dim lItemCnt As Int32 = 0
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then lItemCnt += 1
        Next X
        System.BitConverter.GetBytes(lItemCnt).CopyTo(yMsg, lPos) : lPos += 4
        For X As Int32 = 0 To mlItemUB
            With muItems(X)
                If (.yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                    System.BitConverter.GetBytes(.lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
 
                    yMsg(lPos) = .yFlags : lPos += 1

                    System.BitConverter.GetBytes(.lColonyID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.TaxIncome).CopyTo(yMsg, lPos) : lPos += 4

                    System.BitConverter.GetBytes(.PopUpkeep).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.ResearchCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.FactoryCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.SpaceportCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.OtherFacCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.Unemployment).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.NonAirCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.AirCost).CopyTo(yMsg, lPos) : lPos += 4
                    yMsg(lPos) = .yTaxRate : lPos += 1
                    System.BitConverter.GetBytes(.lJumpToEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.TurretCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lDockedAirUnits).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lDockedNonAirUnits).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.ExcessStorage).CopyTo(yMsg, lPos) : lPos += 4

                    System.BitConverter.GetBytes(.blMiningBidExpense).CopyTo(yMsg, lPos) : lPos += 8
                    System.BitConverter.GetBytes(.blMiningBidIncome).CopyTo(yMsg, lPos) : lPos += 8

                    System.BitConverter.GetBytes(.lHWSupplyLine).CopyTo(yMsg, lPos) : lPos += 4

                    Dim yPCapDetail As Byte = 0

                    Dim lSysID As Int32 = -1
                    If .iEnvirTypeID = ObjectType.ePlanet Then
                        Dim oObj As Planet = GetEpicaPlanet(.lEnvirID)
                        If oObj Is Nothing = False Then
                            With oObj
                                If .ParentSystem Is Nothing = False Then lSysID = .ParentSystem.ObjectID

                                yPCapDetail = CByte((.PlanetSizeID * 32) + Math.Min(31, .lColonysHereUB + 1))
                            End With
                        End If
                    ElseIf .iEnvirTypeID = ObjectType.eSolarSystem Then
                        lSysID = .lEnvirID
                    ElseIf .iEnvirTypeID = ObjectType.eFacility Then
                        Dim oObj As Facility = GetEpicaFacility(.lEnvirID)
                        If oObj Is Nothing = False Then
                            Dim oParent As Epica_GUID = oObj.GetRootParentEnvir()
                            If oParent Is Nothing = False Then
                                If oParent.ObjTypeID = ObjectType.ePlanet Then
                                    If CType(oParent, Planet).ParentSystem Is Nothing = False Then lSysID = CType(oParent, Planet).ParentSystem.ObjectID
                                ElseIf oParent.ObjTypeID = ObjectType.eSolarSystem Then
                                    lSysID = oParent.ObjectID
                                End If
                            End If
                        End If
                    End If
                    System.BitConverter.GetBytes(lSysID).CopyTo(yMsg, lPos) : lPos += 4

                    yMsg(lPos) = .GovScore : lPos += 1
                    yMsg(lPos) = yPCapDetail : lPos += 1
                    'If .lTradePlayerID Is Nothing = False Then
                    '    System.BitConverter.GetBytes(.lTradePlayerID.GetUpperBound(0) + 1).CopyTo(yMsg, lPos) : lPos += 4
                    '    For Y As Int32 = 0 To .lTradePlayerID.GetUpperBound(0)
                    '        System.BitConverter.GetBytes(.lTradePlayerID(Y)).CopyTo(yMsg, lPos) : lPos += 4
                    '        System.BitConverter.GetBytes(.lTradeValue(Y)).CopyTo(yMsg, lPos) : lPos += 4
                    '    Next Y
                    'Else : System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
                    'End If
                End If
            End With
        Next X

        If lTradePlayerID Is Nothing = False Then

            System.BitConverter.GetBytes(lTPCnt).CopyTo(yMsg, lPos) : lPos += 4
            If lTradePlayerID.Length > 29 Then
                Dim lKeepers(29) As Int32

                'ok, gotta pick the top 20
                Dim lCurrMin As Int32 = 0
                Dim lCutCnt As Int32 = 0
                For X As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                    If lTradeValue(X) > lCurrMin OrElse lCutCnt < 30 Then
                        If lCutCnt < 30 Then
                            lCutCnt += 1
                            lKeepers(lCutCnt - 1) = X
                            If lCutCnt = 20 Then
                                lCurrMin = Int32.MaxValue
                                For Y As Int32 = 0 To lKeepers.GetUpperBound(0)
                                    lCurrMin = Math.Min(lCurrMin, lTradeValue(lKeepers(Y)))
                                Next
                            End If
                        Else
                            Dim lNewMin As Int32 = Int32.MaxValue
                            Dim bSet As Boolean = False
                            For Y As Int32 = 0 To lKeepers.GetUpperBound(0)
                                If bSet = False AndAlso lTradeValue(lKeepers(Y)) = lCurrMin Then
                                    lKeepers(Y) = X
                                    bSet = True
                                End If
                                lNewMin = Math.Min(lTradeValue(lKeepers(Y)), lNewMin)
                            Next Y
                            lCurrMin = lNewMin
                        End If
                    End If
                Next X

                For X As Int32 = 0 To lKeepers.GetUpperBound(0)
                    System.BitConverter.GetBytes(lTradePlayerID(lKeepers(X))).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lTradeValue(lKeepers(X))).CopyTo(yMsg, lPos) : lPos += 4
                Next X
            Else
                For X As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                    System.BitConverter.GetBytes(lTradePlayerID(X)).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lTradeValue(X)).CopyTo(yMsg, lPos) : lPos += 4
                Next X
            End If


        Else
            System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
        End If

        Return yMsg
    End Function

    Public Sub AddFacilityToEnvir(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lCnt As Int32)
        Dim lIdx As Int32 = GetOrAddEnvirItem(lEnvirID, iEnvirTypeID)

        With muItems(lIdx)
            .lFacilities += lCnt
            .lEnvirID = lEnvirID
            .iEnvirTypeID = iEnvirTypeID
        End With
    End Sub

    Public Sub AddUnitToEnvir(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal bAirUnit As Boolean, ByVal yExpLevel As Byte)
        Dim lIdx As Int32 = GetOrAddEnvirItem(lEnvirID, iEnvirTypeID)

        With muItems(lIdx)
            If bAirUnit = True Then .lAirUnits += 1 Else .lNonAirUnits += 1
            .lCPUsed += mlExpLevelLookup(yExpLevel) + Me.oPlayer.BadWarDecCPIncrease
        End With
    End Sub

    Public Sub AddDockedUnitToEnvir(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal bAirUnit As Boolean)
        Dim lIdx As Int32 = GetOrAddEnvirItem(lEnvirID, iEnvirTypeID)

        With muItems(lIdx)
            If bAirUnit = True Then .lDockedAirUnits += 1 Else .lDockedNonAirUnits += 1
        End With
    End Sub

    Public Sub AddMiningBidExpense(ByVal blExpense As Int64, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
        Dim lIdx As Int32 = GetOrAddEnvirItem(lEnvirID, iEnvirTypeID)

        With muItems(lIdx)
            .blMiningBidExpense += blExpense
        End With
    End Sub
    Public Sub AddMiningBidIncome(ByVal blIncome As Int64, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
        Dim lIdx As Int32 = GetOrAddEnvirItem(lEnvirID, iEnvirTypeID)

        With muItems(lIdx)
            .blMiningBidIncome += blIncome
        End With
    End Sub

    Public Sub New()
        If mlExpLevelLookup Is Nothing Then
            ReDim mlExpLevelLookup(255)
            For X As Int32 = 0 To 255
                mlExpLevelLookup(X) = Math.Max(10 - (X \ 25), 1)
            Next X
        End If
    End Sub

    Public Function GetDockingDelayMult(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Single
        For X As Int32 = 0 To mlItemUB
            If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iEnvirTypeID Then
                If muItems(X).lCPUsed < oPlayer.oSpecials.iCPLimit Then Return 1.0F

                Return ((muItems(X).lCPUsed * 1.1F) / oPlayer.oSpecials.iCPLimit)
            End If
        Next X
        Return 1.0F
    End Function

    Public Function GetTotalUnitCount() As Int32
        Dim lResult As Int32 = 0
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                lResult += muItems(X).lNonAirUnits + muItems(X).lAirUnits
            End If
        Next X
        Return lResult
    End Function

    Public Function GetCashFlow() As Int64
        Dim lResult As Int64 = 0
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                lResult += (muItems(X).GetTotalRevenue + muItems(X).blMiningBidIncome) - (muItems(X).GetTotalExpense + muItems(X).blMiningBidExpense)
            End If
        Next X
        lResult += mblTotalTradeIncome - mlAgentMaintCost
        Return lResult
    End Function

    Public Function IsInEnvironment(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Boolean
        For X As Int32 = 0 To mlItemUB
            If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iEnvirTypeID Then Return True
        Next X
        Return False
    End Function

    Public Function GetEnvirCPUsage(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Int32
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iEnvirTypeID Then
                    Return muItems(X).lCPUsed
                End If
            End If
        Next X
        Return 0
    End Function

    Public Function EnvironmentInConflict(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Boolean
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iEnvirTypeID Then Return (muItems(X).yFlags And eyBudgetEnvirFlag.ConflictInEnvir) <> 0
            End If
        Next X
        Return False
    End Function

    Public Sub SetEnvironmentInConflict(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal bValue As Boolean)
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iEnvirTypeID Then
                    If bValue = True Then
                        muItems(X).yFlags = muItems(X).yFlags Or eyBudgetEnvirFlag.ConflictInEnvir
                    Else
                        If (muItems(X).yFlags And eyBudgetEnvirFlag.ConflictInEnvir) <> 0 Then muItems(X).yFlags = muItems(X).yFlags Xor eyBudgetEnvirFlag.ConflictInEnvir
                    End If
                    Exit For
                End If
            End If
        Next X
    End Sub

    Public Function GetTotalIncome() As Int64
        Dim lResult As Int64 = 0
        For X As Int32 = 0 To mlItemUB
            If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 Then
                lResult += (muItems(X).GetTotalRevenue + muItems(X).blMiningBidIncome) '- (muItems(X).GetTotalExpense + muItems(X).blMiningBidExpense)
            End If
        Next X
        lResult += mblTotalTradeIncome
        Return lResult
    End Function

    Public Function GetFacilityPointUsage(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Int32
        For X As Int32 = 0 To mlItemUB
            If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iEnvirTypeID Then
 
                Return muItems(X).lFacilities
            End If
        Next X
        Return 0
    End Function

#Region "  Agent Effects versus this Budget  "
    Protected moAgentEffects() As AgentEffect
    Private myAgentEffectUsed() As Byte
    Protected mlAgentEffectUB As Int32 = -1

    Public Function AddAgentEffect(ByVal lStartCycle As Int32, ByVal lDuration As Int32, ByVal yType As AgentEffectType, ByVal lAmount As Int32, ByVal bAsPercentage As Boolean, ByVal lCausedByID As Int32) As AgentEffect

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlAgentEffectUB
            If myAgentEffectUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            mlAgentEffectUB += 1
            ReDim Preserve myAgentEffectUsed(mlAgentEffectUB)
            ReDim Preserve moAgentEffects(mlAgentEffectUB)
            lIdx = mlAgentEffectUB
        End If

        moAgentEffects(lIdx) = New AgentEffect
        With moAgentEffects(lIdx)
            .bAmountAsPerc = bAsPercentage
            .lAmount = lAmount
            .lDuration = lDuration
            .lStartCycle = lStartCycle
            .yType = yType
            .lCausedByID = lCausedByID
        End With
        myAgentEffectUsed(lIdx) = 255

        Return moAgentEffects(lIdx)
    End Function

    Protected Sub RemoveAgentEffect(ByVal lIdx As Int32)
        myAgentEffectUsed(lIdx) = 0
    End Sub

    Protected Function EffectValid(ByVal lIdx As Int32) As Boolean
        Try
            If myAgentEffectUsed(lIdx) = 0 Then Return False
            If moAgentEffects(lIdx).lLastVerification = glCurrentCycle Then Return True
            If moAgentEffects(lIdx).lStartCycle <= glCurrentCycle Then
                If moAgentEffects(lIdx).lStartCycle + moAgentEffects(lIdx).lDuration > glCurrentCycle Then
                    moAgentEffects(lIdx).lLastVerification = glCurrentCycle
                    Return True
                Else : RemoveAgentEffect(lIdx)
                End If
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Budget.EffectValid: " & ex.Message)
        End Try
        Return False
    End Function

    Protected Function SaveAgentEffects() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            sSQL = "DELETE FROM tblAgentEffects WHERE EffectedItemID = " & Me.oPlayer.ObjectID & " AND EffectedItemTypeID = " & ObjectType.eBudget
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            'Now, all effects are inserts
            For X As Int32 = 0 To mlAgentEffectUB
                If myAgentEffectUsed(X) <> 0 Then
                    With moAgentEffects(X)
                        Try
                            sSQL = "INSERT INTO tblAgentEffects (EffectedItemID, EffectedItemTypeID, RemainingCycles, EffectType, EffectAmount, " & _
                              "yAmountAsPerc, CausedByID) VALUES (" & Me.oPlayer.ObjectID & ", " & ObjectType.eBudget & ", " & .lDuration - (glCurrentCycle - .lStartCycle) & _
                              ", " & CByte(.yType) & ", " & .lAmount & ", "
                            If .bAmountAsPerc = True Then sSQL &= "1," Else sSQL &= "0,"
                            sSQL &= .lCausedByID & ")"
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            oComm.ExecuteNonQuery()
                            oComm.Dispose()
                            oComm = Nothing
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "Unable to Save Agent Effect: " & ex.Message)
                        End Try
                    End With
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save agent effect. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    'Public Sub HandleTradeTreatyEspionage(ByRef oAttacker As Player)
    '    If lTradePlayerID Is Nothing OrElse lTradeValue Is Nothing Then Return
    '    For Y As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
    '        Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(lTradePlayerID(Y), ObjectType.eTrade, PlayerItemIntel.PlayerItemIntelType.eFullKnowledge, Me.oPlayer.ObjectID)
    '        If oPII Is Nothing = False Then
    '            oPII.lValue = lTradeValue(Y)
    '            oAttacker.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
    '        End If
    '    Next Y
    'End Sub

    Public Function GetEnvirBudgetText(ByVal lEnvirID As Int32, ByVal iTypeID As Int16) As String
        Dim oSB As New System.Text.StringBuilder
        oSB.AppendLine("Budget Findings:")

        If lEnvirID = -1 OrElse iTypeID = -1 OrElse (iTypeID <> ObjectType.ePlanet AndAlso iTypeID <> ObjectType.eSolarSystem AndAlso iTypeID <> ObjectType.eFacility) Then
            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To mlItemUB
                If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 AndAlso muItems(X).lColonyID <> -1 Then
                    lCnt += 1
                End If
            Next X
            If lCnt <> 0 Then
                Dim lVal As Int32 = CInt(Rnd() * lCnt)
                For X As Int32 = 0 To mlItemUB
                    If (muItems(X).yFlags And eyBudgetEnvirFlag.UsedThisCycle) <> 0 AndAlso muItems(X).lColonyID <> -1 Then
                        lEnvirID = muItems(X).lEnvirID
                        iTypeID = muItems(X).iEnvirTypeID
                        lVal -= 1
                        If lVal <= 0 Then
                            lEnvirID = muItems(X).lEnvirID
                            iTypeID = muItems(X).iEnvirTypeID
                            Exit For
                        End If
                    End If
                Next X
            Else
                oSB.AppendLine("No valid targets found.")
                Return oSB.ToString
            End If
        End If

        Dim bFound As Boolean = False
        For X As Int32 = 0 To mlItemUB
            If muItems(X).lEnvirID = lEnvirID AndAlso muItems(X).iEnvirTypeID = iTypeID Then
                bFound = True
                With muItems(X)
                    Dim oColony As Colony = Nothing
                    If .lColonyID <> -1 Then oColony = GetEpicaColony(.lColonyID)
                    Dim sTargetName As String = ""
                    If oColony Is Nothing = False Then sTargetName = BytesToString(oColony.ColonyName) & " colony"
                    oColony = Nothing

                    If .iEnvirTypeID = ObjectType.ePlanet Then
                        Dim oPlanet As Planet = GetEpicaPlanet(.lEnvirID)
                        If oPlanet Is Nothing = False Then
                            If sTargetName = "" Then sTargetName = BytesToString(oPlanet.PlanetName) Else sTargetName &= " (" & BytesToString(oPlanet.PlanetName) & ")"
                        End If
                        oPlanet = Nothing
                    ElseIf .iEnvirTypeID = ObjectType.eSolarSystem Then
                        Dim oSystem As SolarSystem = GetEpicaSystem(.lEnvirID)
                        If oSystem Is Nothing = False Then
                            If sTargetName = "" Then sTargetName = BytesToString(oSystem.SystemName) Else sTargetName &= " (" & BytesToString(oSystem.SystemName) & ")"
                        End If
                        oSystem = Nothing
                    End If

                    If sTargetName = "" Then
                        oSB.AppendLine("Your agents could not find any budget records on the specified target.")
                    Else
                        oSB.AppendLine("Your agents have finished assessing the budget records stolen from " & Me.oPlayer.sPlayerNameProper & " regarding " & sTargetName & ".")
                        oSB.AppendLine()
                        oSB.AppendLine("The results are:")
                        If .TaxIncome > 0 Then
                            oSB.AppendLine("Tax Income: " & .TaxIncome.ToString("#,##0"))
                            oSB.AppendLine("Tax Rate: " & .yTaxRate.ToString)
                        End If 
                        If mblTotalTradeIncome <> 0 Then
                            oSB.AppendLine("Trade Income: " & mblTotalTradeIncome.ToString("#,##0"))
                        End If

                        oSB.AppendLine()
                        'now, expenses
                        oSB.AppendLine("Homeworld Supply: " & .lHWSupplyLine.ToString("#,##0"))
                        oSB.AppendLine("Population Upkeep: " & .PopUpkeep.ToString("#,##0"))
                        oSB.AppendLine("Research Facilities: " & .ResearchCost.ToString("#,##0"))
                        oSB.AppendLine("Factories: " & .FactoryCost.ToString("#,##0"))
                        oSB.AppendLine("Spaceports: " & .SpaceportCost.ToString("#,##0"))
                        oSB.AppendLine("Defenses: " & .TurretCost.ToString("#,##0"))
                        oSB.AppendLine("Excess Storage: " & .ExcessStorage.ToString("#,##0"))
                        oSB.AppendLine("Other Facilities: " & .OtherFacCost.ToString("#,##0"))
                        oSB.AppendLine("Unemployment: " & .Unemployment.ToString("#,##0"))
                        oSB.AppendLine("Non Aerial Support: " & .NonAirCost.ToString("#,##0"))
                        oSB.AppendLine("Aerial Support: " & .AirCost.ToString("#,##0"))
                        oSB.AppendLine("Docked Non Aerial Support: " & .lDockedNonAirUnits.ToString("#,##0"))
                        oSB.AppendLine("Docked Aerial Support: " & .lDockedAirUnits.ToString("#,##0"))
                        If (.yFlags And eyBudgetEnvirFlag.HasCCInEnvir) <> 0 Then
                            oSB.AppendLine("Has a Command Center")
                        Else : oSB.AppendLine("Does not have a Command Center")
                        End If
                    End If
                End With

                Exit For
            End If
        Next X

        If bFound = False Then
            oSB.AppendLine("No valid targets found.")
        End If

        Return oSB.ToString
    End Function

    Public Function GetTradeIncome(ByVal lPlayerID As Int32) As Int64
        Dim blTotal As Int64 = 0
        If lTradePlayerID Is Nothing OrElse lTradeValue Is Nothing Then Return 0
        For Y As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
            If lTradePlayerID(Y) = lPlayerID Then
                blTotal += lTradeValue(Y)
                Exit For
            End If
        Next Y
        Return blTotal
    End Function

    Public Function AuditRemoveAgentEffects(ByVal lTargetNum As Int32) As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlAgentEffectUB
            If EffectValid(X) = True Then
                If Rnd() * 100 < lTargetNum Then
                    RemoveAgentEffect(X)
                    lCnt += 1
                End If
            End If
        Next X
        Return lCnt
    End Function
#End Region
End Class
