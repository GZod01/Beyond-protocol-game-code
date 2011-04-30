Option Strict On

Public Enum eyBudgetEnvirFlag As Byte
    UsedThisCycle = 1
    HasCCInEnvir = 2
    ConflictInEnvir = 4
End Enum

Public Class Budget
    Public Structure BudgetEnvir
        Public lEnvirID As Int32
        Public iEnvirTypeID As Int16
        Public lColonyID As Int32
        Public TaxIncome As Int32
        Public PopUpkeep As Int32
        Public ResearchCost As Int32
        Public FactoryCost As Int32
        Public SpaceportCost As Int32
        Public OtherFacCost As Int32
		Public UnemploymentCost As Int32
		Public TurretCost As Int32
        Public NonAirCost As Int32
        Public AirCost As Int32
        Public yTaxRate As Byte



		Public DockedAirCost As Int32
		Public DockedNonAirCost As Int32

        Public ExcessStorage As Int32

        Public blMiningBidExpense As Int64
        Public blMiningBidIncome As Int64

        Public yPlanetaryControl As Byte

        Public lJumpToID As Int32

        Public lSystemID As Int32
        Public lHWSupplyLineCost As Int32

        Public lChildrenItems As Int32

        Public bHasCC As Boolean
        Public bHasConflict As Boolean

        Public iTotalStarIncome As Int64
        Public iTotalStarExpense As Int64

        Public lPlanetCap As Int32
        Public lCurrentColonyCount As Int32

        Public Function GetTotalRevenue() As Int64
            Dim blTemp As Int64 = 0
            blTemp += TaxIncome

            Return blTemp + blMiningBidIncome
        End Function
        Public Function GetTotalExpense() As Int64
            Return CLng(PopUpkeep) + CLng(ResearchCost) + CLng(FactoryCost) + CLng(SpaceportCost) + CLng(OtherFacCost) + CLng(UnemploymentCost) + CLng(NonAirCost) + CLng(AirCost) + CLng(TurretCost) + CLng(DockedAirCost) + CLng(DockedNonAirCost) + CLng(ExcessStorage) + blMiningBidExpense + CLng(lHWSupplyLineCost)
        End Function

        Public Sub SetPCapValue(ByVal yVal As Byte)
            lPlanetCap = 0
            lCurrentColonyCount = 0
            If yVal <> 0 Then
                'yPCapDetail = CByte((.PlanetSizeID * 32) + Math.Min(31, .lColonysHereUB + 1))
                lPlanetCap = (yVal And 224)
                yVal -= CByte(lPlanetCap)
                lPlanetCap \= 32
                lPlanetCap += 2
                lCurrentColonyCount = yVal
            End If
        End Sub

        'MSC - 1/1/2009 - removed because players HATED it... might not be a bad idea for something in the future if we can make the initial impact less... impacting
        'Public Function GetHomeworldSupplyLineCost(ByVal lHWID As Int32) As Int64
        '    Dim blValue As Int64 = 0
        '    Try
        '        If PopUpkeep = 0 Then Return 0

        '        If lHWID = -1 Then Return 0

        '        If goGalaxy Is Nothing = False Then

        '            Dim lHWLocX As Int32 = 0
        '            Dim lHWLocY As Int32 = 0
        '            Dim lHWLocZ As Int32 = 0
        '            Dim lGTCSpeed As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eTradeGTCSpeed)
        '            lGTCSpeed = (lGTCSpeed * 200) + 200
        '            lGTCSpeed *= 30

        '            Dim lHWSysID As Int32 = -1

        '            Dim sEnvirName As String = GetCacheObjectValue(lEnvirID, iEnvirTypeID)
        '            Dim sHWName As String = GetCacheObjectValue(lHWID, ObjectType.ePlanet)

        '            Dim oEnvirSys As SolarSystem = Nothing

        '            For X As Int32 = 0 To goGalaxy.mlSystemUB
        '                Dim oSys As SolarSystem = goGalaxy.moSystems(X)
        '                If oSys Is Nothing = False AndAlso oSys.SystemName Is Nothing = False Then
        '                    Dim sCompare As String = oSys.SystemName
        '                    sCompare = sCompare.Replace("(S)", "")
        '                    If sEnvirName.StartsWith(sCompare) = True Then
        '                        oEnvirSys = oSys
        '                    End If
        '                    If sHWName.StartsWith(sCompare) = True Then
        '                        lHWSysID = oSys.ObjectID
        '                        lHWLocX = oSys.LocX
        '                        lHWLocY = oSys.LocY
        '                        lHWLocZ = oSys.LocZ
        '                    End If
        '                End If
        '                If lHWSysID <> -1 AndAlso oEnvirSys Is Nothing = False Then Exit For
        '            Next X

        '            If oEnvirSys Is Nothing = False AndAlso lHWSysID <> -1 Then

        '                If oEnvirSys.ObjectID = lHWSysID Then Return 0

        '                Dim fDX As Single = oEnvirSys.LocX - lHWLocX
        '                Dim fDY As Single = oEnvirSys.LocY - lHWLocY
        '                Dim fDZ As Single = oEnvirSys.LocZ - lHWLocZ
        '                fDX *= fDX
        '                fDY *= fDY
        '                fDZ *= fDZ

        '                Dim dDist As Double = Math.Sqrt(fDX + fDY + fDZ)
        '                'Now, that is in system sectors... so, we'll cheat and add 1 more sector for the inter-system movements
        '                dDist += 1.0F
        '                dDist *= 10000000

        '                Dim lTotalTime As Int32 = CInt(dDist / lGTCSpeed)
        '                lTotalTime \= 2
        '                'Dim fVal As Single = lTotalTime   'lTotalTime / 86400.0F

        '                'Now, determine the population...
        '                'if nocc, lccmult = 5 else lccmult = 1
        '                'Dim lPopUpkeep As Int32 = CInt(Population * mfPopUpkeepMult * .fPopulationUpkeepMult) * lCCMult
        '                Const mfPopUpkeepMult As Single = 0.00480324076F '(16.6F / 17280.0F) * 5
        '                'At present, fPopulationUpkeepMult is always 1
        '                Dim fPopulationUpkeepMult As Single = 1.0F
        '                Dim lCCMult As Int32 = 1
        '                If bHasCC = False Then lCCMult = 5

        '                'determine our population from our popupkeep
        '                Dim fTemp As Single = ((mfPopUpkeepMult * fPopulationUpkeepMult * lCCMult) / PopUpkeep)
        '                Try
        '                    Dim lPopulation As Int32 = CInt(1.0F / fTemp)

        '                    'blValue = CLng((lPopulation * fVal) / 10)
        '                    blValue = Math.Max(0, lTotalTime - (lPopulation \ 10))
        '                Catch
        '                End Try
        '            End If
        '        End If
        '    Catch
        '    End Try
        '    Return blValue
        'End Function

        'This is the accepted method of doing it...
        'Public Function GetHomeworldSupplyLineCost(ByVal lHWID As Int32) As Int32

        '    If PopUpkeep = 0 Then Return 0
        '    Dim lValue As Int32 = 0

        '    Try
        '        Dim lHWSysID As Int32 = -1

        '        Dim sEnvirName As String = GetCacheObjectValue(lEnvirID, iEnvirTypeID)
        '        Dim sHWName As String = GetCacheObjectValue(lHWID, ObjectType.ePlanet)

        '        Dim oEnvirSys As SolarSystem = Nothing

        '        For X As Int32 = 0 To goGalaxy.mlSystemUB
        '            Dim oSys As SolarSystem = goGalaxy.moSystems(X)
        '            If oSys Is Nothing = False AndAlso oSys.SystemName Is Nothing = False Then
        '                Dim sCompare As String = oSys.SystemName
        '                sCompare = sCompare.Replace("(S)", "")
        '                If sEnvirName.StartsWith(sCompare) = True Then
        '                    oEnvirSys = oSys
        '                End If
        '                If sHWName.StartsWith(sCompare) = True Then
        '                    lHWSysID = oSys.ObjectID
        '                End If
        '            End If
        '            If lHWSysID <> -1 AndAlso oEnvirSys Is Nothing = False Then Exit For
        '        Next X

        '        If oEnvirSys Is Nothing = False AndAlso lHWSysID <> -1 Then
        '            If oEnvirSys.ObjectID = lHWSysID Then Return 0

        '            'Now, determine the population...
        '            Const mfPopUpkeepMult As Single = 0.00480324076F '(16.6F / 17280.0F) * 5
        '            'At present, fPopulationUpkeepMult is always 1
        '            Dim fPopulationUpkeepMult As Single = 1.0F
        '            Dim lCCMult As Int32 = 1
        '            If bHasCC = False Then lCCMult = 5

        '            'determine our population from our popupkeep
        '            Dim fTemp As Single = ((mfPopUpkeepMult * fPopulationUpkeepMult * lCCMult) / PopUpkeep)
        '            Try
        '                Dim lPopulation As Int32 = CInt(1.0F / fTemp)

        '                If lPopulation > 2000000 Then Return 0

        '                lValue = CInt((1 - (lPopulation / 2000000.0F)) * 61000)

        '                'blValue = Math.Max(0, lTotalTime - (lPopulation \ 10))
        '            Catch
        '            End Try
        '        End If
        '    Catch
        '    End Try

        '    Return lValue
        'End Function
    End Structure


    Public Enum BudgetSort As Int32
        EnvironmentType = 1
        EnvironmentName = 2
        ColonyName = 3
        Revenue = 4
        Expense = 5
        TaxRate = 6
        Control = 7

        DescendingOrderShift = 1073741824
    End Enum

    Public lPlayerID As Int32

    Public muItems() As BudgetEnvir
    Public mlItemUB As Int32 = -1

    Private muSrcList() As BudgetEnvir
    Private mlSrcListUB As Int32 = -1

	Public lLastUpdateCycle As Int32

	Public lAgentMaintCost As Int32 = 0

	Public lIronCurtainPlanet As Int32 = -1

	Public lMaxDeathBudget As Int32 = 0

    ' Used to sum the total trade income for the player
    Public TotalTradeIncome As Int32

    Public lTradeValue() As Int32
    Public lTradePlayerID() As Int32

    Private mlLastSort As Int32 = 0
    Public Sub ResetLastSortTime()
        mlLastSort = 0
    End Sub

    Private mblTotalRevenue As Int64 = Int64.MinValue
    Public ReadOnly Property TotalRevenue() As Int64
        Get
            If mblTotalRevenue = Int64.MinValue Then
                mblTotalRevenue = 0

                For X As Int32 = 0 To mlItemUB
                    mblTotalRevenue += muItems(X).GetTotalRevenue
                Next X
                If lTradePlayerID Is Nothing = False Then
                    For X As Int32 = 0 To Math.Min(19, lTradePlayerID.GetUpperBound(0))
                        mblTotalRevenue += lTradeValue(X)
                    Next X
                End If
            End If
            
            Return mblTotalRevenue
        End Get
    End Property

    Private mblTotalExpense As Int64 = Int64.MinValue
    Public ReadOnly Property TotalExpense() As Int64
        Get
            If mblTotalExpense = Int64.MinValue Then
                mblTotalExpense = 0

                For X As Int32 = 0 To mlItemUB
                    mblTotalExpense += muItems(X).GetTotalExpense
                Next X
            End If

            Return mblTotalExpense
        End Get
    End Property

    Public Sub FillFromMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += 2   'for typeid

        If lPlayerID < 0 Then
            lPlayerID = Math.Abs(lPlayerID)
            lAgentMaintCost = CInt(System.BitConverter.ToInt64(yData, lPos)) : lPos += 8
        Else
            lAgentMaintCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        End If

        lIronCurtainPlanet = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        lMaxDeathBudget = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
            lMaxDeathBudget = 100000000
        End If

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim muSrcList(lCnt - 1)
        mlSrcListUB = lCnt - 1


        For X As Int32 = 0 To mlSrcListUB
            Dim oNewEnvir As BudgetEnvir
            With oNewEnvir
                .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                '.bHasCC = yData(lPos) <> 0 : lPos += 1
                Dim yTemp As Byte = yData(lPos) : lPos += 1
                .bHasCC = (yTemp And eyBudgetEnvirFlag.HasCCInEnvir) <> 0
                .bHasConflict = (yTemp And eyBudgetEnvirFlag.ConflictInEnvir) <> 0

                .lColonyID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .TaxIncome = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .PopUpkeep = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ResearchCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .FactoryCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .SpaceportCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .OtherFacCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .UnemploymentCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .NonAirCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .AirCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yTaxRate = yData(lPos) : lPos += 1
                .lJumpToID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .TurretCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .DockedAirCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .DockedNonAirCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ExcessStorage = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                .blMiningBidExpense = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .blMiningBidIncome = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

                .lHWSupplyLineCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSystemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                .yPlanetaryControl = yData(lPos) : lPos += 1
                If .iEnvirTypeID = ObjectType.eSolarSystem OrElse .iEnvirTypeID = ObjectType.eFacility Then .yPlanetaryControl = 0
                If .lColonyID < 1 Then .yPlanetaryControl = 0

                Dim yPCapDetails As Byte = yData(lPos) : lPos += 1
                .SetPCapValue(yPCapDetails)
            End With
            muSrcList(X) = oNewEnvir

            'Fear It
            'If isAdmin() = True Then
            '    If muSrcList(X).iEnvirTypeID = ObjectType.eSolarSystem Then
            '        Dim sName As String = GetCacheObjectValue(muSrcList(X).lEnvirID, muSrcList(X).iEnvirTypeID)
            '        For Y As Int32 = 0 To X - 1
            '            Dim sOtherName As String = GetCacheObjectValue(muSrcList(Y).lEnvirID, muSrcList(Y).iEnvirTypeID)
            '            If sName.ToUpper > sOtherName.ToUpper Then
            '                Dim oOldEnvir As BudgetEnvir = muSrcList(Y)
            '                muSrcList(X) = oOldEnvir
            '                muSrcList(Y) = oNewEnvir
            '                Exit For
            '            End If
            '        Next
            '    End If
            'End If
            oNewEnvir = Nothing
        Next X

        For X As Int32 = 0 To mlSrcListUB
            With muSrcList(X)
                If .iEnvirTypeID <> ObjectType.eSolarSystem Then
                    For y As Int32 = 0 To mlSrcListUB
                        If muSrcList(y).lSystemID = .lSystemID AndAlso muSrcList(y).iEnvirTypeID = ObjectType.eSolarSystem Then
                            muSrcList(y).iTotalStarIncome += muSrcList(X).GetTotalRevenue
                            muSrcList(y).iTotalStarExpense += muSrcList(X).GetTotalExpense
                        End If
                    Next
                End If
            End With
        Next X


        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
        ReDim lTradePlayerID(lUB)
        ReDim lTradeValue(lUB)

        For X As Int32 = 0 To lUB
            lTradePlayerID(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lTradeValue(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Next X

        'Ok, sort
        If lTradePlayerID Is Nothing = False Then 'AndAlso lTradePlayerID.Length > 20 Then
            'ok, gotta pick the top 20
            Dim lSorted(lTradePlayerID.GetUpperBound(0)) As Int32 '= Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                Dim lIdx As Int32 = -1

                For Y As Int32 = 0 To lSortedUB
                    If lTradeValue(lSorted(Y)) < lTradeValue(X) Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                'ReDim Preserve lSorted(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X

            Dim lNewID(lTradePlayerID.GetUpperBound(0)) As Int32
            Dim lNewVal(lTradePlayerID.GetUpperBound(0)) As Int32
            For X As Int32 = 0 To lTradePlayerID.GetUpperBound(0)
                lNewID(X) = lTradePlayerID(lSorted(X))
                lNewVal(X) = lTradeValue(lSorted(X))
            Next X
            lTradePlayerID = lNewID
            lTradeValue = lNewVal
        End If

        mblTotalExpense = Int64.MinValue
        mblTotalRevenue = Int64.MinValue

        mlLastSort = 0

        SortItems(mlLastSortType)

        lLastUpdateCycle = glCurrentCycle
    End Sub

    Private mlLastSortType As BudgetSort = BudgetSort.EnvironmentName
    Public Sub SortItems(ByVal lSortType As BudgetSort)
        mlLastSortType = lSortType
        If glCurrentCycle - mlLastSort < 30 Then Return
        mlLastSort = glCurrentCycle

        Dim bDescending As Boolean = (lSortType And BudgetSort.DescendingOrderShift) <> 0
        If bDescending = True Then lSortType = lSortType Xor BudgetSort.DescendingOrderShift

        muItems = muSrcList
        mlItemUB = mlSrcListUB

        'Ok, first... we need to ensure that each envir entry is represented by a system line item
        Dim lNeedToAddSystems() As Int32 = Nothing
        Dim lNeedToAddUB As Int32 = -1
        Dim lSystems As Int32 = 0
        For X As Int32 = 0 To mlItemUB
            If muItems(X).iEnvirTypeID <> ObjectType.eSolarSystem Then
                Dim bFound As Boolean = False

                'verify my system is not already in the need to add list
                For Y As Int32 = 0 To lNeedToAddUB
                    If lNeedToAddSystems(Y) = muItems(X).lSystemID Then
                        bFound = True
                        Exit For
                    End If
                Next Y
                If bFound = True Then Continue For

                For Y As Int32 = 0 To mlItemUB
                    If muItems(Y).iEnvirTypeID = ObjectType.eSolarSystem AndAlso muItems(Y).lEnvirID = muItems(X).lSystemID Then
                        bFound = True
                        Exit For
                    End If
                Next Y
                If bFound = False Then
                    lNeedToAddUB += 1
                    ReDim Preserve lNeedToAddSystems(lNeedToAddUB)
                    lNeedToAddSystems(lNeedToAddUB) = muItems(X).lSystemID
                End If
            Else
                lSystems += 1
            End If
            muItems(X).lChildrenItems = 0
        Next X
        'Now, add any systems that need to be added
        If lNeedToAddUB > -1 Then
            Dim lTmpUB As Int32 = mlItemUB + lNeedToAddUB + 1
            Dim oTmpItems(lTmpUB) As BudgetEnvir
            For X As Int32 = 0 To mlItemUB
                oTmpItems(X) = muItems(X)
            Next X

            For X As Int32 = 0 To lNeedToAddUB
                lSystems += 1
                With oTmpItems(X + mlItemUB + 1)
                    .AirCost = 0
                    .bHasCC = False
                    .bHasConflict = False
                    .blMiningBidExpense = 0
                    .blMiningBidIncome = 0
                    .DockedAirCost = 0
                    .DockedNonAirCost = 0
                    .ExcessStorage = 0
                    .FactoryCost = 0
                    .iEnvirTypeID = ObjectType.eSolarSystem
                    .lColonyID = -1
                    .lEnvirID = lNeedToAddSystems(X)
                    .lHWSupplyLineCost = 0
                    .lJumpToID = .lEnvirID
                    .lSystemID = .lEnvirID
                    .NonAirCost = 0
                    .OtherFacCost = 0
                    .PopUpkeep = 0
                    .ResearchCost = 0
                    .SpaceportCost = 0
                    .TaxIncome = 0
                    .TurretCost = 0
                    .UnemploymentCost = 0
                    .yPlanetaryControl = 0
                    .yTaxRate = 0
                    For y As Int32 = 0 To mlSrcListUB
                        If muSrcList(y).lSystemID = .lEnvirID AndAlso muSrcList(y).iEnvirTypeID = ObjectType.ePlanet Then
                            Try
                                .iTotalStarIncome += muSrcList(y).GetTotalRevenue
                                .iTotalStarExpense += muSrcList(y).GetTotalExpense
                            Catch
                            End Try
                        End If
                    Next
                End With
            Next X

            mlItemUB = lTmpUB
            muItems = oTmpItems
        End If

        'Ok, let's sort our items...
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To mlItemUB
            Dim lIdx As Int32 = -1
            Dim oItem As Budget.BudgetEnvir = muItems(X)

            'Ok, determine what we are sorting by...
            Select Case lSortType
                Case BudgetSort.Control
                    Dim yControl As Int64 = oItem.yPlanetaryControl
                    For Y As Int32 = 0 To lSortedUB
                        Dim ySorted As Int64 = muItems(lSorted(Y)).yPlanetaryControl
                        If (bDescending = False AndAlso ySorted > yControl) OrElse (bDescending = True AndAlso ySorted < yControl) Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                Case BudgetSort.ColonyName
                    Dim sCurrColonyName As String = GetCacheObjectValue(oItem.lColonyID, ObjectType.eColony)
                    Dim lTmpVal As Int32 = sCurrColonyName.LastIndexOf(" "c)
                    If lTmpVal <> -1 Then
                        Dim sTemp As String = sCurrColonyName.Substring(lTmpVal).Trim
                        sCurrColonyName = sCurrColonyName.Substring(0, lTmpVal)
                        sCurrColonyName &= GetRomanNumeralSortStr(sTemp)
                    End If
                    For Y As Int32 = 0 To lSortedUB
                        Dim sSortedColonyName As String = GetCacheObjectValue(muItems(lSorted(Y)).lColonyID, ObjectType.eColony)
                        lTmpVal = sSortedColonyName.LastIndexOf(" "c)
                        If lTmpVal <> -1 Then
                            Dim sTemp As String = sSortedColonyName.Substring(lTmpVal).Trim
                            sSortedColonyName = sSortedColonyName.Substring(0, lTmpVal)
                            sSortedColonyName &= GetRomanNumeralSortStr(sTemp)
                        End If
                        If (bDescending = False AndAlso sSortedColonyName.ToUpper > sCurrColonyName.ToUpper) OrElse (bDescending = True AndAlso sSortedColonyName.ToUpper < sCurrColonyName.ToUpper) Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                Case BudgetSort.EnvironmentName
                    Dim sCurrEnvirName As String = GetCacheObjectValue(oItem.lEnvirID, oItem.iEnvirTypeID)
                    Dim lTmpVal As Int32 = sCurrEnvirName.LastIndexOf(" "c)
                    If lTmpVal <> -1 Then
                        Dim sTemp As String = sCurrEnvirName.Substring(lTmpVal).Trim
                        sCurrEnvirName = sCurrEnvirName.Substring(0, lTmpVal)
                        sCurrEnvirName &= GetRomanNumeralSortStr(sTemp)
                    End If

                    For Y As Int32 = 0 To lSortedUB
                        Dim sSortedEnvirName As String = GetCacheObjectValue(muItems(lSorted(Y)).lEnvirID, muItems(lSorted(Y)).iEnvirTypeID)
                        lTmpVal = sSortedEnvirName.LastIndexOf(" "c)
                        If lTmpVal <> -1 Then
                            Dim sTemp As String = sSortedEnvirName.Substring(lTmpVal).Trim
                            sSortedEnvirName = sSortedEnvirName.Substring(0, lTmpVal)
                            sSortedEnvirName &= GetRomanNumeralSortStr(sTemp)
                        End If
                        If (bDescending = False AndAlso sSortedEnvirName.ToUpper > sCurrEnvirName.ToUpper) OrElse (bDescending = True AndAlso sSortedEnvirName.ToUpper < sCurrEnvirName.ToUpper) Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                Case BudgetSort.EnvironmentType
                    Dim iType As Int16 = oItem.iEnvirTypeID
                    For Y As Int32 = 0 To lSortedUB
                        Dim iSortedType As Int16 = muItems(lSorted(Y)).iEnvirTypeID
                        If (bDescending = False AndAlso iSortedType > iType) OrElse (bDescending = True AndAlso iSortedType < iType) Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                Case BudgetSort.Expense
                    Dim blExpense As Int64 = oItem.GetTotalExpense
                    'If frmBudget.bShowHWSupply = True Then blExpense += oItem.GetHomeworldSupplyLineCost(lIronCurtainPlanet)
                    For Y As Int32 = 0 To lSortedUB
                        Dim blSorted As Int64 = muItems(lSorted(Y)).GetTotalExpense
                        'If frmBudget.bShowHWSupply = True Then blSorted += muItems(lSorted(Y)).GetHomeworldSupplyLineCost(lIronCurtainPlanet)
                        If (bDescending = False AndAlso blSorted > blExpense) OrElse (bDescending = True AndAlso blSorted < blExpense) Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                Case BudgetSort.Revenue
                    Dim blRevenue As Int64 = oItem.GetTotalRevenue
                    For Y As Int32 = 0 To lSortedUB
                        Dim blSorted As Int64 = muItems(lSorted(Y)).GetTotalRevenue
                        If (bDescending = False AndAlso blSorted > blRevenue) OrElse (bDescending = True AndAlso blSorted < blRevenue) Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                Case BudgetSort.TaxRate
                    Dim yTax As Byte = oItem.yTaxRate
                    For Y As Int32 = 0 To lSortedUB
                        Dim ySorted As Byte = muItems(lSorted(Y)).yTaxRate
                        If (bDescending = False AndAlso ySorted > yTax) OrElse (bDescending = True AndAlso ySorted < yTax) Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
            End Select

            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X


        'Ok, need to do a final sort to put the sorted list into order by system...
        If lSystems > 0 Then
            Dim lSystemSorted(-1) As Int32
            For X As Int32 = 0 To mlItemUB
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To lSystemSorted.GetUpperBound(0)
                    If muItems(X).lSystemID = lSystemSorted(Y) Then
                        bFound = True
                        Exit For
                    End If
                Next Y
                If bFound = False Then
                    ReDim Preserve lSystemSorted(lSystemSorted.GetUpperBound(0) + 1)
                    lSystemSorted(lSystemSorted.GetUpperBound(0)) = muItems(X).lSystemID
                End If
            Next X

            'Now, sort them
            Dim lResult() As Int32 = Nothing
            Dim sResultVal() As String = Nothing
            Dim lResultUB As Int32 = -1
            For X As Int32 = 0 To lSystemSorted.GetUpperBound(0)
                Dim lIdx As Int32 = -1
                Dim sCurrent As String = GetCacheObjectValue(lSystemSorted(X), ObjectType.eSolarSystem)
                For Y As Int32 = 0 To lResultUB
                    If sResultVal(Y) > sCurrent Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y

                lResultUB += 1
                ReDim Preserve lResult(lResultUB)
                ReDim Preserve sResultVal(lResultUB)
                If lIdx = -1 Then
                    lResult(lResultUB) = lSystemSorted(X)
                    sResultVal(lResultUB) = sCurrent
                Else
                    For Y As Int32 = lResultUB To lIdx + 1 Step -1
                        lResult(Y) = lResult(Y - 1)
                        sResultVal(Y) = sResultVal(Y - 1)
                    Next Y
                    lResult(lIdx) = lSystemSorted(X)
                    sResultVal(lIdx) = sCurrent
                End If
            Next X


            Dim lTmpSorted() As Int32 = Nothing
            Dim lTmpSortedUB As Int32 = -1

            For X As Int32 = 0 To lResultUB
                'Ok, first, add the system item
                Dim lSystemEntryIdx As Int32 = -1
                For Y As Int32 = 0 To mlItemUB
                    If muItems(lSorted(Y)).lEnvirID = lResult(X) AndAlso muItems(lSorted(Y)).iEnvirTypeID = ObjectType.eSolarSystem Then
                        lTmpSortedUB += 1
                        ReDim Preserve lTmpSorted(lTmpSortedUB)
                        lTmpSorted(lTmpSortedUB) = lSorted(Y)
                        lSystemEntryIdx = lSorted(Y)
                        Exit For
                    End If
                Next Y

                'Now, just go through and add every item that has the system id match
                Dim lCnt As Int32 = 0
                For Y As Int32 = 0 To mlItemUB
                    If muItems(lSorted(Y)).lSystemID = lResult(X) Then
                        If muItems(lSorted(Y)).iEnvirTypeID <> ObjectType.eSolarSystem Then
                            lTmpSortedUB += 1
                            ReDim Preserve lTmpSorted(lTmpSortedUB)
                            lTmpSorted(lTmpSortedUB) = lSorted(Y)
                            lCnt += 1
                        End If
                    End If
                Next Y
                muItems(lSystemEntryIdx).lChildrenItems = lCnt
            Next X

            For X As Int32 = 0 To lTmpSortedUB
                lSorted(X) = lTmpSorted(X)
            Next X

        End If


        Dim uItems(mlItemUB) As BudgetEnvir
        For X As Int32 = 0 To mlItemUB
            uItems(X) = muItems(lSorted(X))
        Next X
        muItems = uItems

    End Sub

    Public Function GetTradeText() As String
        Dim oSB As New System.Text.StringBuilder
        oSB.Length = 0
        TotalTradeIncome = 0

        Dim MaxTradesAllowed As Int32 = 20

        If lTradePlayerID Is Nothing = False Then
            For X As Int32 = 0 To lTradePlayerID.GetUpperBound(0)

                If X = MaxTradesAllowed Then
                    oSB.AppendLine()
                    oSB.AppendLine("Not Received (max " & MaxTradesAllowed & ")")
                End If

                Dim sPlayer As String = GetCacheObjectValue(lTradePlayerID(X), ObjectType.ePlayer)
                Dim sValue As String = lTradeValue(X).ToString("#,##0")
                oSB.AppendLine((sPlayer & ":").PadRight(20, " "c) & sValue.PadLeft(9, " "c))

                If X < MaxTradesAllowed Then
                    TotalTradeIncome += lTradeValue(X)
                End If


            Next X
        End If
        If oSB.Length = 0 Then oSB.AppendLine("No Trade Income" & vbCrLf & " Allies and Guilds generate" & vbCrLf & " trade income")
        Return oSB.ToString
    End Function

End Class
