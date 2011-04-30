Option Strict On

Public MustInherit Class TechBuilderComputer
    Protected Shared moSB As System.Text.StringBuilder

    Public Shared bShowDebug As Boolean = False

    Public Enum elErrorReasons As Int32
        eNoError = 0
        ''' <summary>
        ''' ProdCost > 1% of Rescost or ResCost must be > 1% of ProdCost
        ''' </summary>
        ''' <remarks></remarks>
        eProdCostToResCostRatio = 1
        ''' <summary>
        ''' ProdTime must be 1/1000th of ResTime or ResTime must be 1/1000th of ProdTime
        ''' </summary>
        ''' <remarks></remarks>
        eProdTimeToResTimeRatio = 2
        ''' <summary>
        ''' Sum of Mineral AV must exceed Hull Size AV
        ''' </summary>
        ''' <remarks></remarks>
        eMineralsToHullSize = 4
        ''' <summary>
        ''' HullSizePoints + PowerPoints > 10% total bill
        ''' </summary>
        ''' <remarks></remarks>
        eHullPowerPointsLow = 8
        ''' <summary>
        ''' HullSizePoints + MineralPoints > 15% total bill
        ''' </summary>
        ''' <remarks></remarks>
        eHullMineralPointsLow = 16
        ''' <summary>
        ''' ResTimePoints + ProdTimePoints > 10% and less than 90% of total bill
        ''' </summary>
        ''' <remarks></remarks>
        eResTimeProdTimePoints = 32
        ''' <summary>
        ''' ResCostPoints + ProdCostPoints > 5% and less than 50% of total bill
        ''' </summary>
        ''' <remarks></remarks>
        eResCostProdCostPoints = 64
        ''' <summary>
        ''' HullAV must be > ColAV + EnlAV + OffAV
        ''' </summary>
        ''' <remarks></remarks>
        eHullToCrewRatio = 128
        ''' <summary>
        ''' EnlAV + OffAV must be > ColAV
        ''' </summary>
        ''' <remarks></remarks>
        eMilitaryToCitizen = 256
        ''' <summary>
        ''' EnlAV must be less than OffAV * 5
        ''' </summary>
        ''' <remarks></remarks>
        eLackOfOfficers = 512
        ''' <summary>
        ''' HullAV must be > CrewHullAV
        ''' </summary>
        ''' <remarks></remarks>
        eHullToCrewHullRatio = 1024
        ''' <summary>
        ''' any points paid cannot exceed 80% of the bill
        ''' </summary>
        ''' <remarks></remarks>
        e80PercMarkerExceed = 2048
        ''' <summary>
        ''' MineralAV 1 is less than 15% of all mineralsAV added
        ''' </summary>
        ''' <remarks></remarks>
        eMineral1LessThan15 = 4096
        ''' <summary>
        ''' MineralAV 2 is less than 15% of all mineralsAV added
        ''' </summary>
        ''' <remarks></remarks>
        eMineral2LessThan15 = 8192
        ''' <summary>
        ''' MineralAV 3 is less than 15% of all mineralsAV added
        ''' </summary>
        ''' <remarks></remarks>
        eMineral3LessThan15 = 16384
        ''' <summary>
        ''' MineralAV 4 is less than 15% of all mineralsAV added
        ''' </summary>
        ''' <remarks></remarks>
        eMineral4LessThan15 = 32768
        ''' <summary>
        ''' MineralAV 5 is less than 15% of all mineralsAV added
        ''' </summary>
        ''' <remarks></remarks>
        eMineral5LessThan15 = 65536
        ''' <summary>
        ''' MineralAV 6 is less than 15% of all mineralsAV added
        ''' </summary>
        ''' <remarks></remarks>
        eMineral6LessThan15 = 131072
        ''' <summary>
        ''' No locked value may be less than 1
        ''' </summary>
        ''' <remarks></remarks>
        eValueIsBelowZero = 262144
        ''' <summary>
        ''' Power Perc is Less than 1% (presently only for shields)
        ''' </summary>
        ''' <remarks></remarks>
        ePowerLessThan1Perc = 524288
        ''' <summary>
        ''' Hull Perc is Less than 1% (presently only for shields)
        ''' </summary>
        ''' <remarks></remarks>
        eHullLessThan1Perc = 1048576
        ''' <summary>
        ''' TotalBill cannot be less than total payment
        ''' </summary>
        ''' <remarks></remarks>
        ePaymentOverBill = 2097152
    End Enum
   
    Public Enum elPropCoeffLookup As Int32
        eHull = 0
        ePower = 1
        eColonist = 2
        eEnlisted = 3
        eOfficer = 4
        eResCost = 5
        eResTime = 6
        eProdCost = 7
        eProdTime = 8
        eMin1 = 9
        eMin2 = 10
        eMin3 = 11
        eMin4 = 12
        eMin5 = 13
        eMin6 = 14
    End Enum

    'Protected ml_POINT_PER_HULL As Int32
    'Protected ml_POINT_PER_POWER As Int32
    'Protected ml_POINT_PER_RES_TIME As Int32
    'Protected ml_POINT_PER_RES_COST As Int32
    'Protected ml_POINT_PER_PROD_TIME As Int32
    'Protected ml_POINT_PER_PROD_COST As Int32
    'Protected ml_POINT_PER_COLONIST As Int32
    'Protected ml_POINT_PER_ENLISTED As Int32
    'Protected ml_POINT_PER_OFFICER As Int32

    'Protected ml_POINT_PER_MIN1 As Int32
    'Protected ml_POINT_PER_MIN2 As Int32
    'Protected ml_POINT_PER_MIN3 As Int32
    'Protected ml_POINT_PER_MIN4 As Int32
    'Protected ml_POINT_PER_MIN5 As Int32
    'Protected ml_POINT_PER_MIN6 As Int32

    Public lHullTypeID As Int32 = -1

    'Protected MustOverride Function CalculateBaseHull() As Int32
    'Protected MustOverride Function CalculateBasePower() As Int32
    'Protected MustOverride Function CalculateBaseColonists() As Int32
    'Protected MustOverride Function CalculateBaseEnlisted() As Int32
    'Protected MustOverride Function CalculateBaseOfficers() As Int32
    'Protected MustOverride Function CalculateBaseResCost() As Int64
    'Protected MustOverride Function CalculateBaseResTime() As Int64
    'Protected MustOverride Function CalculateBaseProdCost() As Int64
    'Protected MustOverride Function CalculateBaseProdTime() As Int64
    'Protected MustOverride Function CalculateBaseMineralCosts(ByVal lMaxDA As Int32, ByVal lSumDAForMineral As Int32) As Int32
    Protected MustOverride Sub SetMinReqProps()
    'Protected MustOverride Sub LoadPointVals()
    'Protected MustOverride Function CalculateThresholdValue() As Decimal

    Protected MustOverride Function GetCoefficient(ByVal lLookup As elPropCoeffLookup) As Decimal
    Protected MustOverride Function GetTheBill() As Decimal

    Protected mbIgnoreValueChange As Boolean = False

    Public bImpossibleDesign As Boolean = True

    Public lMineral1ID As Int32 = -1
    Public lMineral2ID As Int32 = -1
    Public lMineral3ID As Int32 = -1
    Public lMineral4ID As Int32 = -1
    Public lMineral5ID As Int32 = -1
    Public lMineral6ID As Int32 = -1

    'MSC - stores DA values
    Private ml_Min1DA As Int32 = 255
    Private ml_Min2DA As Int32 = 255
    Private ml_Min3DA As Int32 = 255
    Private ml_Min4DA As Int32 = 255
    Private ml_Min5DA As Int32 = 255
    Private ml_Min6DA As Int32 = 255

    Public msMin1Name As String
    Public msMin2Name As String
    Public msMin3Name As String
    Public msMin4Name As String
    Public msMin5Name As String
    Public msMin6Name As String

    'These are here for pulling the costs out
    Public decTotalBill As Decimal
    Public decHullPayment As Decimal
    Public decPowerPayment As Decimal
    Public decResCostPayment As Decimal
    Public decResTimePayment As Decimal
    Public decProdCostPayment As Decimal
    Public decProdTimePayment As Decimal
    Public decColonistPayment As Decimal
    Public decEnlistedPayment As Decimal
    Public decOfficersPayment As Decimal
    Public decMineral1Payment As Decimal
    Public decMineral2Payment As Decimal
    Public decMineral3Payment As Decimal
    Public decMineral4Payment As Decimal
    Public decMineral5Payment As Decimal
    Public decMineral6Payment As Decimal
    Public lError As elErrorReasons = elErrorReasons.eNoError


    Protected Structure uMinPropReq
        Public lPropID As Int32
        Public lMinID As Int32
        Public lPropVal As Int32
    End Structure
    Protected muPropReqs() As uMinPropReq

    Public Shared Sub GetTypeValues(ByVal plHullType As Int32, ByRef pdecNormalizer As Decimal, ByRef plMaxGuns As Int32, ByRef plMaxDPS As Int32, ByRef plMaxHullSize As Int32, ByRef plHullAvail As Int32, ByRef plPower As Int32)
        'NOTE: HullAvail is not used anywhere - if it becomes used, we need to recalculate the values in here
        Select Case CType(plHullType, Base_Tech.eyHullType)
            Case EngineTech.eyHullType.BattleCruiser
                pdecNormalizer = 2777.77777777778D : plMaxGuns = 25 : plMaxDPS = 500 : plMaxHullSize = 250000 : plHullAvail = 200000 : plPower = 12500
            Case EngineTech.eyHullType.Battleship
                pdecNormalizer = 11111.1111111111D : plMaxGuns = 20 : plMaxDPS = 1900 : plMaxHullSize = 1000000 : plHullAvail = 800000 : plPower = 46900
            Case EngineTech.eyHullType.Corvette
                pdecNormalizer = 138.888888888D : plMaxGuns = 15 : plMaxDPS = 55 : plMaxHullSize = 12500 : plHullAvail = 10000 : plPower = 1515
            Case EngineTech.eyHullType.Cruiser
                pdecNormalizer = 1222.2222222222D : plMaxGuns = 20 : plMaxDPS = 250 : plMaxHullSize = 110000 : plHullAvail = 88000 : plPower = 6250
            Case EngineTech.eyHullType.Destroyer
                pdecNormalizer = 488.888888888889D : plMaxGuns = 20 : plMaxDPS = 125 : plMaxHullSize = 44000 : plHullAvail = 36000 : plPower = 3125
            Case EngineTech.eyHullType.Escort
                pdecNormalizer = 30D : plMaxGuns = 8 : plMaxDPS = 21 : plMaxHullSize = 2700 : plHullAvail = 2160 : plPower = 536
            Case EngineTech.eyHullType.Facility
                pdecNormalizer = 11111.1111111111D : plMaxGuns = 20 : plMaxDPS = 1900 : plMaxHullSize = 1000000 : plHullAvail = 800000 : plPower = 46900 'pdecNormalizer = 44444.4444444444D : plMaxGuns = 40
            Case EngineTech.eyHullType.Frigate
                pdecNormalizer = 311.111111111111D : plMaxGuns = 12 : plMaxDPS = 60 : plMaxHullSize = 28000 : plHullAvail = 22400 : plPower = 1389
            Case EngineTech.eyHullType.HeavyFighter
                pdecNormalizer = 3.33333333333333D : plMaxGuns = 4 : plMaxDPS = 14 : plMaxHullSize = 300 : plHullAvail = 240 : plPower = 347
            Case EngineTech.eyHullType.LightFighter
                pdecNormalizer = 1 : plMaxGuns = 2 : plMaxDPS = 10 : plMaxHullSize = 90 : plHullAvail = 72 : plPower = 250
            Case EngineTech.eyHullType.MediumFighter
                pdecNormalizer = 1.5555555555D : plMaxGuns = 3 : plMaxDPS = 12 : plMaxHullSize = 140 : plHullAvail = 108 : plPower = 303
            Case EngineTech.eyHullType.NavalBattleship
                pdecNormalizer = 2666.66666666667D : plMaxGuns = 20 : plMaxDPS = 500 : plMaxHullSize = 240000 : plHullAvail = 192000 : plPower = 12500
            Case EngineTech.eyHullType.NavalCarrier
                pdecNormalizer = 2777.77777777778D : plMaxGuns = 20 : plMaxDPS = 500 : plMaxHullSize = 250000 : plHullAvail = 200000 : plPower = 12500
            Case EngineTech.eyHullType.NavalCruiser
                pdecNormalizer = 1111.11111111111D : plMaxGuns = 15 : plMaxDPS = 150 : plMaxHullSize = 100000 : plHullAvail = 80000 : plPower = 6500
            Case EngineTech.eyHullType.NavalDestroyer
                pdecNormalizer = 500D : plMaxGuns = 10 : plMaxDPS = 200 : plMaxHullSize = 45000 : plHullAvail = 36000 : plPower = 3000
            Case EngineTech.eyHullType.NavalFrigate
                pdecNormalizer = 222.222222222222D : plMaxGuns = 10 : plMaxDPS = 75 : plMaxHullSize = 20000 : plHullAvail = 16000 : plPower = 1500
            Case EngineTech.eyHullType.NavalSub
                pdecNormalizer = 555.555555555556D : plMaxGuns = 2 : plMaxDPS = 1000 : plMaxHullSize = 50000 : plHullAvail = 40000 : plPower = 3500
            Case EngineTech.eyHullType.SmallVehicle
                pdecNormalizer = 1.33333333333333D : plMaxGuns = 5 : plMaxDPS = 10 : plMaxHullSize = 120 : plHullAvail = 96 : plPower = 250
            Case EngineTech.eyHullType.SpaceStation
                pdecNormalizer = 44444.4444444444D : plMaxGuns = 40 : plMaxDPS = 15000 : plMaxHullSize = 4000000 : plHullAvail = 3200000 : plPower = 375000
            Case EngineTech.eyHullType.Tank
                pdecNormalizer = 5.555555555555D : plMaxGuns = 12 : plMaxDPS = 40 : plMaxHullSize = 500 : plHullAvail = 400 : plPower = 1042
            Case EngineTech.eyHullType.Utility
                pdecNormalizer = 222.22222222222222D : plMaxGuns = 1 : plMaxDPS = 4 : plMaxHullSize = 36500 : plHullAvail = 30000 : plPower = 100
        End Select
    End Sub

    Protected Shared Sub AddToTestString(ByVal sLine As String)
        If moSB Is Nothing Then moSB = New System.Text.StringBuilder
        moSB.AppendLine(sLine)
    End Sub

    Private Function GetDADiff(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32
        If lMineralID = 157 Then Return 1

        Select Case lMineralIndex
            Case 0
                Return ml_Min1DA
            Case 1
                Return ml_Min2DA
            Case 2
                Return ml_Min3DA
            Case 3
                Return ml_Min4DA
            Case 4
                Return ml_Min5DA
            Case 5
                Return ml_Min6DA
        End Select
        Return 0        'RFI: 0 the correct answer here?
        'Dim oMineral As Mineral = Nothing
        'For X As Int32 = 0 To glMineralUB
        '    If glMineralIdx(X) = lMineralID Then
        '        oMineral = goMinerals(X)
        '        Exit For
        '    End If
        'Next X
        'If oMineral Is Nothing Then Return 0

        'If oMineral.ObjectID = 157 Then Return 1
        ''If oMineral.ObjectID = 9 Then Return 1

        'Dim lResult As Int32 = 0
        'If muPropReqs Is Nothing = False Then
        '    For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
        '        If muPropReqs(X).lMinID = lMineralIndex Then
        '            Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
        '            lResult += Math.Abs(muPropReqs(X).lPropVal - lA)
        '        End If
        '    Next X
        'End If
        ''        lResult = 0
        'If lResult < 1 Then Return 12
        'Return lResult * 23
    End Function
    Private Function GetClientDADiff(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32
        Dim oMineral As Mineral = Nothing
        For X As Int32 = 0 To glMineralUB
            If glMineralIdx(X) = lMineralID Then
                oMineral = goMinerals(X)
                Exit For
            End If
        Next X
        If oMineral Is Nothing Then Return 0

        If oMineral.ObjectID = 157 Then Return 0
        'If oMineral.ObjectID = 9 Then Return 0

        Dim lResult As Int32 = 0
        If muPropReqs Is Nothing = False Then
            For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
                If muPropReqs(X).lMinID = lMineralIndex Then
                    Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
                    lResult += Math.Abs(muPropReqs(X).lPropVal - lA)
                End If
            Next X
        End If
        'lResult = 0
        If lResult < 0 Then lResult = 0
        Return lResult '* 23
    End Function

    Public Sub MineralCBOExpanded(ByVal lMineralIndex As Int32, ByVal iTechTypeID As Int16)
        Try
            Dim ofrmMin As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
            If ofrmMin Is Nothing Then Return

            ofrmMin.ClearHighlights()
            SetMinReqProps()

            If muPropReqs Is Nothing = False AndAlso lHullTypeID > -1 Then
                For lProp As Int32 = 0 To muPropReqs.GetUpperBound(0)
                    With muPropReqs(lProp)
                        If .lMinID = lMineralIndex Then
                            Dim yDVal As Byte = CByte(.lPropVal)
                            Dim ySetting As Byte = GetPropValSettingNum(yDVal)
                            ofrmMin.HighlightProperty(.lPropID, ySetting, yDVal)
                        End If
                    End With
                Next lProp
            ElseIf muPropReqs Is Nothing = False AndAlso iTechTypeID = ObjectType.eArmorTech Then
                Dim lIdx As Int32 = -1
                Dim lVals(-1) As Int32
                Dim lPropID(-1) As Int32
                Dim lCnts(-1) As Int32

                For lProp As Int32 = 0 To muPropReqs.GetUpperBound(0)
                    With muPropReqs(lProp)
                        If .lMinID = lMineralIndex Then
                            Dim bFound As Boolean = False
                            For Y As Int32 = 0 To lIdx
                                If lPropID(Y) = .lPropID Then
                                    bFound = True
                                    lCnts(Y) += 1
                                    lVals(Y) += .lPropVal
                                    Exit For
                                End If
                            Next Y
                            If bFound = False Then
                                lIdx += 1
                                ReDim Preserve lVals(lIdx)
                                ReDim Preserve lPropID(lIdx)
                                ReDim Preserve lCnts(lIdx)
                                lVals(lIdx) = .lPropVal
                                lCnts(lIdx) = 1
                                lPropID(lIdx) = .lPropID
                            End If
                        End If
                    End With
                Next lProp

                'Now, go back through our list
                For X As Int32 = 0 To lIdx
                    If lCnts(X) > 0 Then
                        lVals(X) = lVals(X) \ lCnts(X)
                        Dim yDVal As Byte = CByte(lVals(X))
                        Dim ySetting As Byte = GetPropValSettingNum(yDVal)
                        ofrmMin.HighlightProperty(lPropID(X), ySetting, yDVal)
                    End If
                Next X
            Else
                goUILib.AddNotification("Select a hull type for this technology to see the required properties.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If

            ofrmMin.ForceHighlightFirstItem()
        Catch
        End Try
    End Sub

    Private Shared Function GetPropValSettingNum(ByVal yPropVal As Byte) As Byte
        '0 = not highlighted, 1 = white, 2 = green, 3 = red, 4 = yellow
        If yPropVal < 4 Then Return 3
        If yPropVal < 7 Then Return 4
        Return 2
    End Function

    Public Function GetCoeffValue(ByVal lLookup As elPropCoeffLookup) As Decimal

        'Dim lBaseMin1 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral1ID, 0), GetClientDASum(lMineral1ID, 0))
        'AddToTestString("BaseMin1: " & lBaseMin1.ToString)
        'Dim lBaseMin2 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral2ID, 1), GetClientDASum(lMineral2ID, 1))
        'AddToTestString("BaseMin2: " & lBaseMin2.ToString)
        'Dim lBaseMin3 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral3ID, 2), GetClientDASum(lMineral3ID, 2))
        'AddToTestString("BaseMin3: " & lBaseMin3.ToString)
        'Dim lBaseMin4 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral4ID, 3), GetClientDASum(lMineral4ID, 3))
        'AddToTestString("BaseMin4: " & lBaseMin4.ToString)
        'Dim lBaseMin5 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral5ID, 4), GetClientDASum(lMineral5ID, 4))
        'AddToTestString("BaseMin5: " & lBaseMin5.ToString)
        'Dim lBaseMin6 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral6ID, 5), GetClientDASum(lMineral6ID, 5))
        'AddToTestString("BaseMin6: " & lBaseMin6.ToString)
        Dim decTemp As Decimal = GetCoefficient(lLookup)

        Select Case lLookup
            Case elPropCoeffLookup.eMin1
                Dim decDADiff As Decimal = GetDADiff(lMineral1ID, 0)
                decTemp -= decDADiff / 1000D
            Case elPropCoeffLookup.eMin2
                Dim decDADiff As Decimal = GetDADiff(lMineral2ID, 1)
                decTemp -= decDADiff / 1000D
            Case elPropCoeffLookup.eMin3
                Dim decDADiff As Decimal = GetDADiff(lMineral3ID, 2)
                decTemp -= decDADiff / 1000D
            Case elPropCoeffLookup.eMin4
                Dim decDADiff As Decimal = GetDADiff(lMineral4ID, 3)
                decTemp -= decDADiff / 1000D
            Case elPropCoeffLookup.eMin5
                Dim decDADiff As Decimal = GetDADiff(lMineral5ID, 4)
                decTemp -= decDADiff / 1000D
            Case elPropCoeffLookup.eMin6
                Dim decDADiff As Decimal = GetDADiff(lMineral6ID, 5)
                decTemp -= decDADiff / 1000D
            Case elPropCoeffLookup.eProdCost
                'decTemp -= 0.01D
                'If decTemp < 0 Then decTemp = 0.01D
        End Select

        Return decTemp
    End Function

    'Public Sub DoBuilderCostPreconfigure(ByRef lblDesignFlaw As UILabel, ByRef frmBuildCost As frmTechBuilderCost, ByRef tpMin1 As ctlTechProp, ByRef tpMin2 As ctlTechProp, ByRef tpMin3 As ctlTechProp, ByRef tpMin4 As ctlTechProp, ByRef tpMin5 As ctlTechProp, ByRef tpMin6 As ctlTechProp, ByVal iTechTypeID As Int16, ByVal decMaxThreshold As Decimal)
    '    bImpossibleDesign = False

    '    Try
    '        SetMinReqProps()

    '        decTotalBill = GetTheBill()


    '        'ok, now we have the bill... let's get the locked payment
    '        Dim lLockedHull As Int32 = -1
    '        Dim lLockedPower As Int32 = -1
    '        Dim blLockedResCost As Int64 = -1
    '        Dim blLockedResTime As Int64 = -1
    '        Dim blLockedProdCost As Int64 = -1
    '        Dim blLockedProdTime As Int64 = -1
    '        Dim lLockedColonists As Int32 = -1
    '        Dim lLockedEnlisted As Int32 = -1
    '        Dim lLockedOfficers As Int32 = -1
    '        'frmBuildCost.GetLockedValues(lLockedHull, lLockedPower, blLockedResCost, blLockedResTime, blLockedProdCost, blLockedProdTime, lLockedColonists, lLockedEnlisted, lLockedOfficers)

    '        Dim bColLockedAtStart As Boolean = lLockedColonists <> -1
    '        Dim bEnlLockedAtStart As Boolean = lLockedEnlisted <> -1
    '        Dim bOffLockedAtStart As Boolean = lLockedOfficers <> -1

    '        'Also need to get the locked minerals
    '        Dim lLockedMin1 As Int32 = -1
    '        Dim lLockedMin2 As Int32 = -1
    '        Dim lLockedMin3 As Int32 = -1
    '        Dim lLockedMin4 As Int32 = -1
    '        Dim lLockedMin5 As Int32 = -1
    '        Dim lLockedMin6 As Int32 = -1
    '        Dim bMin1Locked, bMin2Locked, bMin3Locked, bMin4Locked, bMin5Locked, bMin6Locked As Boolean
    '        Dim lMinCnt As Int32 = 0
    '        Dim lTotalDAOff As Int32 = 0
    '        If tpMin1 Is Nothing = False Then
    '            bMin1Locked = False
    '            'If tpMin1.PropertyLocked = True Then
    '            '    lLockedMin1 = CInt(tpMin1.PropertyValue)
    '            '    bMin1Locked = True
    '            'Else : 
    '            lMinCnt += 1
    '            'End If
    '            lTotalDAOff += GetClientDADiff(lMineral1ID, 0)
    '        End If
    '        If tpMin2 Is Nothing = False Then
    '            bMin2Locked = False
    '            'If tpMin2.PropertyLocked = True Then
    '            '    lLockedMin2 = CInt(tpMin2.PropertyValue)
    '            '    bMin2Locked = True
    '            'Else : 
    '            lMinCnt += 1
    '            'End If
    '            lTotalDAOff += GetClientDADiff(lMineral2ID, 1)
    '        End If
    '        If tpMin3 Is Nothing = False Then
    '            bMin3Locked = False
    '            'If tpMin3.PropertyLocked = True Then
    '            '    lLockedMin3 = CInt(tpMin3.PropertyValue)
    '            '    bMin3Locked = True
    '            'Else :
    '            lMinCnt += 1
    '            'End If
    '            lTotalDAOff += GetClientDADiff(lMineral3ID, 2)
    '        End If
    '        If tpMin4 Is Nothing = False Then
    '            bMin4Locked = False
    '            'If tpMin4.PropertyLocked = True Then
    '            '    lLockedMin4 = CInt(tpMin4.PropertyValue)
    '            '    bMin4Locked = True
    '            'Else : 
    '            lMinCnt += 1
    '            'End If
    '            lTotalDAOff += GetClientDADiff(lMineral4ID, 3)
    '        End If
    '        If tpMin5 Is Nothing = False Then
    '            bMin5Locked = False
    '            'If tpMin5.PropertyLocked = True Then
    '            '    lLockedMin5 = CInt(tpMin5.PropertyValue)
    '            '    bMin5Locked = True
    '            'Else : 
    '            lMinCnt += 1
    '            'End If
    '            lTotalDAOff += GetClientDADiff(lMineral5ID, 4)
    '        End If
    '        If tpMin6 Is Nothing = False Then
    '            bMin6Locked = False
    '            'If tpMin6.PropertyLocked = True Then
    '            '    lLockedMin6 = CInt(tpMin6.PropertyValue)
    '            '    bMin6Locked = True
    '            'Else : 
    '            lMinCnt += 1
    '            'End If
    '            lTotalDAOff += GetClientDADiff(lMineral6ID, 5)
    '        End If
    '        decTotalBill *= Math.Max(1, lTotalDAOff)

    '        Dim decMax As Decimal = Math.Max(Math.Max(Math.Max(Math.Max(Math.Max(GetCoeffValue(elPropCoeffLookup.eMin1), GetCoeffValue(elPropCoeffLookup.eMin2)), GetCoeffValue(elPropCoeffLookup.eMin3)), GetCoeffValue(elPropCoeffLookup.eMin4)), GetCoeffValue(elPropCoeffLookup.eMin5)), GetCoeffValue(elPropCoeffLookup.eMin6))

    '        If bMin1Locked = False AndAlso tpMin1 Is Nothing = False Then lLockedMin1 = CInt(Math.Pow((decTotalBill * 0.15D) / lMinCnt, 1D / decMax))
    '        If bMin2Locked = False AndAlso tpMin2 Is Nothing = False Then lLockedMin2 = lLockedMin1
    '        If bMin3Locked = False AndAlso tpMin3 Is Nothing = False Then lLockedMin3 = lLockedMin1
    '        If bMin4Locked = False AndAlso tpMin4 Is Nothing = False Then lLockedMin4 = lLockedMin1
    '        If bMin5Locked = False AndAlso tpMin5 Is Nothing = False Then lLockedMin5 = lLockedMin1
    '        If bMin6Locked = False AndAlso tpMin6 Is Nothing = False Then lLockedMin6 = lLockedMin1

    '        If lLockedHull = -1 Then lLockedHull = CInt(Math.Min(Math.Pow((decTotalBill * 0.5D), (1 / Math.Max(GetCoeffValue(elPropCoeffLookup.eHull), GetCoeffValue(elPropCoeffLookup.ePower)))), (lLockedMin1 + lLockedMin2 + lLockedMin3 + lLockedMin4 + lLockedMin5 + lLockedMin6 - 1)))
    '        If lLockedPower = -1 AndAlso iTechTypeID <> ObjectType.eEngineTech Then lLockedPower = CInt(lLockedHull / 1.5)

    '        Dim lCrewHull As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence)
    '        If lLockedOfficers = -1 Then lLockedOfficers = CInt(Math.Min(lLockedHull / (10 * lCrewHull), Math.Pow((decTotalBill * 0.001D), 1 / GetCoeffValue(elPropCoeffLookup.eOfficer))))
    '        If lLockedEnlisted = -1 Then lLockedEnlisted = (lLockedOfficers * 5) - 1
    '        If lLockedColonists = -1 Then lLockedColonists = CInt(Math.Min(lLockedHull / (4 * lCrewHull), lLockedOfficers + lLockedEnlisted))

    '        Dim decPaymentSoFar As Decimal = CDec(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
    '        If iTechTypeID <> ObjectType.eEngineTech Then decPaymentSoFar += CDec(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
    '        Dim decCoeff As Decimal = GetCoeffValue(elPropCoeffLookup.eColonist)
    '        If decCoeff > 0 Then decPaymentSoFar += CDec(Math.Pow(lLockedColonists, decCoeff))
    '        decCoeff = GetCoeffValue(elPropCoeffLookup.eEnlisted)
    '        If decCoeff > 0 Then decPaymentSoFar += CDec(Math.Pow(lLockedEnlisted, decCoeff))
    '        decCoeff = GetCoeffValue(elPropCoeffLookup.eOfficer)
    '        If decCoeff > 0 Then decPaymentSoFar += CDec(Math.Pow(lLockedOfficers, decCoeff))
    '        If tpMin1 Is Nothing = False Then decPaymentSoFar += CDec(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
    '        If tpMin2 Is Nothing = False Then decPaymentSoFar += CDec(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
    '        If tpMin3 Is Nothing = False Then decPaymentSoFar += CDec(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
    '        If tpMin4 Is Nothing = False Then decPaymentSoFar += CDec(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
    '        If tpMin5 Is Nothing = False Then decPaymentSoFar += CDec(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
    '        If tpMin6 Is Nothing = False Then decPaymentSoFar += CDec(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))

    '        Dim decDiff As Decimal = decTotalBill - decPaymentSoFar
    '        If blLockedProdCost = -1 Then blLockedProdCost = CLng(Math.Pow((decDiff / 2.5D), 1 / GetCoeffValue(elPropCoeffLookup.eProdCost)))
    '        decPaymentSoFar += CDec(Math.Pow(blLockedProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
    '        If blLockedResCost = -1 Then blLockedResCost = CLng(blLockedProdCost * 99)
    '        decPaymentSoFar += CDec(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
    '        decDiff = decTotalBill - decPaymentSoFar
    '        If blLockedProdTime = -1 Then blLockedProdTime = CLng(Math.Ceiling(Math.Pow((decDiff * 0.99D), 1 / GetCoeffValue(elPropCoeffLookup.eProdTime))))
    '        If blLockedResTime = -1 Then blLockedResTime = CLng((blLockedProdTime / 528) * 92000) - 1

    '        frmBuildCost.SetPreConfigValues(lLockedHull, lLockedPower, blLockedResCost, blLockedResTime, blLockedProdCost, -1, lLockedColonists, lLockedEnlisted, lLockedOfficers)
    '        If tpMin1 Is Nothing = False Then
    '            tpMin1.PropertyLocked = True
    '            tpMin1.PropertyValue = lLockedMin1
    '        End If
    '        If tpMin2 Is Nothing = False Then
    '            tpMin2.PropertyLocked = True
    '            tpMin2.PropertyValue = lLockedMin2
    '        End If
    '        If tpMin3 Is Nothing = False Then
    '            tpMin3.PropertyLocked = True
    '            tpMin3.PropertyValue = lLockedMin3
    '        End If
    '        If tpMin4 Is Nothing = False Then
    '            tpMin4.PropertyLocked = True
    '            tpMin4.PropertyValue = lLockedMin4
    '        End If
    '        If tpMin5 Is Nothing = False Then
    '            tpMin5.PropertyLocked = True
    '            tpMin5.PropertyValue = lLockedMin5
    '        End If
    '        If tpMin6 Is Nothing = False Then
    '            tpMin6.PropertyLocked = True
    '            tpMin6.PropertyValue = lLockedMin6
    '        End If

    '        BuilderCostValueChange(lblDesignFlaw, frmBuildCost, tpMin1, tpMin2, tpMin3, tpMin4, tpMin5, tpMin6, iTechTypeID, decMaxThreshold)

    '    Catch
    '        lblDesignFlaw.Visible = True
    '        lblDesignFlaw.Caption = "This design is impossible to comprehend."
    '        bImpossibleDesign = True
    '    End Try

    '    mbIgnoreValueChange = False

    'End Sub

    Public Sub DoBuilderCostPreconfigure(ByRef lblDesignFlaw As UILabel, ByRef frmBuildCost As frmTechBuilderCost, ByRef tpMin1 As ctlTechProp, ByRef tpMin2 As ctlTechProp, ByRef tpMin3 As ctlTechProp, ByRef tpMin4 As ctlTechProp, ByRef tpMin5 As ctlTechProp, ByRef tpMin6 As ctlTechProp, ByVal iTechTypeID As Int16, ByVal decMaxThreshold As Decimal, ByVal bDoNotLockCrew As Boolean)
        bImpossibleDesign = False

        Try
            SetMinReqProps()

            decTotalBill = GetTheBill()


            'ok, now we have the bill... let's get the locked payment
            Dim lLockedHull As Int32 = -1
            Dim lLockedPower As Int32 = -1
            Dim blLockedResCost As Int64 = -1
            Dim blLockedResTime As Int64 = -1
            Dim blLockedProdCost As Int64 = -1
            Dim blLockedProdTime As Int64 = -1
            Dim lLockedColonists As Int32 = -1
            Dim lLockedEnlisted As Int32 = -1
            Dim lLockedOfficers As Int32 = -1
            'frmBuildCost.GetLockedValues(lLockedHull, lLockedPower, blLockedResCost, blLockedResTime, blLockedProdCost, blLockedProdTime, lLockedColonists, lLockedEnlisted, lLockedOfficers)

            Dim bColLockedAtStart As Boolean = lLockedColonists <> -1
            Dim bEnlLockedAtStart As Boolean = lLockedEnlisted <> -1
            Dim bOffLockedAtStart As Boolean = lLockedOfficers <> -1

            'Also need to get the locked minerals
            Dim lLockedMin1 As Int32 = -1
            Dim lLockedMin2 As Int32 = -1
            Dim lLockedMin3 As Int32 = -1
            Dim lLockedMin4 As Int32 = -1
            Dim lLockedMin5 As Int32 = -1
            Dim lLockedMin6 As Int32 = -1
            Dim bMin1Locked, bMin2Locked, bMin3Locked, bMin4Locked, bMin5Locked, bMin6Locked As Boolean
            Dim lMinCnt As Int32 = 0
            Dim lTotalDAOff As Int32 = 0
            If tpMin1 Is Nothing = False Then
                bMin1Locked = False
                'If tpMin1.PropertyLocked = True Then
                '    lLockedMin1 = CInt(tpMin1.PropertyValue)
                '    bMin1Locked = True
                'Else : 
                lMinCnt += 1
                'End If
                lTotalDAOff += GetClientDADiff(lMineral1ID, 0)
            End If
            If tpMin2 Is Nothing = False Then
                bMin2Locked = False
                'If tpMin2.PropertyLocked = True Then
                '    lLockedMin2 = CInt(tpMin2.PropertyValue)
                '    bMin2Locked = True
                'Else : 
                lMinCnt += 1
                'End If
                lTotalDAOff += GetClientDADiff(lMineral2ID, 1)
            End If
            If tpMin3 Is Nothing = False Then
                bMin3Locked = False
                'If tpMin3.PropertyLocked = True Then
                '    lLockedMin3 = CInt(tpMin3.PropertyValue)
                '    bMin3Locked = True
                'Else :
                lMinCnt += 1
                'End If
                lTotalDAOff += GetClientDADiff(lMineral3ID, 2)
            End If
            If tpMin4 Is Nothing = False Then
                bMin4Locked = False
                'If tpMin4.PropertyLocked = True Then
                '    lLockedMin4 = CInt(tpMin4.PropertyValue)
                '    bMin4Locked = True
                'Else : 
                lMinCnt += 1
                'End If
                lTotalDAOff += GetClientDADiff(lMineral4ID, 3)
            End If
            If tpMin5 Is Nothing = False Then
                bMin5Locked = False
                'If tpMin5.PropertyLocked = True Then
                '    lLockedMin5 = CInt(tpMin5.PropertyValue)
                '    bMin5Locked = True
                'Else : 
                lMinCnt += 1
                'End If
                lTotalDAOff += GetClientDADiff(lMineral5ID, 4)
            End If
            If tpMin6 Is Nothing = False Then
                bMin6Locked = False
                'If tpMin6.PropertyLocked = True Then
                '    lLockedMin6 = CInt(tpMin6.PropertyValue)
                '    bMin6Locked = True
                'Else : 
                lMinCnt += 1
                'End If
                lTotalDAOff += GetClientDADiff(lMineral6ID, 5)
            End If
            'decTotalBill *= Math.Max(1, lTotalDAOff)

            Dim decMax As Decimal = Math.Max(Math.Max(Math.Max(Math.Max(Math.Max(GetCoeffValue(elPropCoeffLookup.eMin1), GetCoeffValue(elPropCoeffLookup.eMin2)), GetCoeffValue(elPropCoeffLookup.eMin3)), GetCoeffValue(elPropCoeffLookup.eMin4)), GetCoeffValue(elPropCoeffLookup.eMin5)), GetCoeffValue(elPropCoeffLookup.eMin6))

            If bMin1Locked = False AndAlso tpMin1 Is Nothing = False Then lLockedMin1 = CInt(Math.Pow((decTotalBill * 0.15D) / lMinCnt, 1D / decMax))
            If bMin2Locked = False AndAlso tpMin2 Is Nothing = False Then lLockedMin2 = lLockedMin1
            If bMin3Locked = False AndAlso tpMin3 Is Nothing = False Then lLockedMin3 = lLockedMin1
            If bMin4Locked = False AndAlso tpMin4 Is Nothing = False Then lLockedMin4 = lLockedMin1
            If bMin5Locked = False AndAlso tpMin5 Is Nothing = False Then lLockedMin5 = lLockedMin1
            If bMin6Locked = False AndAlso tpMin6 Is Nothing = False Then lLockedMin6 = lLockedMin1

            lLockedOfficers = 1
            lLockedEnlisted = 1
            lLockedColonists = 1

            If bDoNotLockCrew = True Then
                lLockedColonists = -1
                lLockedEnlisted = -1
                lLockedOfficers = -1
            End If

            Dim decPaymentSoFar As Decimal = 0D 'CDec(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
            Dim decCoeff As Decimal = GetCoeffValue(elPropCoeffLookup.eColonist)
            If decCoeff > 0 AndAlso bDoNotLockCrew = False Then decPaymentSoFar += CDec(Math.Pow(lLockedColonists, decCoeff))
            decCoeff = GetCoeffValue(elPropCoeffLookup.eEnlisted)
            If decCoeff > 0 AndAlso bDoNotLockCrew = False Then decPaymentSoFar += CDec(Math.Pow(lLockedEnlisted, decCoeff))
            decCoeff = GetCoeffValue(elPropCoeffLookup.eOfficer)
            If decCoeff > 0 AndAlso bDoNotLockCrew = False Then decPaymentSoFar += CDec(Math.Pow(lLockedOfficers, decCoeff))

            'now, prod cost to 5%
            If blLockedProdCost = -1 Then
                decCoeff = GetCoeffValue(elPropCoeffLookup.eProdCost)
                blLockedProdCost = CLng(Math.Pow(decTotalBill * 0.05D, (1D / decCoeff)))
            End If
            blLockedResCost = CLng(Math.Ceiling(blLockedProdCost * 0.01D))

            'hull and power to 10%
            If iTechTypeID = ObjectType.eEngineTech Then
                lLockedHull = CInt(Math.Pow(decTotalBill * 0.1D, (1 / GetCoeffValue(elPropCoeffLookup.eHull)))) + 1
                lLockedPower = 0
            Else
                lLockedHull = CInt(Math.Pow(decTotalBill * 0.09D, (1 / GetCoeffValue(elPropCoeffLookup.eHull)))) + 1
                lLockedPower = CInt(Math.Pow(decTotalBill * 0.01D, (1 / GetCoeffValue(elPropCoeffLookup.ePower)))) + 1
            End If

            'now, res time as 1% of prodtime... problem is, we need to determine what points are left, and then
            '  distribute those among the minerals and the prod time to determine an estimate of the prod time
            '  then, set the res time to 1% of that value
            decCoeff = GetCoeffValue(elPropCoeffLookup.eProdTime)
            Dim decTotalCoeff As Decimal = 1D / decCoeff
            If tpMin1 Is Nothing = False Then decTotalCoeff += 1D / GetCoeffValue(elPropCoeffLookup.eMin1)
            If tpMin2 Is Nothing = False Then decTotalCoeff += 1D / GetCoeffValue(elPropCoeffLookup.eMin2)
            If tpMin3 Is Nothing = False Then decTotalCoeff += 1D / GetCoeffValue(elPropCoeffLookup.eMin3)
            If tpMin4 Is Nothing = False Then decTotalCoeff += 1D / GetCoeffValue(elPropCoeffLookup.eMin4)
            If tpMin5 Is Nothing = False Then decTotalCoeff += 1D / GetCoeffValue(elPropCoeffLookup.eMin5)
            If tpMin6 Is Nothing = False Then decTotalCoeff += 1D / GetCoeffValue(elPropCoeffLookup.eMin6)
            decTotalCoeff = decCoeff / decTotalCoeff


            Dim blTempProdTime As Int64 = CLng(Math.Pow(decTotalBill * decTotalCoeff, (1 / decCoeff)))
            blLockedResTime = (CLng(blTempProdTime * 0.01D) \ 92) * 528

            frmBuildCost.SetPreConfigValues(lLockedHull, lLockedPower, blLockedResCost, blLockedResTime, blLockedProdCost, -1, lLockedColonists, lLockedEnlisted, lLockedOfficers)

            BuilderCostValueChange(lblDesignFlaw, frmBuildCost, tpMin1, tpMin2, tpMin3, tpMin4, tpMin5, tpMin6, iTechTypeID, decMaxThreshold)

        Catch
            lblDesignFlaw.Visible = True
            lblDesignFlaw.Caption = "This design is impossible to comprehend."
            bImpossibleDesign = True
        End Try

        mbIgnoreValueChange = False

    End Sub

    Public Sub BuilderCostValueChange(ByRef lblDesignFlaw As UILabel, ByRef frmBuildCost As frmTechBuilderCost, ByRef tpMin1 As ctlTechProp, ByRef tpMin2 As ctlTechProp, ByRef tpMin3 As ctlTechProp, ByRef tpMin4 As ctlTechProp, ByRef tpMin5 As ctlTechProp, ByRef tpMin6 As ctlTechProp, ByVal iTechTypeID As Int16, ByVal decMaxThreshold As Decimal)
        bImpossibleDesign = False

        Try
            SetMinReqProps()
            If moSB Is Nothing = False Then moSB.Length = 0

            decTotalBill = GetTheBill()
            AddToTestString("Bill: " & decTotalBill.ToString("#,###"))

            'ok, now we have the bill... let's get the locked payment
            Dim lLockedHull As Int32 = -1
            Dim lLockedPower As Int32 = -1
            Dim blLockedResCost As Int64 = -1
            Dim blLockedResTime As Int64 = -1
            Dim blLockedProdCost As Int64 = -1
            Dim blLockedProdTime As Int64 = -1
            Dim lLockedColonists As Int32 = -1
            Dim lLockedEnlisted As Int32 = -1
            Dim lLockedOfficers As Int32 = -1
            frmBuildCost.GetLockedValues(lLockedHull, lLockedPower, blLockedResCost, blLockedResTime, blLockedProdCost, blLockedProdTime, lLockedColonists, lLockedEnlisted, lLockedOfficers)

            Dim bColLockedAtStart As Boolean = lLockedColonists <> -1
            Dim bEnlLockedAtStart As Boolean = lLockedEnlisted <> -1
            Dim bOffLockedAtStart As Boolean = lLockedOfficers <> -1

            'Also need to get the locked minerals
            Dim lLockedMin1 As Int32 = -1
            Dim lLockedMin2 As Int32 = -1
            Dim lLockedMin3 As Int32 = -1
            Dim lLockedMin4 As Int32 = -1
            Dim lLockedMin5 As Int32 = -1
            Dim lLockedMin6 As Int32 = -1
            Dim bMin1Locked, bMin2Locked, bMin3Locked, bMin4Locked, bMin5Locked, bMin6Locked As Boolean
            Dim lTotalDAOff As Int32 = 0
            If tpMin1 Is Nothing = False Then
                bMin1Locked = False
                If tpMin1.PropertyLocked = True Then
                    lLockedMin1 = CInt(tpMin1.PropertyValue)
                    bMin1Locked = True
                End If
                lTotalDAOff += GetClientDADiff(lMineral1ID, 0)
            End If
            If tpMin2 Is Nothing = False Then
                bMin2Locked = False
                If tpMin2.PropertyLocked = True Then
                    lLockedMin2 = CInt(tpMin2.PropertyValue)
                    bMin2Locked = True
                End If
                lTotalDAOff += GetClientDADiff(lMineral2ID, 1)
            End If
            If tpMin3 Is Nothing = False Then
                bMin3Locked = False
                If tpMin3.PropertyLocked = True Then
                    lLockedMin3 = CInt(tpMin3.PropertyValue)
                    bMin3Locked = True
                End If
                lTotalDAOff += GetClientDADiff(lMineral3ID, 2)
            End If
            If tpMin4 Is Nothing = False Then
                bMin4Locked = False
                If tpMin4.PropertyLocked = True Then
                    lLockedMin4 = CInt(tpMin4.PropertyValue)
                    bMin4Locked = True
                End If
                lTotalDAOff += GetClientDADiff(lMineral4ID, 3)
            End If
            If tpMin5 Is Nothing = False Then
                bMin5Locked = False
                If tpMin5.PropertyLocked = True Then
                    lLockedMin5 = CInt(tpMin5.PropertyValue)
                    bMin5Locked = True
                End If
                lTotalDAOff += GetClientDADiff(lMineral5ID, 4)
            End If
            If tpMin6 Is Nothing = False Then
                bMin6Locked = False
                If tpMin6.PropertyLocked = True Then
                    lLockedMin6 = CInt(tpMin6.PropertyValue)
                    bMin6Locked = True
                End If
                lTotalDAOff += GetClientDADiff(lMineral6ID, 5)
            End If
            'decTotalBill *= Math.Max(lTotalDAOff, 1)

            Dim decPayment As Decimal = 0
            Dim lDistributable As Int32 = 0
            Dim decDistVals(elPropCoeffLookup.eMin6) As Decimal
            For X As Int32 = 0 To elPropCoeffLookup.eMin6
                decDistVals(X) = 0D
                'AddToTestString(X & ": " & GetCoeffValue(CType(X, elPropCoeffLookup)))
            Next X

            If lLockedHull <> -1 Then
                decHullPayment = CDec(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                decPayment += decHullPayment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eHull) = GetCoeffValue(elPropCoeffLookup.eHull)
            End If
            If lLockedPower <> -1 AndAlso GetCoeffValue(elPropCoeffLookup.ePower) <> 0 Then
                decPowerPayment = CDec(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                decPayment += decPowerPayment
            ElseIf GetCoeffValue(elPropCoeffLookup.ePower) <> 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.ePower) = GetCoeffValue(elPropCoeffLookup.ePower)
            Else
                lLockedPower = 0
            End If
            If blLockedResCost <> -1 Then
                decResCostPayment = CDec(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
                decPayment += decResCostPayment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eResCost) = GetCoeffValue(elPropCoeffLookup.eResCost)
            End If
            If blLockedResTime <> -1 Then
                decResTimePayment = CDec(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
                decPayment += decResTimePayment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eResTime) = GetCoeffValue(elPropCoeffLookup.eResTime)
            End If
            If blLockedProdCost <> -1 Then
                decProdCostPayment = CDec(Math.Pow(blLockedProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
                decPayment += decProdCostPayment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eProdCost) = GetCoeffValue(elPropCoeffLookup.eProdCost)
            End If
            If blLockedProdTime <> -1 Then
                decProdTimePayment = CDec(Math.Pow(blLockedProdTime, GetCoeffValue(elPropCoeffLookup.eProdTime)))
                decPayment += decProdTimePayment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eProdTime) = GetCoeffValue(elPropCoeffLookup.eProdTime)
            End If
            If GetCoeffValue(elPropCoeffLookup.eColonist) <> 0 Then
                If lLockedColonists <> -1 Then
                    decColonistPayment = CDec(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                    decPayment += decColonistPayment
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eColonist) = GetCoeffValue(elPropCoeffLookup.eColonist)
                End If
            Else : lLockedColonists = 0
            End If

            If GetCoeffValue(elPropCoeffLookup.eEnlisted) <> 0 Then
                If lLockedEnlisted <> -1 Then
                    decEnlistedPayment = CDec(Math.Pow(lLockedEnlisted, GetCoeffValue(elPropCoeffLookup.eEnlisted)))
                    decPayment += decEnlistedPayment
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eEnlisted) = GetCoeffValue(elPropCoeffLookup.eEnlisted)
                End If
            Else : lLockedEnlisted = 0
            End If

            If GetCoeffValue(elPropCoeffLookup.eOfficer) <> 0 Then
                If lLockedOfficers <> -1 Then
                    decOfficersPayment = CDec(Math.Pow(lLockedOfficers, GetCoeffValue(elPropCoeffLookup.eOfficer)))
                    decPayment += decOfficersPayment
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eOfficer) = GetCoeffValue(elPropCoeffLookup.eOfficer)
                End If
            Else : lLockedOfficers = 0
            End If

            If bMin1Locked = True Then
                decMineral1Payment = CDec(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
                decPayment += decMineral1Payment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin1) = GetCoeffValue(elPropCoeffLookup.eMin1)
            End If
            If bMin2Locked = True Then
                decMineral2Payment = CDec(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
                decPayment += decMineral2Payment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin2) = GetCoeffValue(elPropCoeffLookup.eMin2)
            End If
            If bMin3Locked = True Then
                decMineral3Payment = CDec(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
                decPayment += decMineral3Payment
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin3) = GetCoeffValue(elPropCoeffLookup.eMin3)
            End If
            If bMin4Locked = True Then
                decMineral4Payment = CDec(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
                decPayment += decMineral4Payment
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin4) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin4) = GetCoeffValue(elPropCoeffLookup.eMin4)
            Else
                bMin4Locked = True
                lLockedMin4 = 0
            End If
            If bMin5Locked = True Then
                decMineral5Payment = CDec(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
                decPayment += decMineral5Payment
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin5) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin5) = GetCoeffValue(elPropCoeffLookup.eMin5)
            Else
                bMin5Locked = True
                lLockedMin5 = 0
            End If
            If bMin6Locked = True Then
                decMineral6Payment = CDec(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))
                decPayment += decMineral6Payment
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin6) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin6) = GetCoeffValue(elPropCoeffLookup.eMin6)
            Else
                bMin6Locked = True
                lLockedMin6 = 0
            End If

            For X As Int32 = 0 To 14
                AddToTestString("Coeff " & X & ": " & GetCoeffValue(CType(X, elPropCoeffLookup)).ToString("###0.#####"))
            Next

            AddToTestString("PAYMENT ITEMS")
            AddToTestString("Payment (Locked): " & decPayment.ToString)

            'ok, we have our payment and the bill...
            If decPayment < decTotalBill Then

                Dim lResDistCnt As Int32 = 0
                If blLockedResCost = -1 Then lResDistCnt += 1
                If blLockedResTime = -1 Then lResDistCnt += 1
                If lLockedColonists = -1 Then lResDistCnt += 1
                If lResDistCnt <> lDistributable Then
                    If blLockedResCost = -1 Then
                        blLockedResCost = 100000000 'CLng(Math.Pow(100000000, GetCoeffValue(elPropCoeffLookup.eResCost)))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
                        If blNew + decPayment > decTotalBill Then
                            blLockedResCost = -1
                        Else
                            decResCostPayment = blNew
                            decPayment += blNew
                            AddToTestString("ResCost: " & blNew.ToString("#,###"))
                            decDistVals(elPropCoeffLookup.eResCost) = 0
                            lDistributable -= 1
                        End If
                    End If
                    If blLockedResTime = -1 Then
                        blLockedResTime = 2000000000 'CLng(Math.Pow(2000000000, GetCoeffValue(elPropCoeffLookup.eResTime)))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
                        If blNew + decPayment > decTotalBill Then
                            blLockedResTime = -1
                        Else
                            decResTimePayment = blNew
                            decPayment += blNew
                            AddToTestString("ResTime: " & blNew.ToString("#,###"))
                            decDistVals(elPropCoeffLookup.eResTime) = 0
                            lDistributable -= 1
                        End If
                    End If
                    If lLockedColonists = -1 Then
                        lLockedColonists = 20 'CInt(Math.Pow(decDistVals(elPropCoeffLookup.eColonist) * decDiffTemp, (1 / GetCoeffValue(elPropCoeffLookup.eColonist)))) \ 20
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                        If blNew + decPayment > decTotalBill Then
                            lLockedColonists = -1
                        Else
                            decColonistPayment = blNew
                            decPayment += blNew
                            AddToTestString("Colonists: " & blNew.ToString("#,###"))
                            decDistVals(elPropCoeffLookup.eColonist) = 0
                            lDistributable -= 1
                        End If


                    End If
                End If

                'ok, we owe more
                Dim decDiff As Decimal = decTotalBill - decPayment
                'we owe decDiff in points...
                If lDistributable <> 0 Then

                    Dim decTotal As Decimal = 0
                    For X As Int32 = 0 To decDistVals.GetUpperBound(0)
                        decTotal += decDistVals(X)
                    Next X
                    For X As Int32 = 0 To decDistVals.GetUpperBound(0)
                        If decDistVals(X) > 0 Then decDistVals(X) = decTotal / decDistVals(X)
                    Next X
                    decTotal = 0
                    For X As Int32 = 0 To decDistVals.GetUpperBound(0)
                        decTotal += decDistVals(X)
                    Next X
                    For X As Int32 = 0 To decDistVals.GetUpperBound(0)
                        If decDistVals(X) > 0 Then decDistVals(X) = decDistVals(X) / decTotal
                    Next X

                    'For X As Int32 = 0 To decDistVals.GetUpperBound(0)
                    '    AddToTestString("Distribute" & X & ": " & decDistVals(X).ToString)
                    'Next X

                    Dim blMorePay As Int64 = 0

                    'ok, take decDiff / lDistributable = Points to distribute to each unlocked property (lPointToDist)
                    ' each unlocked property is then raised by lPointToDist ^ (1 / GetCoeff)
                    'Dim blPointsToDistribute As Int64 = CLng(decDiff / lDistributable)
                    If lLockedHull = -1 Then
                        lLockedHull = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eHull) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eHull))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                        decHullPayment = blNew
                        blMorePay += blNew
                        AddToTestString("Hull: " & blNew.ToString("#,###"))
                    End If
                    If lLockedPower = -1 AndAlso GetCoeffValue(elPropCoeffLookup.ePower) <> 0 Then
                        lLockedPower = CInt(Math.Pow(decDistVals(elPropCoeffLookup.ePower) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.ePower))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                        decPowerPayment = blNew
                        blMorePay += blNew
                        AddToTestString("Power: " & blNew.ToString("#,###"))
                    End If
                    If blLockedResCost = -1 Then
                        blLockedResCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResCost))))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
                        decResCostPayment = blNew
                        blMorePay += blNew
                        AddToTestString("ResCost: " & blNew.ToString("#,###"))
                    End If
                    If blLockedResTime = -1 Then
                        blLockedResTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResTime))))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
                        decResTimePayment = blNew
                        blMorePay += blNew
                        AddToTestString("ResTime: " & blNew.ToString("#,###"))
                    End If
                    If blLockedProdCost = -1 Then
                        blLockedProdCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdCost))))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
                        decProdCostPayment = blNew
                        blMorePay += blNew
                        AddToTestString("ProdCost: " & blNew.ToString("#,###"))
                    End If
                    If blLockedProdTime = -1 Then
                        blLockedProdTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdTime))))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedProdTime, GetCoeffValue(elPropCoeffLookup.eProdTime)))
                        decProdTimePayment = blNew
                        blMorePay += blNew
                        AddToTestString("ProdTime: " & blNew.ToString("#,###"))
                    End If
                    If lLockedColonists = -1 Then
                        lLockedColonists = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eColonist) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eColonist))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                        decColonistPayment = blNew
                        blMorePay += blNew
                        AddToTestString("Colonist: " & blNew.ToString("#,###"))
                    End If
                    If lLockedEnlisted = -1 Then
                        lLockedEnlisted = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eEnlisted) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eEnlisted))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedEnlisted, GetCoeffValue(elPropCoeffLookup.eEnlisted)))
                        decEnlistedPayment = blNew
                        blMorePay += blNew
                        AddToTestString("Enlisted: " & blNew.ToString("#,###"))
                    End If
                    If lLockedOfficers = -1 Then
                        lLockedOfficers = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eOfficer) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eOfficer))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedOfficers, GetCoeffValue(elPropCoeffLookup.eOfficer)))
                        decOfficersPayment = blNew
                        blMorePay += blNew
                        AddToTestString("Officer: " & blNew.ToString("#,###"))
                    End If
                    If bMin1Locked = False Then
                        lLockedMin1 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin1) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin1))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
                        decMineral1Payment = blNew
                        blMorePay += blNew
                        AddToTestString("Min1: " & blNew.ToString("#,###"))
                    End If
                    If bMin2Locked = False Then
                        lLockedMin2 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin2) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin2))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
                        decMineral2Payment = blNew
                        blMorePay += blNew
                        AddToTestString("Min2: " & blNew.ToString("#,###"))
                    End If
                    If bMin3Locked = False Then
                        lLockedMin3 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin3) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin3))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
                        decMineral3Payment = blNew
                        blMorePay += blNew
                        AddToTestString("Min3: " & blNew.ToString("#,###"))
                    End If
                    If bMin4Locked = False Then
                        lLockedMin4 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin4) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin4))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
                        decMineral4Payment = blNew
                        blMorePay += blNew
                        AddToTestString("Min4: " & blNew.ToString("#,###"))
                    End If
                    If bMin5Locked = False Then
                        lLockedMin5 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin5) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin5))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
                        decMineral5Payment = blNew
                        blMorePay += blNew
                        AddToTestString("Min5: " & blNew.ToString("#,###"))
                    End If
                    If bMin6Locked = False Then
                        lLockedMin6 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin6) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin6))))
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))
                        decMineral6Payment = blNew
                        blMorePay += blNew
                        AddToTestString("Min6: " & blNew.ToString("#,###"))
                    End If

                    Dim decTotalPay As Decimal = (blMorePay + decPayment)
                    AddToTestString("MorePay: " & blMorePay.ToString("#,###"))
                    AddToTestString("Perc: " & ((decTotalPay / decTotalBill) * 100).ToString("##0.##") & "%")

                    frmBuildCost.SetBaseValues(lLockedHull, lLockedPower, blLockedResCost, blLockedResTime, blLockedProdCost, blLockedProdTime, lLockedColonists, lLockedEnlisted, lLockedOfficers, Me.lHullTypeID)
                    frmBuildCost.SetBillPaymentValues(decTotalBill, decTotalPay)

                    If tpMin1 Is Nothing = False AndAlso tpMin1.PropertyLocked = False Then
                        tpMin1.MaxValue = Math.Max(tpMin1.MaxValue, lLockedMin1)
                        tpMin1.PropertyValue = lLockedMin1
                    End If
                    If tpMin2 Is Nothing = False AndAlso tpMin2.PropertyLocked = False Then
                        tpMin2.MaxValue = Math.Max(tpMin2.MaxValue, lLockedMin2)
                        tpMin2.PropertyValue = lLockedMin2
                    End If
                    If tpMin3 Is Nothing = False AndAlso tpMin3.PropertyLocked = False Then
                        tpMin3.MaxValue = Math.Max(tpMin3.MaxValue, lLockedMin3)
                        tpMin3.PropertyValue = lLockedMin3
                    End If
                    If tpMin4 Is Nothing = False AndAlso tpMin4.PropertyLocked = False Then
                        tpMin4.MaxValue = Math.Max(tpMin4.MaxValue, lLockedMin4)
                        tpMin4.PropertyValue = lLockedMin4
                    End If
                    If tpMin5 Is Nothing = False AndAlso tpMin5.PropertyLocked = False Then
                        tpMin5.MaxValue = Math.Max(tpMin5.MaxValue, lLockedMin5)
                        tpMin5.PropertyValue = lLockedMin5
                    End If
                    If tpMin6 Is Nothing = False AndAlso tpMin6.PropertyLocked = False Then
                        tpMin6.MaxValue = Math.Max(tpMin6.MaxValue, lLockedMin6)
                        tpMin6.PropertyValue = lLockedMin6
                    End If

                    'ProdCost must be > 1% of ResCost or ResCost must be 1% of ProdCost
                    Dim decProdMin As Decimal = 0.01D
                    'If iTechTypeID = ObjectType.eRadarTech Then decProdMin = 0.001D

                    If blLockedProdCost < blLockedResCost Then
                        If blLockedProdCost < blLockedResCost * decProdMin Then
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "The Production Cost is too low." & vbCrLf & "(Raise Production Cost Or Lower Research Cost)"
                            lblDesignFlaw.Visible = True
                            lError = lError Or elErrorReasons.eProdCostToResCostRatio
                        End If
                    Else
                        If blLockedResCost < blLockedProdCost * decProdMin Then
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "The Research Cost is too low." & vbCrLf & "(Raise Research Cost Or Lower Production Cost)"
                            lblDesignFlaw.Visible = True
                            lError = lError Or elErrorReasons.eProdCostToResCostRatio
                        End If
                    End If

                    'factory is 528, reslab is 92

                    'prodtime must be 1/1000th of res time
                    Dim blCompareProdTime As Int64 = blLockedProdTime \ 528
                    Dim blCompareResTime As Int64 = blLockedResTime \ 92
                    If blCompareProdTime < blCompareResTime Then
                        If blCompareProdTime < blCompareResTime * 0.001 Then
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "The Production Time is too low." & vbCrLf & "(Raise Production Time or Lower Research Time)"
                            lblDesignFlaw.Visible = True
                            lError = lError Or elErrorReasons.eProdTimeToResTimeRatio
                        End If
                    Else
                        If blCompareResTime < blCompareProdTime * 0.001 Then
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "The Research Time is too low." & vbCrLf & "(Raise Research Time or Lower Production Time)"
                            lblDesignFlaw.Visible = True
                            lError = lError Or elErrorReasons.eProdTimeToResTimeRatio
                        End If
                    End If


                    'Sum Minerals AV > Hull Size AV
                    If lLockedMin1 + lLockedMin2 + lLockedMin3 + lLockedMin4 + lLockedMin5 + lLockedMin6 < lLockedHull Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "We need minerals to make this thing as large" & vbCrLf & "as the hull requirements specify. (Raise Minerals or Lower Hull)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eMineralsToHullSize
                    End If

                    'HullSize PP + Power PP > 10% of total points to be paid
                    'Dim decHull As Decimal = CDec(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                    'Dim decPower As Decimal = CDec(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                    Dim decMin As Decimal = 0.1D
                    'If iTechTypeID = ObjectType.eRadarTech Then decMin = 0.005D
                    If decHullPayment + decPowerPayment < decTotalBill * decMin Then
                        bImpossibleDesign = True
                        lError = lError Or elErrorReasons.eHullPowerPointsLow
                        If iTechTypeID = ObjectType.eEngineTech Then
                            lblDesignFlaw.Caption = "Unable to produce this design with the hull constraints" & vbCrLf & "placed on them. (Raise Hull Required)"
                        Else
                            lblDesignFlaw.Caption = "Unable to produce this design with the hull and power" & vbCrLf & "constraints placed on them. (Raise Hull or Power Required)"
                        End If

                        lblDesignFlaw.Visible = True
                    End If

                    'HullSize PP + Minerals PP > 15%
                    'Dim decMin1 As Decimal = CDec(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
                    'Dim decMin2 As Decimal = CDec(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
                    'Dim decMin3 As Decimal = CDec(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
                    'Dim decMin4 As Decimal = CDec(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
                    'Dim decMin5 As Decimal = CDec(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
                    'Dim decMin6 As Decimal = CDec(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))
                    Dim decHMMin As Decimal = 0.15D
                    'If iTechTypeID = ObjectType.eRadarTech Then decHMMin = 0.1D
                    If decHullPayment + decMineral1Payment + decMineral2Payment + decMineral3Payment + decMineral4Payment + decMineral5Payment + decMineral6Payment < decTotalBill * decHMMin Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "The hull size and minerals are insufficient to fit all of the" & vbCrLf & "capabilities we need from this component. (Raise Hull or Minerals)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eHullMineralPointsLow
                    End If

                    'Col PP + Enl PP + Off PP > 10%
                    'Dim decCol As Decimal = CDec(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                    'Dim decEnl As Decimal = CDec(Math.Pow(lLockedEnlisted, GetCoeffValue(elPropCoeffLookup.eEnlisted)))
                    'Dim decOff As Decimal = CDec(Math.Pow(lLockedOfficers, GetCoeffValue(elPropCoeffLookup.eOfficer)))
                    'If decColonistPayment + decEnlistedPayment + decOfficersPayment < decBill * 0.05D Then
                    '    bImpossibleDesign = True
                    '    lblDesignFlaw.Caption = "Not enough crew for this design."
                    '    lblDesignFlaw.Visible = True
                    'End If

                    'ResTime PP + ProdTime PP > 30% of total points to be paid
                    'ResTime PP + ProdTime PP < 60% of total points to be paid
                    'Dim decResTime As Decimal = CDec(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
                    'Dim decProdTime As Decimal = CDec(Math.Pow(blLockedProdTime, GetCoeffValue(elPropCoeffLookup.eProdTime)))
                    If decResTimePayment + decProdTimePayment < decTotalBill * 0.1D Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "Our scientists cannot get the design completed with the " & vbCrLf & "specified time constraints. (Increase Production or Research Time)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eResTimeProdTimePoints
                    ElseIf decResTimePayment + decProdTimePayment > decTotalBill * 0.9D Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "The time it will take to do this is not going to serve us well." & vbCrLf & "We need to reduce some of these times."
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eResTimeProdTimePoints
                    End If

                    'ResCost PP + ProdCost PP > 5% < 50%
                    'Dim decResCost As Decimal = CDec(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
                    'Dim decProdCost As Decimal = CDec(Math.Pow(blLockedProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
                    Dim decPCMin As Decimal = 0.05D
                    'If iTechTypeID = ObjectType.eRadarTech Then decPCMin = 0.02D
                    If decResCostPayment + decProdCostPayment < decPCMin * decTotalBill Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "The budgetary limitations make this design impossible." & vbCrLf & "(Raise Production or Research Cost)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eResCostProdCostPoints
                    ElseIf decProdCostPayment + decResCostPayment > 0.5D * decTotalBill Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "Our budget analysis is showing that the empire accountants" & vbCrLf & "are not going to be happy with this result. (Lower Costs)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eResCostProdCostPoints
                    End If

                    'HullRequired AV > Col AV + Enl AV + Off AV
                    If lLockedHull < lLockedColonists + lLockedEnlisted + lLockedOfficers Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "All these people for such a small component." & vbCrLf & "Let's get rid of some of these people. (Reduce Crew)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eHullToCrewRatio
                    End If

                    'Enl AV + Off AV > Col AV
                    If lLockedEnlisted + lLockedOfficers < lLockedColonists Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "We are not building a family cruise vessel. We need more" & vbCrLf & "qualified personnel. (Raise Enlisted/Officers or Reduce Colonists)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eMilitaryToCitizen
                    End If

                    'Enl AV < Off AV * 5
                    If lLockedEnlisted > lLockedOfficers * 5 Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "Not having enough officers will make the enlisted unruly." & vbCrLf & "(Raise Officers or Reduce Enlisted)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eLackOfOfficers
                    End If

                    'HullRequired AV > Crew Hull Size AV
                    'this rule only applies to hulltypes other than light fighter, medium fighter, heavy fighter, small vehicles and tanks
                    If lHullTypeID <> Base_Tech.eyHullType.Tank AndAlso lHullTypeID <> Base_Tech.eyHullType.LightFighter AndAlso lHullTypeID <> Base_Tech.eyHullType.MediumFighter AndAlso _
                       lHullTypeID <> Base_Tech.eyHullType.HeavyFighter AndAlso lHullTypeID <> Base_Tech.eyHullType.SmallVehicle Then
                        If lLockedHull < (lLockedOfficers + lLockedColonists + lLockedEnlisted) * goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence) Then
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "We are wasting far too much space on the crew given" & vbCrLf & "the hull required for this component. (Reduce Crew or Raise Hull)"
                            lblDesignFlaw.Visible = True
                            lError = lError Or elErrorReasons.eHullToCrewHullRatio
                        End If
                    End If

                    Dim decBill80 As Decimal = decTotalBill * 0.8D
                    If decResCostPayment > decBill80 OrElse decResTimePayment > decBill80 OrElse decProdCostPayment > decBill80 OrElse decProdTimePayment > decBill80 OrElse decMineral1Payment > decBill80 OrElse decMineral2Payment > decBill80 Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "All of your scientists have a look of disbelief." & vbCrLf & "(One of your values exceeds 80%)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.e80PercMarkerExceed
                    ElseIf decMineral3Payment > decBill80 OrElse decMineral4Payment > decBill80 OrElse decMineral5Payment > decBill80 OrElse decMineral6Payment > decBill80 OrElse decHullPayment > decBill80 OrElse decColonistPayment > decBill80 OrElse decEnlistedPayment > decBill80 OrElse decOfficersPayment > decBill80 Then
                        bImpossibleDesign = True
                        lblDesignFlaw.Caption = "All of your scientists have a look of disbelief." & vbCrLf & "(One of your values exceeds 80%)"
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.e80PercMarkerExceed
                    End If

                    'If iTechTypeID = ObjectType.eWeaponTech Then
                    '    'dec
                    '    If lLockedPower < decMaxThreshold Then
                    '        bImpossibleDesign = True
                    '        lblDesignFlaw.Caption = "Power cannot be less than " & (decMaxThreshold).ToString("#,###")
                    '        lblDesignFlaw.Visible = True
                    '    End If
                    'End If

                    'No mineral used can be less than 15% of total minerals...
                    Dim l15thMinerals As Int32 = CInt((lLockedMin1 + lLockedMin2 + lLockedMin3 + lLockedMin4 + lLockedMin5 + lLockedMin6) * 0.15)
                    If lMineral1ID > 0 AndAlso lLockedMin1 < l15thMinerals Then
                        bImpossibleDesign = True
                        If iTechTypeID = ObjectType.eEngineTech Then
                            lblDesignFlaw.Caption = "Unable to get the Structural Body that low."
                        Else
                            lblDesignFlaw.Caption = "Unable to get the " & tpMin1.PropertyName.Replace(":", "") & " that low."
                        End If
                        lError = lError Or elErrorReasons.eMineral1LessThan15
                        lblDesignFlaw.Visible = True
                    ElseIf lMineral2ID > 0 AndAlso lLockedMin2 < l15thMinerals Then
                        bImpossibleDesign = True
                        If iTechTypeID = ObjectType.eEngineTech Then
                            lblDesignFlaw.Caption = "Unable to get the Structural Frame that low."
                        Else
                            lblDesignFlaw.Caption = "Unable to get the " & tpMin2.PropertyName.Replace(":", "") & " that low."
                        End If
                        lError = lError Or elErrorReasons.eMineral2LessThan15
                        lblDesignFlaw.Visible = True
                    ElseIf lMineral3ID > 0 AndAlso lLockedMin3 < l15thMinerals Then
                        bImpossibleDesign = True
                        If iTechTypeID = ObjectType.eEngineTech Then
                            lblDesignFlaw.Caption = "Unable to get the Structural Meld that low."
                        Else
                            lblDesignFlaw.Caption = "Unable to get the " & tpMin3.PropertyName.Replace(":", "") & " that low."
                        End If
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eMineral3LessThan15
                    ElseIf lMineral4ID > 0 AndAlso lLockedMin4 < l15thMinerals Then
                        bImpossibleDesign = True
                        If iTechTypeID = ObjectType.eEngineTech Then
                            lblDesignFlaw.Caption = "Unable to get the Drive Body that low."
                        Else
                            lblDesignFlaw.Caption = "Unable to get the " & tpMin4.PropertyName.Replace(":", "") & " that low."
                        End If
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eMineral4LessThan15
                    ElseIf lMineral5ID > 0 AndAlso lLockedMin5 < l15thMinerals Then
                        bImpossibleDesign = True
                        If iTechTypeID = ObjectType.eEngineTech Then
                            lblDesignFlaw.Caption = "Unable to get the Drive Frame that low."
                        Else
                            lblDesignFlaw.Caption = "Unable to get the " & tpMin5.PropertyName.Replace(":", "") & " that low."
                        End If
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eMineral5LessThan15
                    ElseIf lMineral6ID > 0 AndAlso lLockedMin6 < l15thMinerals Then
                        bImpossibleDesign = True
                        If iTechTypeID = ObjectType.eEngineTech Then
                            lblDesignFlaw.Caption = "Unable to get the Drive Meld that low."
                        Else
                            lblDesignFlaw.Caption = "Unable to get the " & tpMin6.PropertyName.Replace(":", "") & " that low."
                        End If
                        lblDesignFlaw.Visible = True
                        lError = lError Or elErrorReasons.eMineral6LessThan15
                    End If

                    If iTechTypeID <> ObjectType.eEngineTech Then
                        decMin = 0.01D
                        'If iTechTypeID = ObjectType.eRadarTech Then decMin = 0.001D
                        If decPowerPayment < decTotalBill * decMin Then
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "It is impossible to power that component with" & vbCrLf & "such low of power needs. (Raise Power Required)"
                            lblDesignFlaw.Visible = True
                            lError = lError Or elErrorReasons.ePowerLessThan1Perc
                        End If
                        If decHullPayment < decTotalBill * decMin Then
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "It is impossible to fit that component into" & vbCrLf & "such a small space. (Raise Hull Required)"
                            lblDesignFlaw.Visible = True
                            lError = lError Or elErrorReasons.eHullLessThan1Perc
                        End If
                    End If

                    AddToTestString("ResPoints: " & blLockedResTime.ToString("#,###"))
                    AddToTestString("ProdPoints: " & blLockedProdTime.ToString("#,###"))

                    If bImpossibleDesign = False Then

                        Dim bNoCrewHull As Boolean = (lHullTypeID = Base_Tech.eyHullType.Escort OrElse lHullTypeID = Base_Tech.eyHullType.HeavyFighter OrElse lHullTypeID = Base_Tech.eyHullType.LightFighter OrElse lHullTypeID = Base_Tech.eyHullType.MediumFighter OrElse lHullTypeID = Base_Tech.eyHullType.Tank)

                        If lLockedHull < 1 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the hull that low."
                        ElseIf lLockedPower < 1 AndAlso iTechTypeID <> ObjectType.eEngineTech Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the power that low."
                        ElseIf lLockedOfficers < 1 AndAlso bOffLockedAtStart = True AndAlso bNoCrewHull = False Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the officers that low." & vbCrLf & "(Consider Unlocking)"
                        ElseIf lLockedEnlisted < 1 AndAlso bEnlLockedAtStart = True AndAlso bNoCrewHull = False Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the enlisted that low." & vbCrLf & "(Consider Unlocking)"
                        ElseIf lLockedColonists < 1 AndAlso bColLockedAtStart = True AndAlso bNoCrewHull = False Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the colonists that low." & vbCrLf & "(Consider Unlocking)"
                        ElseIf blLockedProdCost < 1 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the production cost that low."
                        ElseIf blLockedProdTime < 1 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the production time that low."
                        ElseIf blLockedResCost < 1 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the research cost that low."
                        ElseIf blLockedResTime < 1 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the research time that low."
                        ElseIf lLockedMin1 < 1 AndAlso lMineral1ID > 0 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the minerals that low."
                        ElseIf lLockedMin2 < 1 AndAlso lMineral2ID > 0 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the minerals that low."
                        ElseIf lLockedMin3 < 1 AndAlso lMineral3ID > 0 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the minerals that low."
                        ElseIf lLockedMin4 < 1 AndAlso lMineral4ID > 0 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the minerals that low."
                        ElseIf lLockedMin5 < 1 AndAlso lMineral5ID > 0 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the minerals that low."
                        ElseIf lLockedMin6 < 1 AndAlso lMineral6ID > 0 Then
                            lblDesignFlaw.Visible = True
                            bImpossibleDesign = True
                            lblDesignFlaw.Caption = "Your scientists cannot reduce the minerals that low."
                        End If
                        If bImpossibleDesign = False Then
                            lblDesignFlaw.Visible = True
                            lblDesignFlaw.Caption = "Your scientists believe this design to be possible."
                        Else
                            lError = lError Or elErrorReasons.eValueIsBelowZero
                        End If
                    End If
                Else
                    bImpossibleDesign = True
                    lblDesignFlaw.Caption = "This design is impossible to comprehend." & vbCrLf & "Your scientists must use costs somewhere."
                    lblDesignFlaw.Visible = True
                End If

            Else
                frmBuildCost.SetBillPaymentValues(decTotalBill, decPayment)
                bImpossibleDesign = True
                lblDesignFlaw.Caption = "Total of all percentages cannot exceed 100%" & vbCrLf & "Reduce some cost-related values to balance out."
                lblDesignFlaw.Visible = True
                lError = lError Or elErrorReasons.ePaymentOverBill
            End If
        Catch
            lblDesignFlaw.Visible = True
            lblDesignFlaw.Caption = "This design is impossible to comprehend."
            bImpossibleDesign = True
        End Try

        If bShowDebug = True Then
            Dim oResFrm As frmTechBuilderDisplay = CType(goUILib.GetWindow("frmTechBuilderDisplay"), frmTechBuilderDisplay)
            If oResFrm Is Nothing Then oResFrm = New frmTechBuilderDisplay(goUILib)
            oResFrm.SetValues(moSB.ToString)
            moSB.Length = 0
        End If

        If bImpossibleDesign = True Then
            lblDesignFlaw.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        Else
            lblDesignFlaw.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        End If

        frmBuildCost.SetPaymentPercs(Me)
        If tpMin1 Is Nothing = False Then tpMin1.SetPercOfPayment(GetPercValueFromDecs(decTotalBill, decMineral1Payment))
        If tpMin2 Is Nothing = False Then tpMin2.SetPercOfPayment(GetPercValueFromDecs(decTotalBill, decMineral2Payment))
        If tpMin3 Is Nothing = False Then tpMin3.SetPercOfPayment(GetPercValueFromDecs(decTotalBill, decMineral3Payment))
        If tpMin4 Is Nothing = False Then tpMin4.SetPercOfPayment(GetPercValueFromDecs(decTotalBill, decMineral4Payment))
        If tpMin5 Is Nothing = False Then tpMin5.SetPercOfPayment(GetPercValueFromDecs(decTotalBill, decMineral5Payment))
        If tpMin6 Is Nothing = False Then tpMin6.SetPercOfPayment(GetPercValueFromDecs(decTotalBill, decMineral6Payment))

        frmBuildCost.SetFromErrorCode(lError, iTechTypeID, tpMin1, tpMin2, tpMin3, tpMin4, tpMin5, tpMin6)

        mbIgnoreValueChange = False

    End Sub


    Private Function GetPercValueFromDecs(ByVal decTotal As Decimal, ByVal decValue As Decimal) As String
        If decTotal < 0.0001D Then Return "0%"

        Dim decTemp As Decimal = (decValue / decTotal) * 100
        If decTemp < 0 Then Return "0%"
        If decTemp > 255 Then Return "255%"
        Return decTemp.ToString("##0.#") & "%"
    End Function

    'MSC - added for setting mineral DA values from the builders
    Public Sub SetMinDAValues(ByVal l1Val As Int32, ByVal l2Val As Int32, ByVal l3Val As Int32, ByVal l4Val As Int32, ByVal l5Val As Int32, ByVal l6Val As Int32)
        ml_Min1DA = l1Val
        ml_Min2DA = l2Val
        ml_Min3DA = l3Val
        ml_Min4DA = l4Val
        ml_Min5DA = l5Val
        ml_Min6DA = l6Val
    End Sub
End Class
