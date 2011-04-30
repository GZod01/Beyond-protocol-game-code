Option Strict On

Public MustInherit Class TechBuilderComputer
    'Protected Shared moSB As System.Text.StringBuilder

    Protected Enum elPropCoeffLookup As Int32
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
    'Protected MustOverride Function CalculateThresholdValue() As Decimal
    'Protected MustOverride Sub LoadPointVals()
    Protected MustOverride Sub SetMinReqProps()
    Protected MustOverride Function GetCoefficient(ByVal lLookup As elPropCoeffLookup) As Decimal
    Protected MustOverride Function GetTheBill() As Decimal

    Public lMineral1ID As Int32 = -1
    Public lMineral2ID As Int32 = -1
    Public lMineral3ID As Int32 = -1
    Public lMineral4ID As Int32 = -1
    Public lMineral5ID As Int32 = -1
    Public lMineral6ID As Int32 = -1

    Public lLockedMin1 As Int32 = -1
    Public lLockedMin2 As Int32 = -1
    Public lLockedMin3 As Int32 = -1
    Public lLockedMin4 As Int32 = -1
    Public lLockedMin5 As Int32 = -1
    Public lLockedMin6 As Int32 = -1
    Public lLockedHull As Int32 = -1
    Public lLockedPower As Int32 = -1
    Public blLockedResCost As Int64 = -1
    Public blLockedResTime As Int64 = -1
    Public blLockedProdCost As Int64 = -1
    Public blLockedProdTime As Int64 = -1
    Public lLockedColonists As Int32 = -1
    Public lLockedEnlisted As Int32 = -1
    Public lLockedOfficers As Int32 = -1

    Public lResultMin1 As Int32 = -1
    Public lResultMin2 As Int32 = -1
    Public lResultMin3 As Int32 = -1
    Public lResultMin4 As Int32 = -1
    Public lResultMin5 As Int32 = -1
    Public lResultMin6 As Int32 = -1
    Public lResultHull As Int32 = -1
    Public lResultPower As Int32 = -1
    Public blResultResCost As Int64 = -1
    Public blResultResTime As Int64 = -1
    Public blResultProdCost As Int64 = -1
    Public blResultProdTime As Int64 = -1
    Public lResultColonists As Int32 = -1
    Public lResultEnlisted As Int32 = -1
    Public lResultOfficers As Int32 = -1

    Public bIsMicroTech As Boolean = False

    Protected Structure uMinPropReq
        Public lPropID As Int32
        Public lMinID As Int32
        Public lPropVal As Int32
    End Structure
    Protected muPropReqs() As uMinPropReq

    Public Shared Sub GetTypeValues(ByVal plHullType As Int32, ByRef pdecNormalizer As Decimal, ByRef plMaxGuns As Int32, ByRef plMaxDPS As Int32, ByRef plMaxHullSize As Int32, ByRef plHullAvail As Int32, ByRef plPower As Int32)
        'NOTE: HullAvail is not used anywhere - if it becomes used, we need to recalculate the values in here
        Select Case CType(plHullType, Epica_Tech.eyHullType)
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

    'Protected Shared Function InvalidHalfBaseCheck(ByVal blBaseValue As Int64, ByVal blLockedValue As Int64, ByVal yType As Byte) As Boolean
    '    Dim blMinVal As Int64 = 0
    '    If yType = 0 Then
    '        blMinVal = blBaseValue \ 2
    '        If blMinVal < 1 Then blMinVal = 1
    '    ElseIf yType = 1 Then
    '        blMinVal = CLng(blBaseValue * 0.1F)
    '        If blMinVal < 0 Then blMinVal = 0
    '    ElseIf yType = 2 Then
    '        blMinVal = CLng(Math.Ceiling(blBaseValue * 0.25F))
    '        If blMinVal < 1 Then blMinVal = 1
    '    Else
    '        blMinVal = CLng(Math.Min(1, blBaseValue))
    '    End If

    '    If blLockedValue < blMinVal Then
    '        Return True
    '    End If

    '    Return False
    '    'If yType = 0 Then
    '    '    If blLockedValue < blBaseValue / 2 Then
    '    '        Return True
    '    '    End If
    '    'ElseIf yType = 1 Then
    '    '    If blLockedValue < blBaseValue * 0.01F Then
    '    '        Return True
    '    '    End If
    '    'ElseIf yType = 2 Then
    '    '    If blLockedValue < blBaseValue * 0.25F Then
    '    '        Return True
    '    '    End If
    '    'Else
    '    '    If blLockedValue < Math.Min(1, blBaseValue) Then
    '    '        Return True
    '    '    End If
    '    'End If

    '    'Return False
    'End Function

    'Protected Shared Function DoLockedToBaseCheck(ByRef lLockedValue As Int32, ByRef lBaseValue As Int32) As Int32
    '    Dim lResult As Int32 = 0
    '    If lLockedValue = -1 Then
    '        'value is not locked, set the value so we can set it later...
    '        'lLockedValue = lBaseValue
    '        'and clear the lBaseHull value
    '        'lBaseValue = 0
    '        'lLockedValue = 0
    '        lResult = 0
    '    Else
    '        'value is locked, determine our difference
    '        lResult = lBaseValue - lLockedValue
    '        'If lResult > 0 Then
    '        '    Dim decTemp As Decimal = CDec(lResult) * CDec(lResult)
    '        '    lResult = CInt(Math.Max(Int32.MinValue, Math.Min(Int32.MaxValue, decTemp)))
    '        'End If
    '        'Dim decTemp As Decimal = CDec(lResult) * CDec(lResult)
    '        'If lResult < 0 Then
    '        '    decTemp = -decTemp
    '        'End If
    '        'lResult = CInt(Math.Max(Int32.MinValue, Math.Min(Int32.MaxValue, decTemp)))

    '        'Dim decTemp As Decimal = lBaseValue - lLockedValue
    '        'Dim lMult As Int32 = 1
    '        'If decTemp < 1 Then lMult = -1
    '        'decTemp = CDec(Math.Pow(Math.Abs(decTemp), 1.2)) * lMult
    '        'lBaseValue = CInt(Math.Max(Int32.MinValue, Math.Min(Int32.MaxValue, decTemp)))

    '        ''Trying this on for size...
    '        ''  (BaseValue - DesiredValue) ^ (1 + (1 - (DesiredValue / BaseValue)))
    '        'Dim decTemp As Decimal = (lBaseValue - lLockedValue)
    '        'Dim lTemp As Int32 = Math.Max(1, lBaseValue)
    '        'decTemp = CDec(Math.Pow(decTemp, 1 + (1 - (lLockedValue / lTemp))))
    '        'If decTemp > Int32.MaxValue Then
    '        '    lBaseValue = Int32.MaxValue
    '        'ElseIf decTemp < Int32.MinValue Then
    '        '    lBaseValue = Int32.MinValue
    '        'Else : lBaseValue = CInt(decTemp)
    '        'End If
    '    End If
    '    'AddToTestString("L2B Result: " & lResult.ToString)
    '    Return lResult
    'End Function
    'Protected Shared Function DoLockedToBaseCheck(ByRef lLockedValue As Int64, ByRef lBaseValue As Int64) As Int64
    '    Dim blResult As Int64 = 0
    '    If lLockedValue = -1 Then
    '        'value is not locked, set the value so we can set it later...
    '        'lLockedValue = lBaseValue
    '        'and clear the lBaseHull value
    '        'lBaseValue = 0
    '        'lLockedValue = 0
    '        blResult = 0
    '    Else
    '        'value is locked, determine our difference
    '        blResult = lBaseValue - lLockedValue
    '        'If blResult > 0 Then
    '        '    Dim decTemp As Decimal = CDec(blResult) * CDec(blResult)
    '        '    blResult = CLng(Math.Max(Int64.MinValue, Math.Min(Int64.MaxValue, decTemp)))
    '        'End If
    '        'Dim decTemp As Decimal = CDec(blResult) * CDec(blResult)
    '        'If blResult < 0 Then
    '        '    decTemp = -decTemp
    '        'End If
    '        'blResult = CLng(Math.Max(Int64.MinValue, Math.Min(Int64.MaxValue, decTemp)))

    '        'Dim decTemp As Decimal = lBaseValue - lLockedValue
    '        'Dim lMult As Int32 = 1
    '        'If decTemp < 1 Then lMult = -1
    '        'decTemp = CDec(Math.Pow(Math.Abs(decTemp), 1.2)) * lMult
    '        'lBaseValue = CLng(Math.Max(Int64.MinValue, Math.Min(Int64.MaxValue, decTemp)))

    '        ''trying this on for size...
    '        ''  (BaseValue - DesiredValue) ^ (1 + (1 - (DesiredValue / BaseValue)))
    '        'Dim decTemp As Decimal = (lBaseValue - lLockedValue)
    '        'Dim lTemp As Int64 = Math.Max(1, lBaseValue)
    '        'decTemp = CDec(Math.Pow(decTemp, 1 + (1 - (lLockedValue / lTemp))))
    '        'If decTemp > Int64.MaxValue Then
    '        '    lBaseValue = Int64.MaxValue
    '        'ElseIf decTemp < Int64.MinValue Then
    '        '    lBaseValue = Int64.MinValue
    '        'Else : lBaseValue = CLng(decTemp)
    '        'End If
    '    End If
    '    'AddToTestString("L2B Result: " & blResult.ToString)
    '    Return blResult
    'End Function

    ''Protected Shared Sub AddToTestString(ByVal sLine As String)
    ''    If moSB Is Nothing Then moSB = New System.Text.StringBuilder
    ''    moSB.AppendLine(sLine)
    ''End Sub

    'Protected Shared Sub AffectValueByPoints(ByRef blPoints As Int64, ByVal lPointPerValue As Int32, ByRef lValue As Int32, ByVal lModByVal As Int32)
    '    If blPoints < lPointPerValue Then Return
    '    If lValue = Int32.MaxValue Then Return
    '    lValue += lModByVal
    '    blPoints -= lPointPerValue
    'End Sub
    'Protected Shared Sub AffectValueByPoints(ByRef blPoints As Int64, ByVal lPointPerValue As Int32, ByRef blValue As Int64, ByVal lModByVal As Int32)
    '    If blPoints < lPointPerValue Then Return
    '    If blValue = Int64.MaxValue Then Return
    '    blValue += lModByVal
    '    blPoints -= lPointPerValue
    'End Sub

    'Private Function GetPointValue(ByVal blVal As Int64, ByVal sName As String, ByVal blPerPointMult As Int64) As Int64
    '    Dim blResult As Int64 = 0
    '    If blVal < 0 Then
    '        blVal = Math.Abs(blVal * blPerPointMult)

    '        'If sName.ToUpper = "RESTIME" Then
    '        '    blVal \= 180000
    '        'ElseIf sName.ToUpper = "PRODTIME" Then
    '        '    blVal \= 900000
    '        'End If

    '        blResult = -CLng((1D / Math.Log10(blVal)) * blVal)  '-CLng((1 / Math.Log10(blVal)) * blVal)
    '        blVal *= blPerPointMult

    '        'If sName.ToUpper = "RESTIME" Then
    '        '    blResult *= 180000
    '        'ElseIf sName.ToUpper = "PRODTIME" Then
    '        '    blResult *= 900000
    '        'End If
    '        'AddToTestString(sName & " grants " & blResult.ToString("#,##0"))
    '    Else
    '        blResult = blVal * blPerPointMult
    '    End If
    '    Return blResult
    'End Function

    'Public Function BuilderCostValueChange(ByVal iTechTypeID As Int16, ByVal decMaxThreshold As Decimal, ByVal lPopIntel As Int32) As Boolean
    '    'If moSB Is Nothing = False Then moSB.Length = 0

    '    Try
    '        LoadPointVals()
    '        SetMinReqProps()

    '        Dim decNormalize As Decimal = 1D
    '        Dim lMaxGuns As Int32 = 1
    '        Dim lMaxDPS As Int32 = 1
    '        Dim lMaxHullSize As Int32 = 1
    '        Dim lHullAvail As Int32 = 1
    '        Dim lMaxPower As Int32 = 1

    '        If lHullTypeID = -1 AndAlso iTechTypeID <> ObjectType.eArmorTech Then
    '            'lblDesignFlaw.Caption = "Invalid Design: a Hull Type must be selected!"
    '            'mbIgnoreValueChange = False
    '            Return False
    '        End If

    '        GetTypeValues(lHullTypeID, decNormalize, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

    '        'AddToTestString("N: " & decNormalize.ToString & ", MG: " & lMaxGuns.ToString & ", MDPS: " & lMaxDPS & ", MH: " & lMaxHullSize & ", HA: " & lHullAvail & ", MP: " & lMaxPower)

    '        'Ok, something has changed, let's calculate our values
    '        'This occurs in the following steps/order
    '        ' 1) Determine the required values lookups for the mineral properties
    '        ' 2) Calculate all base values for the costs
    '        '   a) Hull
    '        Dim lBaseHull As Int32 = CalculateBaseHull()
    '        'AddToTestString("Basehull: " & lBaseHull.ToString)
    '        '   b) Power - we do not use power here
    '        Dim lBasePower As Int32 = CalculateBasePower()
    '        'AddToTestString("lBasePower: " & lBasePower.ToString)
    '        '   c) Colonists
    '        Dim lBaseColonists As Int32 = CalculateBaseColonists()
    '        'AddToTestString("BaseColonists: " & lBaseColonists.ToString)
    '        '   d) Enlisted
    '        Dim lBaseEnlisted As Int32 = CalculateBaseEnlisted()
    '        'AddToTestString("BaseEnlisted: " & lBaseEnlisted.ToString)
    '        '   e) Officers
    '        Dim lBaseOfficers As Int32 = CalculateBaseOfficers()
    '        'AddToTestString("BaseOfficers: " & lBaseOfficers.ToString)
    '        '   f) Research Cost
    '        Dim blBaseResCost As Int64 = CalculateBaseResCost()
    '        'AddToTestString("BaseResCost: " & blBaseResCost.ToString)
    '        '   g) Research Time
    '        Dim blBaseResTime As Int64 = CalculateBaseResTime()
    '        'AddToTestString("BaseResTime: " & blBaseResTime.ToString)
    '        '   h) Production Cost
    '        Dim blBaseProdCost As Int64 = CalculateBaseProdCost()
    '        'AddToTestString("BaseProdCost: " & blBaseProdCost.ToString)
    '        '   i) Production Time
    '        Dim blBaseProdTime As Int64 = CalculateBaseProdTime()
    '        'AddToTestString("BaseProdTime: " & blBaseProdTime.ToString)
    '        '   j) Material Costs

    '        'MyBase.mfrmBuilderCost.SetMinValsFromCalcs(lBaseHull, 0, blBaseResCost, blBaseProdCost, lBaseColonists, lBaseEnlisted, lBaseOfficers)
    '        lSumAllDA = 0
    '        Dim lBaseMin1 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral1ID, 0), GetSumClientDA(lMineral1ID, 0))
    '        'AddToTestString("BaseMin1: " & lBaseMin1.ToString)
    '        Dim lBaseMin2 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral2ID, 1), GetSumClientDA(lMineral2ID, 1))
    '        'AddToTestString("BaseMin2: " & lBaseMin2.ToString)
    '        Dim lBaseMin3 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral3ID, 2), GetSumClientDA(lMineral3ID, 2))
    '        'AddToTestString("BaseMin3: " & lBaseMin3.ToString)
    '        Dim lBaseMin4 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral4ID, 3), GetSumClientDA(lMineral4ID, 3))
    '        'AddToTestString("BaseMin4: " & lBaseMin4.ToString)
    '        Dim lBaseMin5 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral5ID, 4), GetSumClientDA(lMineral5ID, 4))
    '        'AddToTestString("BaseMin5: " & lBaseMin5.ToString)
    '        Dim lBaseMin6 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral6ID, 5), GetSumClientDA(lMineral6ID, 5))
    '        'AddToTestString("BaseMin6: " & lBaseMin6.ToString)

    '        'tpStructBody.MinValue = lBaseStructBody \ 2
    '        'tpStructFrame.MinValue = lBaseStructFrame \ 2
    '        'tpStructMeld.MinValue = lBaseStructMeld \ 2
    '        'tpDriveBody.MinValue = lBaseDriveBody \ 2
    '        'tpDriveFrame.MinValue = lBaseDriveFrame \ 2
    '        'tpDriveMeld.MinValue = lBaseDriveMeld \ 2

    '        ' 3) If the value is locked for the cost, determine the difference between the locked value and the calculated value
    '        Dim bMin1Locked, bMin2Locked, bMin3Locked, bMin4Locked, bMin5Locked, bMin6Locked As Boolean
    '        bMin1Locked = lLockedMin1 <> -1
    '        bMin2Locked = lLockedMin2 <> -1
    '        bMin3Locked = lLockedMin3 <> -1
    '        bMin4Locked = lLockedMin4 <> -1
    '        bMin5Locked = lLockedMin5 <> -1
    '        bMin6Locked = lLockedMin6 <> -1

    '        If lMineral1ID < 1 Then
    '            lLockedMin1 = 0 : lBaseMin1 = 0
    '        End If
    '        If lMineral2ID < 1 Then
    '            lLockedMin2 = 0 : lBaseMin2 = 0
    '        End If
    '        If lMineral3ID < 1 Then
    '            lLockedMin3 = 0 : lBaseMin3 = 0
    '        End If
    '        If lMineral4ID < 1 Then
    '            lLockedMin4 = 0 : lBaseMin4 = 0
    '        End If
    '        If lMineral5ID < 1 Then
    '            lLockedMin5 = 0 : lBaseMin5 = 0
    '        End If
    '        If lMineral6ID < 1 Then
    '            lLockedMin6 = 0 : lBaseMin6 = 0
    '        End If

    '        Dim bHullLocked As Boolean = lLockedHull <> -1
    '        Dim bPowerLocked As Boolean = lLockedPower <> -1
    '        Dim bResCostLocked As Boolean = blLockedResCost <> -1
    '        Dim bResTimeLocked As Boolean = blLockedResTime <> -1
    '        Dim bProdCostLocked As Boolean = blLockedProdCost <> -1
    '        Dim bProdTimeLocked As Boolean = blLockedProdTime <> -1
    '        Dim bColonistLocked As Boolean = lLockedColonists <> -1
    '        Dim bEnlistedLocked As Boolean = lLockedEnlisted <> -1
    '        Dim bOfficersLocked As Boolean = lLockedOfficers <> -1

    '        Dim blOrigLockedResTime As Int64 = blLockedResTime
    '        Dim blOrigLockedProdTime As Int64 = blLockedProdTime

    '        'AddToTestString("ShiftedResTime: " & blBaseResTime.ToString)
    '        'AddToTestString("ShiftedProdTime: " & blBaseProdTime.ToString)

    '        lResultHull = 0 'lBaseHull
    '        If bHullLocked = True Then lResultHull = lLockedHull Else lResultHull = lBaseHull
    '        lResultPower = 0
    '        If bPowerLocked = True Then lResultPower = lLockedPower Else lResultPower = lBasePower
    '        blResultResCost = 0
    '        If bResCostLocked = True Then blResultResCost = blLockedResCost Else blResultResCost = blBaseResCost
    '        blResultResTime = 0
    '        If bResTimeLocked = True Then blResultResTime = blLockedResTime Else blResultResTime = blBaseResTime
    '        blResultProdCost = 0
    '        If bProdCostLocked = True Then blResultProdCost = blLockedProdCost Else blResultProdCost = blBaseProdCost
    '        blResultProdTime = 0
    '        If bProdTimeLocked = True Then blResultProdTime = blLockedProdTime Else blResultProdTime = blBaseProdTime
    '        lResultColonists = 0
    '        If bColonistLocked = True Then lResultColonists = lLockedColonists Else lResultColonists = lBaseColonists
    '        lResultEnlisted = 0
    '        If bEnlistedLocked = True Then lResultEnlisted = lLockedEnlisted Else lResultEnlisted = lBaseEnlisted
    '        lResultOfficers = 0
    '        If bOfficersLocked = True Then lResultOfficers = lLockedOfficers Else lResultOfficers = lBaseOfficers
    '        lResultMin1 = 0
    '        If bMin1Locked = True Then lResultMin1 = lLockedMin1 Else lResultMin1 = lBaseMin1
    '        lResultMin2 = 0
    '        If bMin2Locked = True Then lResultMin2 = lLockedMin2 Else lResultMin2 = lBaseMin2
    '        lResultMin3 = 0
    '        If bMin3Locked = True Then lResultMin3 = lLockedMin3 Else lResultMin3 = lBaseMin3
    '        lResultMin4 = 0
    '        If bMin4Locked = True Then lResultMin4 = lLockedMin4 Else lResultMin4 = lBaseMin4
    '        lResultMin5 = 0
    '        If bMin5Locked = True Then lResultMin5 = lLockedMin5 Else lResultMin5 = lBaseMin5
    '        lResultMin6 = 0
    '        If bMin6Locked = True Then lResultMin6 = lLockedMin6 Else lResultMin6 = lBaseMin6

    '        If bHullLocked = True AndAlso InvalidHalfBaseCheck(lBaseHull, lLockedHull, 2) = True Then Return False
    '        If bPowerLocked = True AndAlso InvalidHalfBaseCheck(lBasePower, lLockedPower, 2) = True Then Return False
    '        If bResCostLocked = True AndAlso InvalidHalfBaseCheck(blBaseResCost, blLockedResCost, 0) = True Then Return False
    '        If bProdCostLocked = True AndAlso InvalidHalfBaseCheck(blBaseProdCost, blLockedProdCost, 1) = True Then Return False
    '        If bColonistLocked = True AndAlso InvalidHalfBaseCheck(lBaseColonists, lLockedColonists, 3) = True Then Return False
    '        If bEnlistedLocked = True AndAlso InvalidHalfBaseCheck(lBaseEnlisted, lLockedEnlisted, 3) = True Then Return False
    '        If bOfficersLocked = True AndAlso InvalidHalfBaseCheck(lBaseOfficers, lLockedOfficers, 3) = True Then Return False
    '        If bMin1Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin1, lLockedMin1, 0) = True Then Return False
    '        If bMin2Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin2, lLockedMin2, 0) = True Then Return False
    '        If bMin3Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin3, lLockedMin3, 0) = True Then Return False
    '        If bMin4Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin4, lLockedMin4, 0) = True Then Return False
    '        If bMin5Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin5, lLockedMin5, 0) = True Then Return False
    '        If bMin6Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin6, lLockedMin6, 0) = True Then Return False
    '        If (bResTimeLocked = True AndAlso blLockedResTime < 1) OrElse (bProdTimeLocked = True AndAlso blLockedProdTime < 1) Then
    '            'lblDesignFlaw.Caption = "Your scientists cannot get the time that low."
    '            'lblDesignFlaw.Visible = True
    '            'mbIgnoreValueChange = False
    '            Return False
    '        End If

    '        Dim lHullDiff As Int32 = DoLockedToBaseCheck(lLockedHull, lBaseHull)
    '        Dim lPowerDiff As Int32 = DoLockedToBaseCheck(lLockedPower, lBasePower)
    '        Dim blResCostDiff As Int64 = DoLockedToBaseCheck(blLockedResCost, blBaseResCost)
    '        Dim blResTimeDiff As Int64 = DoLockedToBaseCheck(blLockedResTime, blBaseResTime)
    '        Dim blProdCostDiff As Int64 = DoLockedToBaseCheck(blLockedProdCost, blBaseProdCost)
    '        Dim blProdTimeDiff As Int64 = DoLockedToBaseCheck(blLockedProdTime, blBaseProdTime)
    '        Dim lColonistDiff As Int32 = DoLockedToBaseCheck(lLockedColonists, lBaseColonists)
    '        Dim lEnlistedDiff As Int32 = DoLockedToBaseCheck(lLockedEnlisted, lBaseEnlisted)
    '        Dim lOfficerDiff As Int32 = DoLockedToBaseCheck(lLockedOfficers, lBaseOfficers)
    '        Dim lMin1Diff As Int32 = DoLockedToBaseCheck(lLockedMin1, lBaseMin1)
    '        Dim lMin2Diff As Int32 = DoLockedToBaseCheck(lLockedMin2, lBaseMin2)
    '        Dim lMin3Diff As Int32 = DoLockedToBaseCheck(lLockedMin3, lBaseMin3)
    '        Dim lMin4Diff As Int32 = DoLockedToBaseCheck(lLockedMin4, lBaseMin4)
    '        Dim lMin5Diff As Int32 = DoLockedToBaseCheck(lLockedMin5, lBaseMin5)
    '        Dim lMin6Diff As Int32 = DoLockedToBaseCheck(lLockedMin6, lBaseMin6)

    '        ' 4) Determine the point costs for each item 
    '        Dim blPoints As Int64 = 0
    '        'blPoints += CLng(lHullDiff) * ml_POINT_PER_HULL
    '        'blPoints += CLng(lPowerDiff) * ml_POINT_PER_POWER
    '        'blPoints += blResTimeDiff * ml_POINT_PER_RES_TIME
    '        'blPoints += blResCostDiff * ml_POINT_PER_RES_COST
    '        'blPoints += blProdTimeDiff * ml_POINT_PER_PROD_TIME
    '        'blPoints += blProdCostDiff * ml_POINT_PER_PROD_COST
    '        'blPoints += lColonistDiff * ml_POINT_PER_COLONIST
    '        'blPoints += lEnlistedDiff * ml_POINT_PER_ENLISTED
    '        'blPoints += lOfficerDiff * ml_POINT_PER_OFFICER
    '        'blPoints += lMin1Diff * ml_POINT_PER_MIN1
    '        'blPoints += lMin2Diff * ml_POINT_PER_MIN2
    '        'blPoints += lMin3Diff * ml_POINT_PER_MIN3
    '        'blPoints += lMin4Diff * ml_POINT_PER_MIN4
    '        'blPoints += lMin5Diff * ml_POINT_PER_MIN5
    '        'blPoints += lMin6Diff * ml_POINT_PER_MIN6
    '        blPoints += GetPointValue(lHullDiff, "Hull", ml_POINT_PER_HULL)
    '        blPoints += GetPointValue(lPowerDiff, "Power", ml_POINT_PER_POWER)
    '        blPoints += GetPointValue(blResTimeDiff, "ResTime", ml_POINT_PER_RES_TIME)
    '        blPoints += GetPointValue(blResCostDiff, "ResCost", ml_POINT_PER_RES_COST)
    '        blPoints += GetPointValue(blProdTimeDiff, "ProdTime", ml_POINT_PER_PROD_TIME)
    '        blPoints += GetPointValue(blProdCostDiff, "ProdCost", ml_POINT_PER_PROD_COST)
    '        blPoints += GetPointValue(lColonistDiff, "Colonist", ml_POINT_PER_COLONIST)
    '        blPoints += GetPointValue(lEnlistedDiff, "Enlisted", ml_POINT_PER_ENLISTED)
    '        blPoints += GetPointValue(lOfficerDiff, "Officer", ml_POINT_PER_OFFICER)
    '        blPoints += GetPointValue(lMin1Diff, "Min1", ml_POINT_PER_MIN1)
    '        blPoints += GetPointValue(lMin2Diff, "Min2", ml_POINT_PER_MIN2)
    '        blPoints += GetPointValue(lMin3Diff, "Min3", ml_POINT_PER_MIN3)
    '        blPoints += GetPointValue(lMin4Diff, "Min4", ml_POINT_PER_MIN4)
    '        blPoints += GetPointValue(lMin5Diff, "Min5", ml_POINT_PER_MIN5)
    '        blPoints += GetPointValue(lMin6Diff, "Min6", ml_POINT_PER_MIN6)

    '        'AddToTestString("BasePoints: " & blPoints.ToString)
    '        If blPoints < 0 Then Return False

    '        'Now, check for our additive points from being uber
    '        Dim decBA As Decimal = CDec(lMaxGuns) / decNormalize '/ CDec(lMaxGuns)
    '        Dim decThreshold As Decimal = CalculateThresholdValue()
    '        'IF Power*Thrust*Manuever*Speed*B/A > 1,350,000,000 THEN
    '        'AddToTestString("Threshold: " & decThreshold.ToString)
    '        If iTechTypeID = ObjectType.eShieldTech Then
    '            If CDec(blPoints + decThreshold) > Int64.MaxValue Then blPoints = Int64.MaxValue Else blPoints = CLng(blPoints + decThreshold)
    '        Else
    '            If decThreshold > decMaxThreshold Then
    '                Dim blAllConstants As Int64 = ml_POINT_PER_HULL
    '                blAllConstants += ml_POINT_PER_POWER
    '                blAllConstants += ml_POINT_PER_RES_TIME
    '                blAllConstants += ml_POINT_PER_RES_COST
    '                blAllConstants += ml_POINT_PER_PROD_TIME
    '                blAllConstants += ml_POINT_PER_PROD_COST
    '                blAllConstants += ml_POINT_PER_COLONIST
    '                blAllConstants += ml_POINT_PER_ENLISTED
    '                blAllConstants += ml_POINT_PER_OFFICER
    '                blAllConstants += ml_POINT_PER_MIN1
    '                blAllConstants += ml_POINT_PER_MIN2
    '                blAllConstants += ml_POINT_PER_MIN3
    '                blAllConstants += ml_POINT_PER_MIN4
    '                blAllConstants += ml_POINT_PER_MIN5
    '                blAllConstants += ml_POINT_PER_MIN6

    '                Dim decAdd As Decimal = 0
    '                If iTechTypeID = ObjectType.eArmorTech Then
    '                    decAdd = CType(Me, ArmorTechComputer).GetThresholdModifier()
    '                Else
    '                    decAdd = decThreshold - decMaxThreshold
    '                    If iTechTypeID = ObjectType.eWeaponTech Then decAdd *= 1000000D
    '                End If
    '                decAdd *= CDec(blAllConstants)
    '                If CDec(blPoints + decAdd) > Int64.MaxValue Then blPoints = Int64.MaxValue Else blPoints = CLng(blPoints + decAdd)
    '            End If
    '        End If


    '        'AddToTestString("Points: " & blPoints.ToString)

    '        ' 5) Take the sum of the remaining points and divide it by 1 point of each type not locked summed together rounded down
    '        Dim blPointSum As Int64 = 0
    '        Dim lModByVal As Int32 = 0
    '        If blPoints < 0 Then

    '            Dim lAdjusters As Int32 = 0

    '            If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 AndAlso lResultHull > 0 Then
    '                blPointSum += ml_POINT_PER_HULL
    '                lAdjusters += 1
    '            End If
    '            If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 AndAlso lResultPower > 0 Then
    '                blPointSum += ml_POINT_PER_POWER
    '                lAdjusters += 1
    '            End If
    '            If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 AndAlso blResultResCost > 0 Then
    '                blPointSum += ml_POINT_PER_RES_COST
    '                lAdjusters += 1
    '            End If
    '            If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 AndAlso blResultResTime > 0 Then
    '                blPointSum += ml_POINT_PER_RES_TIME
    '                lAdjusters += 1
    '            End If
    '            If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 AndAlso blResultProdCost > 0 Then
    '                blPointSum += ml_POINT_PER_PROD_COST
    '                lAdjusters += 1
    '            End If
    '            If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 AndAlso blResultProdTime > 0 Then
    '                blPointSum += ml_POINT_PER_PROD_TIME
    '                lAdjusters += 1
    '            End If
    '            If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 AndAlso lResultColonists > 0 Then
    '                blPointSum += ml_POINT_PER_COLONIST
    '                lAdjusters += 1
    '            End If
    '            If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 AndAlso lResultEnlisted > 0 Then
    '                blPointSum += ml_POINT_PER_ENLISTED
    '                lAdjusters += 1
    '            End If
    '            If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 AndAlso lResultOfficers > 0 Then
    '                blPointSum += ml_POINT_PER_OFFICER
    '                lAdjusters += 1
    '            End If

    '            If bMin1Locked = False AndAlso lResultMin1 > 0 Then
    '                blPointSum += ml_POINT_PER_MIN1
    '                lAdjusters += 1
    '            End If
    '            If bMin2Locked = False AndAlso lResultMin2 > 0 Then
    '                blPointSum += ml_POINT_PER_MIN2
    '                lAdjusters += 1
    '            End If
    '            If bMin3Locked = False AndAlso lResultMin3 > 0 Then
    '                blPointSum += ml_POINT_PER_MIN3
    '                lAdjusters += 1
    '            End If
    '            If bMin4Locked = False AndAlso lResultMin4 > 0 Then
    '                blPointSum += ml_POINT_PER_MIN4
    '                lAdjusters += 1
    '            End If
    '            If bMin5Locked = False AndAlso lResultMin5 > 0 Then
    '                blPointSum += ml_POINT_PER_MIN5
    '                lAdjusters += 1
    '            End If
    '            If bMin6Locked = False AndAlso lResultMin6 > 0 Then
    '                blPointSum += ml_POINT_PER_MIN6
    '                lAdjusters += 1
    '            End If

    '            If blPointSum > 0 Then

    '                'AddToTestString("PointSum: " & blPointSum.ToString)

    '                ' 6) The result is applied to all non-locked values
    '                'Dim blTemp As Int64 = Math.Abs(blPoints) \ blPointSum
    '                Dim blTemp As Int64 = CLng(Math.Log10(blPoints) * blPoints) \ blPointSum 'CLng(Math.Pow(blPoints, 1 / 0.9)) \ blPointSum
    '                Dim blBaseAdjust As Int64 = blPoints \ blPointSum

    '                Dim lAddPts As Int32 = 0
    '                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
    '                If blTemp < 0 Then blTemp = 0
    '                lAddPts = CInt(blTemp)
    '                If lAddPts > 0 Then
    '                    If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 AndAlso lResultHull > lAddPts Then
    '                        lResultHull -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 AndAlso lResultPower > lAddPts Then
    '                        lAdjusters -= 1
    '                        lResultPower -= lAddPts
    '                    End If
    '                    If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 AndAlso blResultResCost > lAddPts Then
    '                        blResultResCost -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 AndAlso blResultResTime > lAddPts Then
    '                        blResultResTime -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 AndAlso blResultProdCost > lAddPts Then
    '                        blResultProdCost -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 AndAlso blResultProdTime > lAddPts Then
    '                        blResultProdTime -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 AndAlso lResultColonists > lAddPts Then
    '                        lResultColonists -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 AndAlso lResultEnlisted > lAddPts Then
    '                        lResultEnlisted -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 AndAlso lResultOfficers > lAddPts Then
    '                        lResultOfficers -= lAddPts
    '                        lAdjusters -= 1
    '                    End If

    '                    If bMin1Locked = False AndAlso lResultMin1 > lAddPts Then
    '                        lResultMin1 -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bMin2Locked = False AndAlso lResultMin2 > lAddPts Then
    '                        lResultMin2 -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bMin3Locked = False AndAlso lResultMin3 > lAddPts Then
    '                        lResultMin3 -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bMin4Locked = False AndAlso lResultMin4 > lAddPts Then
    '                        lResultMin4 -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bMin5Locked = False AndAlso lResultMin5 > lAddPts Then
    '                        lResultMin5 -= lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bMin6Locked = False AndAlso lResultMin6 > lAddPts Then
    '                        lResultMin6 -= lAddPts
    '                        lAdjusters -= 1
    '                    End If

    '                    If lAdjusters <> 0 Then Return False

    '                    blPoints += blBaseAdjust * blPointSum
    '                End If
    '            End If

    '            lModByVal = -1
    '        ElseIf blPoints > 0 Then
    '            Dim lAdjusters As Int32 = 0

    '            If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 Then
    '                blPointSum += ml_POINT_PER_HULL
    '                lAdjusters += 1
    '            End If
    '            If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 Then
    '                blPointSum += ml_POINT_PER_POWER
    '                lAdjusters += 1
    '            End If
    '            If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 Then
    '                blPointSum += ml_POINT_PER_RES_COST
    '                lAdjusters += 1
    '            End If
    '            If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 Then
    '                blPointSum += ml_POINT_PER_RES_TIME
    '                lAdjusters += 1
    '            End If
    '            If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 Then
    '                blPointSum += ml_POINT_PER_PROD_COST
    '                lAdjusters += 1
    '            End If
    '            If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 Then
    '                blPointSum += ml_POINT_PER_PROD_TIME
    '                lAdjusters += 1
    '            End If
    '            If ml_POINT_PER_MIN1 > 0 AndAlso bMin1Locked = False Then
    '                blPointSum += ml_POINT_PER_MIN1
    '                lAdjusters += 1
    '            End If
    '            If ml_POINT_PER_MIN2 > 0 AndAlso bMin2Locked = False Then
    '                blPointSum += ml_POINT_PER_MIN2
    '                lAdjusters += 1
    '            End If
    '            If ml_POINT_PER_MIN3 > 0 AndAlso bMin3Locked = False Then
    '                blPointSum += ml_POINT_PER_MIN3
    '                lAdjusters += 1
    '            End If
    '            If ml_POINT_PER_MIN4 > 0 AndAlso bMin4Locked = False Then
    '                blPointSum += ml_POINT_PER_MIN4
    '                lAdjusters += 1
    '            End If
    '            If ml_POINT_PER_MIN5 > 0 AndAlso bMin5Locked = False Then
    '                blPointSum += ml_POINT_PER_MIN5
    '                lAdjusters += 1
    '            End If
    '            If ml_POINT_PER_MIN6 > 0 AndAlso bMin6Locked = False Then
    '                blPointSum += ml_POINT_PER_MIN6
    '                lAdjusters += 1
    '            End If

    '            If blPointSum > 0 Then
    '                ' AddToTestString("PointSum: " & blPointSum.ToString)

    '                ' 6) The result is applied to all non-locked values
    '                'Dim blTemp As Int64 = blPoints \ blPointSum
    '                Dim blTemp As Int64 = CLng(Math.Log10(blPoints) * blPoints) \ blPointSum 'CLng(Math.Pow(blPoints, 1 / 0.9)) \ blPointSum
    '                Dim blBaseAdjust As Int64 = blPoints \ blPointSum

    '                Dim lAddPts As Int32 = 0
    '                'If blTemp > Int32.MaxValue Then lAddPts = Int32.MaxValue
    '                'If blTemp < 0 Then lAddPts = 0
    '                If blTemp > 0 Then
    '                    If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 Then
    '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_HULL
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultHull += lAddPts 'lBaseHull += lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 Then
    '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_HULL
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultPower += lAddPts 'lBaseHull += lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 Then
    '                        blResultResCost += blTemp ' \ ml_POINT_PER_RES_COST 'blBaseResCost += blTemp \ ml_POINT_PER_RES_COST
    '                        lAdjusters -= 1
    '                    End If
    '                    If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 Then
    '                        blResultResTime += blTemp '\ ml_POINT_PER_RES_TIME
    '                        lAdjusters -= 1
    '                    End If
    '                    If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 Then
    '                        blResultProdCost += blTemp '\ ml_POINT_PER_PROD_COST
    '                        lAdjusters -= 1
    '                    End If
    '                    If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 Then
    '                        blResultProdTime += blTemp '\ ml_POINT_PER_PROD_TIME
    '                        lAdjusters -= 1
    '                    End If

    '                    If ml_POINT_PER_MIN1 > 0 AndAlso bMin1Locked = False AndAlso ml_POINT_PER_MIN1 > 0 Then
    '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_DBody
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultMin1 += lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If ml_POINT_PER_MIN2 > 0 AndAlso bMin2Locked = False AndAlso ml_POINT_PER_MIN2 > 0 Then
    '                        Dim blAdd As Int64 = blTemp ' \ ml_POINT_PER_DFrame
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultMin2 += lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If ml_POINT_PER_MIN3 > 0 AndAlso bMin3Locked = False AndAlso ml_POINT_PER_MIN3 > 0 Then
    '                        Dim blAdd As Int64 = blTemp ' \ ml_POINT_PER_DMeld
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultMin3 += lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If ml_POINT_PER_MIN4 > 0 AndAlso bMin4Locked = False AndAlso ml_POINT_PER_MIN4 > 0 Then
    '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_SBody
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultMin4 += lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If ml_POINT_PER_MIN5 > 0 AndAlso bMin5Locked = False AndAlso ml_POINT_PER_MIN5 > 0 Then
    '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_SFrame
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultMin5 += lAddPts
    '                        lAdjusters -= 1
    '                    End If
    '                    If ml_POINT_PER_MIN6 > 0 AndAlso bMin6Locked = False AndAlso ml_POINT_PER_MIN6 > 0 Then
    '                        Dim blAdd As Int64 = blTemp ' \ ml_POINT_PER_SMeld
    '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
    '                        lResultMin6 += lAddPts
    '                        lAdjusters -= 1
    '                    End If

    '                    If lAdjusters <> 0 Then Return False

    '                    blPoints -= blBaseAdjust * blPointSum
    '                End If
    '            End If

    '            lModByVal = 1
    '        End If

    '        If blPoints > 0 Then
    '            blPointSum = 0
    '            If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 Then blPointSum += ml_POINT_PER_COLONIST
    '            If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 Then blPointSum += ml_POINT_PER_ENLISTED
    '            If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 Then blPointSum += ml_POINT_PER_OFFICER

    '            If blPointSum > 0 Then
    '                Dim blTemp As Int64 = blPoints \ blPointSum
    '                Dim lAddPts As Int32 = 0
    '                If blTemp > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blTemp)
    '                If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 Then lResultColonists += lAddPts
    '                If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 Then lResultEnlisted += lAddPts
    '                If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 Then lResultOfficers += lAddPts

    '                blPoints -= CLng(blTemp) * blPointSum
    '            End If
    '        End If

    '        ' 7) The remainder points are applied to the non-locked items until the remainder < 0
    '        Dim lCnt As Int32 = 0
    '        While blPoints > 0
    '            Dim blPrevPoints As Int64 = blPoints
    '            If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_HULL, lResultHull, lModByVal)
    '            If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_POWER, lResultPower, lModByVal)
    '            If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_RES_COST, blResultResCost, lModByVal)
    '            If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_RES_TIME, blResultResTime, lModByVal)
    '            If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_PROD_COST, blResultProdCost, lModByVal)
    '            If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_PROD_TIME, blResultProdTime, lModByVal)
    '            If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_COLONIST, lResultColonists, lModByVal)
    '            If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_ENLISTED, lResultEnlisted, lModByVal)
    '            If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_OFFICER, lResultOfficers, lModByVal)
    '            If bMin1Locked = False AndAlso ml_POINT_PER_MIN1 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN1, lResultMin1, lModByVal)
    '            If bMin2Locked = False AndAlso ml_POINT_PER_MIN2 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN2, lResultMin2, lModByVal)
    '            If bMin3Locked = False AndAlso ml_POINT_PER_MIN3 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN3, lResultMin3, lModByVal)
    '            If bMin4Locked = False AndAlso ml_POINT_PER_MIN4 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN4, lResultMin4, lModByVal)
    '            If bMin5Locked = False AndAlso ml_POINT_PER_MIN5 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN5, lResultMin5, lModByVal)
    '            If bMin6Locked = False AndAlso ml_POINT_PER_MIN6 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN6, lResultMin6, lModByVal)
    '            ' 8) If the remainder does not reach 0 after several attempts, break out
    '            If blPrevPoints = blPoints Then Exit While
    '            lCnt += 1
    '            If lCnt > 100 Then Exit While
    '        End While

    '        'frmBuildCost.SetBaseValues(lResultHull, lResultPower, blResultResCost, blResultResTime, blResultProdCost, blResultProdTime, lResultColonists, lResultEnlisted, lResultOfficers)


    '        'AddToTestString("RESULTS: ")
    '        'AddToTestString(" Remaining Points: " & blPoints.ToString)
    '        'AddToTestString(" Hull: " & lResultHull.ToString)
    '        'AddToTestString(" Power: " & lResultPower.ToString)
    '        'AddToTestString(" ResCost: " & blResultResCost.ToString)
    '        'AddToTestString(" ResTime: " & blResultResTime.ToString)
    '        'AddToTestString(" ProdCost: " & blResultProdCost.ToString)
    '        'AddToTestString(" ProdTime: " & blResultProdTime.ToString)
    '        'AddToTestString(" Colonist: " & lResultColonists.ToString)
    '        'AddToTestString(" Enlisted: " & lResultEnlisted.ToString)
    '        'AddToTestString(" Officer: " & lResultOfficers.ToString)
    '        'AddToTestString(" Min1: " & lResultMin1.ToString)
    '        'AddToTestString(" Min2: " & lResultMin2.ToString)
    '        'AddToTestString(" Min3: " & lResultMin3.ToString)
    '        'AddToTestString(" Min4: " & lResultMin4.ToString)
    '        'AddToTestString(" Min5: " & lResultMin5.ToString)
    '        'AddToTestString(" Min6: " & lResultMin6.ToString)
    '        Dim decIntelMult As Decimal = 160D / lPopIntel
    '        blResultResCost = CLng(blResultResCost * decIntelMult)
    '        blResultResTime = CLng(blResultResTime * decIntelMult)

    '        If Math.Abs(blPoints) > 1000 Then
    '            'lblDesignFlaw.Visible = True
    '            'lblDesignFlaw.Caption = "Your scientists believe this design to be possible."
    '            Return False
    '        End If

    '        Return True
    '    Catch
    '        'lblDesignFlaw.Visible = True
    '        'lblDesignFlaw.Caption = "This design is impossible to comprehend."
    '        Return False
    '    End Try

    'End Function

    'Public lSumAllDA As Int32
    'Private Function GetMaxDA(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32

    '    Dim oMineral As Mineral = GetEpicaMineral(lMineralID)
    '    If oMineral Is Nothing Then Return 0

    '    Dim lTempMax As Int32 = 0
    '    If muPropReqs Is Nothing = False Then
    '        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    '            If muPropReqs(X).lMinID = lMineralIndex Then
    '                Dim lA As Int32 = oMineral.GetPropertyValue(muPropReqs(X).lPropID) 'oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))

    '                lTempMax = Math.Max(Math.Abs(muPropReqs(X).lPropVal - lA), lTempMax)

    '                'Now, for our D-A Sum
    '                Dim lClientD As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(muPropReqs(X).lPropVal)
    '                Dim lClientA As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lA)
    '                lSumAllDA += Math.Abs(lClientD - lClientA)

    '            End If
    '        Next X
    '    End If
    '    Dim lMaxDA As Int32 = lTempMax

    '    ''TODO: Remove me when ready!!!
    '    'If oMineral.ObjectID = 1 Then
    '    '    lMaxDA = 5 * 23
    '    'ElseIf oMineral.ObjectID = 5 Then
    '    '    lMaxDA = 23
    '    'ElseIf oMineral.ObjectID = 9 Then
    '    '    lMaxDA = 255
    '    'End If

    '    Return lMaxDA
    'End Function

    'Private Function GetSumClientDA(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32

    '    Dim oMineral As Mineral = GetEpicaMineral(lMineralID)
    '    If oMineral Is Nothing Then Return 0

    '    Dim lResult As Int32 = 0

    '    If muPropReqs Is Nothing = False Then
    '        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    '            If muPropReqs(X).lMinID = lMineralIndex Then
    '                Dim lA As Int32 = oMineral.GetPropertyValue(muPropReqs(X).lPropID) 'oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))

    '                'Now, for our D-A Sum
    '                Dim lClientD As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(muPropReqs(X).lPropVal)
    '                Dim lClientA As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lA)
    '                lResult += Math.Abs(lClientD - lClientA)

    '            End If
    '        Next X
    '    End If
    '    If lResult < 1 Then lResult = 1

    '    Return lResult
    'End Function
  

    Public Function GetDADiff(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32
        Dim oMineral As Mineral = GetEpicaMineral(lMineralID)
        If oMineral Is Nothing Then Return 0
        If oMineral.ObjectID = 157 Then Return 0

        Dim lTempMax As Int32 = 0
        Dim lResult As Int32 = 0
        If muPropReqs Is Nothing Then SetMinReqProps()

        If muPropReqs Is Nothing = False Then
            For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
                If muPropReqs(X).lMinID = lMineralIndex Then
                    Dim lA As Int32 = oMineral.GetPropertyValue(muPropReqs(X).lPropID) 'oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))

                    lResult += Math.Abs(muPropReqs(X).lPropVal - lA)
                End If
            Next X
        End If

        If lResult < 0 Then Return 0
        Return lResult
    End Function
    Private Function GetClientDADiff(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32
        Dim oMineral As Mineral = GetEpicaMineral(lMineralID)
        If oMineral Is Nothing Then Return 0
        If oMineral.ObjectID = 157 Then Return 0

        Dim lTempMax As Int32 = 0
        Dim lResult As Int32 = 0
        If muPropReqs Is Nothing = False Then
            For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
                If muPropReqs(X).lMinID = lMineralIndex Then
                    Dim lA As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(oMineral.GetPropertyValue(muPropReqs(X).lPropID))
                    Dim lD As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(muPropReqs(X).lPropID)

                    lResult += Math.Abs(lD - lA)
                End If
            Next X
        End If

        If lResult < 0 Then Return 0
        Return lResult
    End Function

    Private Function GetCoeffValue(ByVal lLookup As elPropCoeffLookup) As Decimal

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
                decTemp -= 0.01D
                If decTemp < 0 Then decTemp = 0.01D
        End Select

        Return decTemp
    End Function

    Public bRunCheckForMicro As Boolean = False
    Public Function BuilderCostValueChange(ByVal iTechTypeID As Int16, ByVal decMaxThreshold As Decimal, ByVal lPopIntel As Int32, ByVal lHullToResidence As Int32) As Boolean 'ByRef lblDesignFlaw As UILabel, ByRef frmBuildCost As frmTechBuilderCost, ByRef tpMin1 As ctlTechProp, ByRef tpMin2 As ctlTechProp, ByRef tpMin3 As ctlTechProp, ByRef tpMin4 As ctlTechProp, ByRef tpMin5 As ctlTechProp, ByRef tpMin6 As ctlTechProp, ByVal iTechTypeID As Int16, ByVal decMaxThreshold As Decimal)
        Try
            SetMinReqProps()

            Dim decBill As Decimal = GetTheBill()

            'ok, now we have the bill... let's get the locked payment
            Dim bColLockedAtStart As Boolean = lLockedColonists <> -1
            Dim bEnlLockedAtStart As Boolean = lLockedEnlisted <> -1
            Dim bOffLockedAtStart As Boolean = lLockedOfficers <> -1

            'Also need to get the locked minerals
            Dim bMin1Locked, bMin2Locked, bMin3Locked, bMin4Locked, bMin5Locked, bMin6Locked As Boolean
            If lLockedMin1 <> -1 Then bMin1Locked = True
            If lLockedMin2 <> -1 Then bMin2Locked = True
            If lLockedMin3 <> -1 Then bMin3Locked = True
            If lLockedMin4 <> -1 Then bMin4Locked = True
            If lLockedMin5 <> -1 Then bMin5Locked = True
            If lLockedMin6 <> -1 Then bMin6Locked = True

            Dim decHullPayment As Decimal = 0D
            Dim decPowerPayment As Decimal = 0D
            Dim decProdCostPayment As Decimal = 0D
            Dim decProdTimePayment As Decimal = 0D

            Dim decPayment As Decimal = 0
            Dim lDistributable As Int32 = 0
            Dim decDistVals(elPropCoeffLookup.eMin6) As Decimal
            For X As Int32 = 0 To elPropCoeffLookup.eMin6
                decDistVals(X) = 0D
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
                decPayment += CDec(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eResCost) = GetCoeffValue(elPropCoeffLookup.eResCost)
            End If
            If blLockedResTime <> -1 Then
                decPayment += CDec(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
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
                    decPayment += CDec(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eColonist) = GetCoeffValue(elPropCoeffLookup.eColonist)
                End If
            Else : lLockedColonists = 0
            End If

            If GetCoeffValue(elPropCoeffLookup.eEnlisted) <> 0 Then
                If lLockedEnlisted <> -1 Then
                    decPayment += CDec(Math.Pow(lLockedEnlisted, GetCoeffValue(elPropCoeffLookup.eEnlisted)))
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eEnlisted) = GetCoeffValue(elPropCoeffLookup.eEnlisted)
                End If
            Else : lLockedEnlisted = 0
            End If

            If GetCoeffValue(elPropCoeffLookup.eOfficer) <> 0 Then
                If lLockedOfficers <> -1 Then
                    decPayment += CDec(Math.Pow(lLockedOfficers, GetCoeffValue(elPropCoeffLookup.eOfficer)))
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eOfficer) = GetCoeffValue(elPropCoeffLookup.eOfficer)
                End If
            Else : lLockedOfficers = 0
            End If

            Dim lTotalDAOff As Int32 = 0
            If bMin1Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin1) = GetCoeffValue(elPropCoeffLookup.eMin1)
            End If
            If bMin2Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin2) = GetCoeffValue(elPropCoeffLookup.eMin2)
            End If
            If bMin3Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin3) = GetCoeffValue(elPropCoeffLookup.eMin3)
            End If
            If bMin4Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin4) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin4) = GetCoeffValue(elPropCoeffLookup.eMin4)
            Else
                bMin4Locked = True
                lLockedMin4 = 0
            End If
            If bMin5Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin5) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin5) = GetCoeffValue(elPropCoeffLookup.eMin5)
            Else
                bMin5Locked = True
                lLockedMin5 = 0
            End If
            If bMin6Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin6) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin6) = GetCoeffValue(elPropCoeffLookup.eMin6)
            Else
                bMin6Locked = True
                lLockedMin6 = 0
            End If

            If lMineral1ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral1ID, 0)
            If lMineral2ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral2ID, 1)
            If lMineral3ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral3ID, 2)
            If lMineral4ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral4ID, 3)
            If lMineral5ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral5ID, 4)
            If lMineral6ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral6ID, 5)
            'decBill *= Math.Max(1, lTotalDAOff)

            'ok, we have our payment and the bill...
            If decPayment < decBill Then
                Dim lResDistCnt As Int32 = 0
                If blLockedResCost = -1 Then lResDistCnt += 1
                If blLockedResTime = -1 Then lResDistCnt += 1
                If lLockedColonists = -1 Then lResDistCnt += 1
                If lResDistCnt <> lDistributable Then
                    If blLockedResCost = -1 Then
                        blLockedResCost = 100000000 'CLng(Math.Pow(100000000, GetCoeffValue(elPropCoeffLookup.eResCost)))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
                        If blNew + decPayment > decBill Then
                            blLockedResCost = -1
                        Else
                            decPayment += blNew
                            decDistVals(elPropCoeffLookup.eResCost) = 0
                            lDistributable -= 1
                        End If
                    End If
                    If blLockedResTime = -1 Then
                        blLockedResTime = 2000000000 'CLng(Math.Pow(2000000000, GetCoeffValue(elPropCoeffLookup.eResTime)))
                        Dim blNew As Int64 = CLng(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
                        If blNew + decPayment > decBill Then
                            blLockedResTime = -1
                        Else
                            decPayment += blNew
                            decDistVals(elPropCoeffLookup.eResTime) = 0
                            lDistributable -= 1
                        End If
                    End If
                    If lLockedColonists = -1 Then
                        lLockedColonists = 20 'CInt(Math.Pow(decDistVals(elPropCoeffLookup.eColonist) * decDiffTemp, (1 / GetCoeffValue(elPropCoeffLookup.eColonist)))) \ 20
                        Dim blNew As Int64 = CLng(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                        If blNew + decPayment > decBill Then
                            lLockedColonists = -1
                        Else
                            decPayment += blNew
                            decDistVals(elPropCoeffLookup.eColonist) = 0
                            lDistributable -= 1
                        End If


                    End If
                End If

                'ok, we owe more
                Dim decDiff As Decimal = decBill - decPayment
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

                    'Dim blMorePay As Int64 = 0

                    'ok, take decDiff / lDistributable = Points to distribute to each unlocked property (lPointToDist)
                    ' each unlocked property is then raised by lPointToDist ^ (1 / GetCoeff)
                    'Dim blPointsToDistribute As Int64 = CLng(decDiff / lDistributable)
                    If lLockedHull = -1 Then
                        lResultHull = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eHull) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eHull))))
                        decHullPayment = CDec(Math.Pow(lResultHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                    Else : lResultHull = lLockedHull
                    End If
                    If lLockedPower = -1 AndAlso GetCoeffValue(elPropCoeffLookup.ePower) <> 0 Then
                        lResultPower = CInt(Math.Pow(decDistVals(elPropCoeffLookup.ePower) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.ePower))))
                        decPowerPayment = CDec(Math.Pow(lResultPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                    Else : lResultPower = lLockedPower
                    End If
                    If blLockedResCost = -1 Then
                        blResultResCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResCost))))
                    Else : blResultResCost = blLockedResCost
                    End If
                    If blLockedResTime = -1 Then
                        blResultResTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResTime))))
                    Else : blResultResTime = blLockedResTime
                    End If
                    If blLockedProdCost = -1 Then
                        blResultProdCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdCost))))
                        decProdCostPayment = CDec(Math.Pow(blResultProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
                    Else : blResultProdCost = blLockedProdCost
                    End If
                    If blLockedProdTime = -1 Then
                        blResultProdTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdTime))))
                        decProdTimePayment = CDec(Math.Pow(blResultProdTime, GetCoeffValue(elPropCoeffLookup.eProdTime)))
                    Else : blResultProdTime = blLockedProdTime
                    End If
                    If lLockedColonists = -1 Then
                        lResultColonists = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eColonist) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eColonist))))
                    Else : lResultColonists = lLockedColonists
                    End If
                    If lLockedEnlisted = -1 Then
                        lResultEnlisted = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eEnlisted) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eEnlisted))))
                    Else : lResultEnlisted = lLockedEnlisted
                    End If
                    If lLockedOfficers = -1 Then
                        lResultOfficers = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eOfficer) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eOfficer))))
                    Else : lResultOfficers = lLockedOfficers
                    End If
                    If bMin1Locked = False Then
                        lResultMin1 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin1) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin1))))
                    Else : lResultMin1 = lLockedMin1
                    End If
                    If bMin2Locked = False Then
                        lResultMin2 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin2) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin2))))
                    Else : lResultMin2 = lLockedMin2
                    End If
                    If bMin3Locked = False Then
                        lResultMin3 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin3) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin3))))
                    Else : lResultMin3 = lLockedMin3
                    End If
                    If bMin4Locked = False Then
                        lResultMin4 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin4) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin4))))
                    Else : lResultMin4 = lLockedMin4
                    End If
                    If bMin5Locked = False Then
                        lResultMin5 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin5) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin5))))
                    Else : lResultMin5 = lLockedMin5
                    End If
                    If bMin6Locked = False Then
                        lResultMin6 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin6) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin6))))
                    Else : lResultMin6 = lLockedMin6
                    End If

                    'Now, check our payments
                    bIsMicroTech = False
                    'Components will no longer be marked as micro tech
                    'If bRunCheckForMicro = True Then
                    'If GetCoeffValue(elPropCoeffLookup.ePower) <> 0 AndAlso decPowerPayment < decBill * 0.005D Then
                    '    bIsMicroTech = True
                    'ElseIf GetCoeffValue(elPropCoeffLookup.eHull) <> 0 AndAlso decHullPayment < decBill * 0.005D Then
                    '    bIsMicroTech = True
                    'ElseIf GetCoeffValue(elPropCoeffLookup.eProdCost) <> 0 AndAlso decProdCostPayment < decBill * 0.01D Then
                    '    bIsMicroTech = True
                    'ElseIf GetCoeffValue(elPropCoeffLookup.eProdTime) <> 0 AndAlso decProdTimePayment < decBill * 0.05D Then
                    '    bIsMicroTech = True
                    'End If
                    'End If

                    ''Sum Minerals AV > Hull Size AV
                    'If lLockedMin1 + lLockedMin2 + lLockedMin3 + lLockedMin4 + lLockedMin5 + lLockedMin6 < lLockedHull Then Return False

                    ''HullSize PP + Power PP > 10% of total points to be paid
                    'Dim decHull As Decimal = CDec(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                    'Dim decPower As Decimal = CDec(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                    'If decHull + decPower < decBill * 0.1D Then Return False

                    ''HullSize PP + Minerals PP > 15%
                    'Dim decMin1 As Decimal = CDec(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
                    'Dim decMin2 As Decimal = CDec(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
                    'Dim decMin3 As Decimal = CDec(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
                    'Dim decMin4 As Decimal = CDec(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
                    'Dim decMin5 As Decimal = CDec(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
                    'Dim decMin6 As Decimal = CDec(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))
                    'If decHull + decMin1 + decMin2 + decMin3 + decMin4 + decMin5 + decMin6 < decBill * 0.15D Then Return False

                    ''Col PP + Enl PP + Off PP > 10%
                    'Dim decCol As Decimal = CDec(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                    'Dim decEnl As Decimal = CDec(Math.Pow(lLockedEnlisted, GetCoeffValue(elPropCoeffLookup.eEnlisted)))
                    'Dim decOff As Decimal = CDec(Math.Pow(lLockedOfficers, GetCoeffValue(elPropCoeffLookup.eOfficer)))
                    ''If decCol + decEnl + decOff < decBill * 0.05D Then
                    ''    bImpossibleDesign = True
                    ''    lblDesignFlaw.Caption = "Not enough crew for this design."
                    ''    lblDesignFlaw.Visible = True
                    ''End If

                    ''ResTime PP + ProdTime PP > 30% of total points to be paid
                    ''ResTime PP + ProdTime PP < 60% of total points to be paid
                    'Dim decResTime As Decimal = CDec(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
                    'Dim decProdTime As Decimal = CDec(Math.Pow(blLockedProdTime, GetCoeffValue(elPropCoeffLookup.eProdTime)))
                    'If decResTime + decProdTime < decBill * 0.1D Then Return False
                    'If decResTime + decProdTime > decBill * 0.9D Then Return False

                    ''ResCost PP + ProdCost PP > 5% < 50%
                    'Dim decResCost As Decimal = CDec(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
                    'Dim decProdCost As Decimal = CDec(Math.Pow(blLockedProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
                    'If decResCost + decProdCost < 0.05 * decBill Then Return False
                    'If decProdCost + decResCost > 0.5D * decBill Then Return False

                    ''HullRequired AV > Col AV + Enl AV + Off AV
                    'If lLockedHull < lLockedColonists + lLockedEnlisted + lLockedOfficers Then Return False

                    ''Enl AV + Off AV > Col AV
                    'If lLockedEnlisted + lLockedOfficers < lLockedColonists Then Return False

                    ''Enl AV < Off AV * 5
                    'If lLockedEnlisted > lLockedOfficers * 5 Then Return False

                    ''HullRequired AV > Crew Hull Size AV
                    'If lLockedHull < (lLockedOfficers + lLockedColonists + lLockedEnlisted) * lHullToResidence Then Return False

                    'Dim decBill80 As Decimal = decBill * 0.8D
                    'If decResCost > decBill80 OrElse decResTime > decBill80 OrElse decProdCost > decBill80 OrElse decProdTime > decBill80 OrElse decMin1 > decBill80 OrElse decMin2 > decBill80 Then
                    '    Return False
                    'ElseIf decMin3 > decBill80 OrElse decMin4 > decBill80 OrElse decMin5 > decBill80 OrElse decMin6 > decBill80 OrElse decHull > decBill80 OrElse decCol > decBill80 OrElse decEnl > decBill80 OrElse decOff > decBill80 Then
                    '    Return False
                    'End If

                    If bColLockedAtStart = True Then lResultColonists = Math.Max(1, lResultColonists)
                    If bEnlLockedAtStart = True Then lResultEnlisted = Math.Max(1, lResultEnlisted)
                    If bOffLockedAtStart = True Then lResultOfficers = Math.Max(1, lResultOfficers)
                    lResultHull = Math.Max(1, lResultHull)
                    If iTechTypeID <> ObjectType.eEngineTech Then lResultPower = Math.Max(lResultPower, 1)
                    If lMineral1ID > 0 Then lResultMin1 = Math.Max(1, lResultMin1)
                    If lMineral2ID > 0 Then lResultMin2 = Math.Max(1, lResultMin2)
                    If lMineral3ID > 0 Then lResultMin3 = Math.Max(1, lResultMin3)
                    If lMineral4ID > 0 Then lResultMin4 = Math.Max(1, lResultMin4)
                    If lMineral5ID > 0 Then lResultMin5 = Math.Max(1, lResultMin5)
                    If lMineral6ID > 0 Then lResultMin6 = Math.Max(1, lResultMin6)
                    blResultProdCost = Math.Max(1, blResultProdCost)
                    blResultProdTime = Math.Max(1, blResultProdTime)
                    blResultResCost = Math.Max(1, blResultResCost)
                    blResultResTime = Math.Max(1, blResultResTime)

                Else
                    Return False
                End If

            Else
                Return False
            End If
        Catch
            Return False
        End Try

        Return True

    End Function


    Public Function IsDesignImpossible(ByVal iTechTypeID As Int16, ByVal oOwner As Player) As Boolean
        Try
            SetMinReqProps()
            Dim decTotalBill As Decimal = GetTheBill()

            'ok, now we have the bill... let's get the locked payment
            Dim bColLockedAtStart As Boolean = lLockedColonists <> -1
            Dim bEnlLockedAtStart As Boolean = lLockedEnlisted <> -1
            Dim bOffLockedAtStart As Boolean = lLockedOfficers <> -1

            'Also need to get the locked minerals
            Dim bMin1Locked, bMin2Locked, bMin3Locked, bMin4Locked, bMin5Locked, bMin6Locked As Boolean
            If lLockedMin1 <> -1 Then bMin1Locked = True
            If lLockedMin2 <> -1 Then bMin2Locked = True
            If lLockedMin3 <> -1 Then bMin3Locked = True
            If lLockedMin4 <> -1 Then bMin4Locked = True
            If lLockedMin5 <> -1 Then bMin5Locked = True
            If lLockedMin6 <> -1 Then bMin6Locked = True


            Dim decHullPayment As Decimal = 0D
            Dim decPowerPayment As Decimal = 0D
            Dim decProdCostPayment As Decimal = 0D
            Dim decProdTimePayment As Decimal = 0D

            Dim decPayment As Decimal = 0
            Dim lDistributable As Int32 = 0
            Dim decDistVals(elPropCoeffLookup.eMin6) As Decimal
            For X As Int32 = 0 To elPropCoeffLookup.eMin6
                decDistVals(X) = 0D
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
                decPayment += CDec(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eResCost) = GetCoeffValue(elPropCoeffLookup.eResCost)
            End If
            If blLockedResTime <> -1 Then
                decPayment += CDec(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
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
                    decPayment += CDec(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eColonist) = GetCoeffValue(elPropCoeffLookup.eColonist)
                End If
            Else : lLockedColonists = 0
            End If

            If GetCoeffValue(elPropCoeffLookup.eEnlisted) <> 0 Then
                If lLockedEnlisted <> -1 Then
                    decPayment += CDec(Math.Pow(lLockedEnlisted, GetCoeffValue(elPropCoeffLookup.eEnlisted)))
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eEnlisted) = GetCoeffValue(elPropCoeffLookup.eEnlisted)
                End If
            Else : lLockedEnlisted = 0
            End If

            If GetCoeffValue(elPropCoeffLookup.eOfficer) <> 0 Then
                If lLockedOfficers <> -1 Then
                    decPayment += CDec(Math.Pow(lLockedOfficers, GetCoeffValue(elPropCoeffLookup.eOfficer)))
                Else
                    lDistributable += 1
                    decDistVals(elPropCoeffLookup.eOfficer) = GetCoeffValue(elPropCoeffLookup.eOfficer)
                End If
            Else : lLockedOfficers = 0
            End If

            Dim lTotalDAOff As Int32 = 0
            If bMin1Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin1) = GetCoeffValue(elPropCoeffLookup.eMin1)
            End If
            If bMin2Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin2) = GetCoeffValue(elPropCoeffLookup.eMin2)
            End If
            If bMin3Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
            Else
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin3) = GetCoeffValue(elPropCoeffLookup.eMin3)
            End If
            If bMin4Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin4) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin4) = GetCoeffValue(elPropCoeffLookup.eMin4)
            Else
                bMin4Locked = True
                lLockedMin4 = 0
            End If
            If bMin5Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin5) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin5) = GetCoeffValue(elPropCoeffLookup.eMin5)
            Else
                bMin5Locked = True
                lLockedMin5 = 0
            End If
            If bMin6Locked = True Then
                decPayment += CDec(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))
            ElseIf GetCoeffValue(elPropCoeffLookup.eMin6) > 0 Then
                lDistributable += 1
                decDistVals(elPropCoeffLookup.eMin6) = GetCoeffValue(elPropCoeffLookup.eMin6)
            Else
                bMin6Locked = True
                lLockedMin6 = 0
            End If

            If lMineral1ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral1ID, 0)
            If lMineral2ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral2ID, 1)
            If lMineral3ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral3ID, 2)
            If lMineral4ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral4ID, 3)
            If lMineral5ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral5ID, 4)
            If lMineral6ID > 0 Then lTotalDAOff += GetClientDADiff(lMineral6ID, 5)

            If decPayment < decTotalBill AndAlso lDistributable <> 0 Then
                Dim decProdMin As Decimal = 0.01D
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
                            decPayment += blNew
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
                            decPayment += blNew
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
                            decPayment += blNew
                            decDistVals(elPropCoeffLookup.eColonist) = 0
                            lDistributable -= 1
                        End If


                    End If
                End If

                Dim decDiff As Decimal = decTotalBill - decPayment

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
                'If lLockedHull = -1 Then
                '    lResultHull = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eHull) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eHull))))
                '    decHullPayment = CDec(Math.Pow(lResultHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                'Else : lResultHull = lLockedHull
                'End If
                'If lLockedPower = -1 AndAlso GetCoeffValue(elPropCoeffLookup.ePower) <> 0 Then
                '    lResultPower = CInt(Math.Pow(decDistVals(elPropCoeffLookup.ePower) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.ePower))))
                '    decPowerPayment = CDec(Math.Pow(lResultPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                'Else : lResultPower = lLockedPower
                'End If
                'If blLockedResCost = -1 Then
                '    blResultResCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResCost))))
                'Else : blResultResCost = blLockedResCost
                'End If
                'If blLockedResTime = -1 Then
                '    blResultResTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResTime))))
                'Else : blResultResTime = blLockedResTime
                'End If
                'If blLockedProdCost = -1 Then
                '    blResultProdCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdCost))))
                '    decProdCostPayment = CDec(Math.Pow(blResultProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
                'Else : blResultProdCost = blLockedProdCost
                'End If
                'If blLockedProdTime = -1 Then
                '    blResultProdTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdTime))))
                '    decProdTimePayment = CDec(Math.Pow(blResultProdTime, GetCoeffValue(elPropCoeffLookup.eProdTime)))
                'Else : blResultProdTime = blLockedProdTime
                'End If
                'If lLockedColonists = -1 Then
                '    lResultColonists = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eColonist) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eColonist))))
                'Else : lResultColonists = lLockedColonists
                'End If
                'If lLockedEnlisted = -1 Then
                '    lResultEnlisted = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eEnlisted) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eEnlisted))))
                'Else : lResultEnlisted = lLockedEnlisted
                'End If
                'If lLockedOfficers = -1 Then
                '    lResultOfficers = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eOfficer) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eOfficer))))
                'Else : lResultOfficers = lLockedOfficers
                'End If
                'If bMin1Locked = False Then
                '    lResultMin1 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin1) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin1))))
                'Else : lResultMin1 = lLockedMin1
                'End If
                'If bMin2Locked = False Then
                '    lResultMin2 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin2) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin2))))
                'Else : lResultMin2 = lLockedMin2
                'End If
                'If bMin3Locked = False Then
                '    lResultMin3 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin3) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin3))))
                'Else : lResultMin3 = lLockedMin3
                'End If
                'If bMin4Locked = False Then
                '    lResultMin4 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin4) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin4))))
                'Else : lResultMin4 = lLockedMin4
                'End If
                'If bMin5Locked = False Then
                '    lResultMin5 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin5) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin5))))
                'Else : lResultMin5 = lLockedMin5
                'End If
                'If bMin6Locked = False Then
                '    lResultMin6 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin6) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin6))))
                'Else : lResultMin6 = lLockedMin6
                'End If
                Dim blNewPay As Int64 = 0 'CLng(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
                If lLockedHull = -1 Then
                    lLockedHull = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eHull) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eHull))))
                End If
                decHullPayment = CDec(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                If lLockedPower = -1 AndAlso GetCoeffValue(elPropCoeffLookup.ePower) <> 0 Then
                    lLockedPower = CInt(Math.Pow(decDistVals(elPropCoeffLookup.ePower) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.ePower))))
                    decPowerPayment = CDec(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                End If
                decPowerPayment = CDec(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                If blLockedResCost = -1 Then
                    blLockedResCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResCost))))
                End If
                Dim decResCostPayment As Decimal = CLng(Math.Pow(blLockedResCost, GetCoeffValue(elPropCoeffLookup.eResCost)))
                If blLockedResTime = -1 Then
                    blLockedResTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eResTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eResTime))))
                End If
                Dim decResTimePayment As Decimal = CLng(Math.Pow(blLockedResTime, GetCoeffValue(elPropCoeffLookup.eResTime)))
                If blLockedProdCost = -1 Then
                    blLockedProdCost = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdCost) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdCost))))
                End If
                decProdCostPayment = CDec(Math.Pow(blLockedProdCost, GetCoeffValue(elPropCoeffLookup.eProdCost)))
                If blLockedProdTime = -1 Then
                    blLockedProdTime = CLng(Math.Pow(decDistVals(elPropCoeffLookup.eProdTime) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eProdTime))))
                End If
                decProdTimePayment = CDec(Math.Pow(blLockedProdTime, GetCoeffValue(elPropCoeffLookup.eProdTime)))
                If lLockedColonists = -1 Then
                    lLockedColonists = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eColonist) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eColonist))))
                End If
                Dim decColonistPayment As Decimal = CLng(Math.Pow(lLockedColonists, GetCoeffValue(elPropCoeffLookup.eColonist)))
                If lLockedEnlisted = -1 Then
                    lLockedEnlisted = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eEnlisted) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eEnlisted))))
                End If
                Dim decEnlistedPayment As Decimal = CLng(Math.Pow(lLockedEnlisted, GetCoeffValue(elPropCoeffLookup.eEnlisted)))
                If lLockedOfficers = -1 Then
                    lLockedOfficers = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eOfficer) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eOfficer))))
                End If
                Dim decOfficersPayment As Decimal = CLng(Math.Pow(lLockedOfficers, GetCoeffValue(elPropCoeffLookup.eOfficer)))
                If bMin1Locked = False Then
                    lLockedMin1 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin1) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin1))))
                End If
                blNewPay = 0
                If lLockedMin1 > 0 Then blNewPay = CLng(Math.Pow(lLockedMin1, GetCoeffValue(elPropCoeffLookup.eMin1)))
                Dim decMineral1Payment As Decimal = blNewPay
                If bMin2Locked = False Then
                    lLockedMin2 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin2) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin2))))
                End If
                blNewPay = 0
                If lLockedMin2 > 0 Then blNewPay = CLng(Math.Pow(lLockedMin2, GetCoeffValue(elPropCoeffLookup.eMin2)))
                Dim decMineral2Payment As Decimal = blNewPay
                If bMin3Locked = False Then
                    lLockedMin3 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin3) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin3))))
                End If
                blNewPay = 0
                If lLockedMin3 > 0 Then blNewPay = CLng(Math.Pow(lLockedMin3, GetCoeffValue(elPropCoeffLookup.eMin3)))
                Dim decMineral3Payment As Decimal = blNewPay
                If bMin4Locked = False Then
                    lLockedMin4 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin4) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin4))))
                End If
                blNewPay = 0
                If lLockedMin4 > 0 Then blNewPay = CLng(Math.Pow(lLockedMin4, GetCoeffValue(elPropCoeffLookup.eMin4)))
                Dim decMineral4Payment As Decimal = blNewPay
                If bMin5Locked = False Then
                    lLockedMin5 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin5) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin5))))
                End If
                blNewPay = 0
                If lLockedMin5 > 0 Then blNewPay = CLng(Math.Pow(lLockedMin5, GetCoeffValue(elPropCoeffLookup.eMin5)))
                Dim decMineral5Payment As Decimal = blNewPay
                If bMin6Locked = False Then
                    lLockedMin6 = CInt(Math.Pow(decDistVals(elPropCoeffLookup.eMin6) * decDiff, (1 / GetCoeffValue(elPropCoeffLookup.eMin6))))
                End If
                blNewPay = 0
                If lLockedMin6 > 0 Then blNewPay = CLng(Math.Pow(lLockedMin6, GetCoeffValue(elPropCoeffLookup.eMin6)))
                Dim decMineral6Payment As Decimal = blNewPay

                If blLockedProdCost < blLockedResCost Then
                    If blLockedProdCost < blLockedResCost * decProdMin Then
                        Return True
                    End If
                Else
                    If blLockedResCost < blLockedProdCost * decProdMin Then
                        Return True
                    End If
                End If

                'prodtime must be 1/1000th of res time
                Dim blCompareProdTime As Int64 = blLockedProdTime \ 528
                Dim blCompareResTime As Int64 = blLockedResTime \ 92
                If blCompareProdTime < blCompareResTime Then
                    If blCompareProdTime < blCompareResTime * 0.001 Then
                        Return True
                    End If
                Else
                    If blCompareResTime < blCompareProdTime * 0.001 Then
                        Return True
                    End If
                End If

                If lLockedMin1 + lLockedMin2 + lLockedMin3 + lLockedMin4 + lLockedMin5 + lLockedMin6 < lLockedHull Then
                    Return True
                End If

                'HullSize PP + Power PP > 10% of total points to be paid
                'Dim decHull As Decimal = CDec(Math.Pow(lLockedHull, GetCoeffValue(elPropCoeffLookup.eHull)))
                'Dim decPower As Decimal = CDec(Math.Pow(lLockedPower, GetCoeffValue(elPropCoeffLookup.ePower)))
                Dim decMin As Decimal = 0.08D   '.095
                'If iTechTypeID = ObjectType.eRadarTech Then decMin = 0.005D
                If decHullPayment + decPowerPayment < decTotalBill * decMin Then
                    Return True
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
                    Return True
                End If

                If decResTimePayment + decProdTimePayment < decTotalBill * 0.08D Then           '.1
                    Return True
                ElseIf decResTimePayment + decProdTimePayment > decTotalBill * 0.9D Then
                    Return True
                End If

                Dim decPCMin As Decimal = 0.03D     '.05
                'If iTechTypeID = ObjectType.eRadarTech Then decPCMin = 0.02D
                If decResCostPayment + decProdCostPayment < decPCMin * decTotalBill Then
                    Return True
                ElseIf decProdCostPayment + decResCostPayment > 0.5D * decTotalBill Then
                    Return True
                End If

                'HullRequired AV > Col AV + Enl AV + Off AV
                If lLockedHull < lLockedColonists + lLockedEnlisted + lLockedOfficers Then
                    Return True
                End If

                'Enl AV + Off AV > Col AV
                If lLockedEnlisted + lLockedOfficers < lLockedColonists Then
                    Return True
                End If

                'Enl AV < Off AV * 5
                If lLockedEnlisted > lLockedOfficers * 5 Then
                    Return True
                End If

                'HullRequired AV > Crew Hull Size AV
                'this rule only applies to hulltypes other than light fighter, medium fighter, heavy fighter, small vehicles and tanks
                If lHullTypeID <> Epica_Tech.eyHullType.Tank AndAlso lHullTypeID <> Epica_Tech.eyHullType.LightFighter AndAlso lHullTypeID <> Epica_Tech.eyHullType.MediumFighter AndAlso _
                   lHullTypeID <> Epica_Tech.eyHullType.HeavyFighter AndAlso lHullTypeID <> Epica_Tech.eyHullType.SmallVehicle Then
                    If lLockedHull < (lLockedOfficers + lLockedColonists + lLockedEnlisted) * Player_Specials.GetPropertyValueByID(oOwner.oSpecials, PlayerSpecialAttributeID.eBetterHullToResidence) Then
                        Return True
                    End If
                End If

                Dim decBill80 As Decimal = decTotalBill * 0.8D
                If decResCostPayment > decBill80 OrElse decResTimePayment > decBill80 OrElse decProdCostPayment > decBill80 OrElse decProdTimePayment > decBill80 OrElse decMineral1Payment > decBill80 OrElse decMineral2Payment > decBill80 Then
                    Return True
                ElseIf decMineral3Payment > decBill80 OrElse decMineral4Payment > decBill80 OrElse decMineral5Payment > decBill80 OrElse decMineral6Payment > decBill80 OrElse decHullPayment > decBill80 OrElse decColonistPayment > decBill80 OrElse decEnlistedPayment > decBill80 OrElse decOfficersPayment > decBill80 Then
                    Return True
                End If

                'Dim l15thMinerals As Int32 = CInt((lLockedMin1 + lLockedMin2 + lLockedMin3 + lLockedMin4 + lLockedMin5 + lLockedMin6) * 0.1)
                'If lMineral1ID > 0 AndAlso lLockedMin1 < l15thMinerals Then
                '    Return True
                'ElseIf lMineral2ID > 0 AndAlso lLockedMin2 < l15thMinerals Then
                '    Return True
                'ElseIf lMineral3ID > 0 AndAlso lLockedMin3 < l15thMinerals Then
                '    Return True
                'ElseIf lMineral4ID > 0 AndAlso lLockedMin4 < l15thMinerals Then
                '    Return True
                'ElseIf lMineral5ID > 0 AndAlso lLockedMin5 < l15thMinerals Then
                '    Return True
                'ElseIf lMineral6ID > 0 AndAlso lLockedMin6 < l15thMinerals Then
                '    Return True
                'End If


                If iTechTypeID <> ObjectType.eEngineTech Then
                    decMin = 0.01D
                    'If iTechTypeID = ObjectType.eRadarTech Then decMin = 0.001D
                    If decPowerPayment < decTotalBill * decMin Then
                        Return True
                    End If
                    If decHullPayment < decTotalBill * decMin Then
                        Return True
                    End If
                End If

                Dim bNoCrewHull As Boolean = (lHullTypeID = Epica_Tech.eyHullType.Escort OrElse lHullTypeID = Epica_Tech.eyHullType.HeavyFighter OrElse lHullTypeID = Epica_Tech.eyHullType.LightFighter OrElse lHullTypeID = Epica_Tech.eyHullType.MediumFighter OrElse lHullTypeID = Epica_Tech.eyHullType.Tank)

                If lLockedHull < 1 Then
                    Return True
                ElseIf lLockedPower < 1 AndAlso iTechTypeID <> ObjectType.eEngineTech Then
                    Return True
                ElseIf lLockedOfficers < 1 AndAlso bOffLockedAtStart = True AndAlso bNoCrewHull = False Then
                    Return True
                ElseIf lLockedEnlisted < 1 AndAlso bEnlLockedAtStart = True AndAlso bNoCrewHull = False Then
                    Return True
                ElseIf lLockedColonists < 1 AndAlso bColLockedAtStart = True AndAlso bNoCrewHull = False Then
                    Return True
                ElseIf blLockedProdCost < 1 Then
                    Return True
                ElseIf blLockedProdTime < 1 Then
                    Return True
                ElseIf blLockedResCost < 1 Then
                    Return True
                ElseIf blLockedResTime < 1 Then
                    Return True
                ElseIf lLockedMin1 < 1 AndAlso lMineral1ID > 0 Then
                    Return True
                ElseIf lLockedMin2 < 1 AndAlso lMineral2ID > 0 Then
                    Return True
                ElseIf lLockedMin3 < 1 AndAlso lMineral3ID > 0 Then
                    Return True
                ElseIf lLockedMin4 < 1 AndAlso lMineral4ID > 0 Then
                    Return True
                ElseIf lLockedMin5 < 1 AndAlso lMineral5ID > 0 Then
                    Return True
                ElseIf lLockedMin6 < 1 AndAlso lMineral6ID > 0 Then
                    Return True
                End If

            Else
                Return False
            End If
        Catch
        End Try
        Return False
    End Function

End Class
