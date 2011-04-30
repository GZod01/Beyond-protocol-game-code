Option Strict On

Public Class RadarTechComputer
    Inherits TechBuilderComputer

    Public blWepAcc As Int64
    Public blScanRes As Int64
    Public blOptRng As Int64
    Public blMaxRng As Int64
    Public blDisRes As Int64
    Public blJamStrength As Int64

    Public blJamTargets As Int64
    Public lJamType As Int32 = -1

    'Protected Overrides Function CalculateBaseHull() As Integer
    '    Dim lMaxDPS As Int32 = 0
    '    GetTypeValues(Me.lHullTypeID, 0, 0, lMaxDPS, 0, 0, 0)
    '    Return lMaxDPS
    'End Function

    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
    '    '(1+abs(Desired-Actual))*(WeaponAccuracy+ScanRes+OptRange+MaxRange+DisruptStrength+JammingStrength)/600
    '    Dim decTemp As Decimal = (1D + lMaxDA) * (blWepAcc + blScanRes + blOptRng + blMaxRng + blDisRes + blJamStrength) / 600D
    '    decTemp *= lSumDAForMineral
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateThresholdValue() As Decimal
    '    Return 0
    'End Function

    'Protected Overrides Sub LoadPointVals()
    '    'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
    '    'If sPath.EndsWith("\") = False Then sPath &= "\"
    '    'sPath &= "TECH.dat"
    '    'Dim oINI As New InitFile(sPath)
    '    'Dim sHeader As String = "RADAR"
    '    ml_POINT_PER_HULL = 1000 ' CInt(Val(oINI.GetString(sHeader, "PointPerHull", "1000")))
    '    ml_POINT_PER_POWER = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerPower", "1000")))
    '    ml_POINT_PER_RES_TIME = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "100")))
    '    ml_POINT_PER_RES_COST = 230 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "230")))
    '    ml_POINT_PER_PROD_TIME = 300 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "300")))
    '    ml_POINT_PER_PROD_COST = 180 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "180")))
    '    ml_POINT_PER_COLONIST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "1")))
    '    ml_POINT_PER_ENLISTED = 10 'CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "10")))
    '    ml_POINT_PER_OFFICER = 40 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "40")))

    '    ml_POINT_PER_MIN1 = 70 'CInt(Val(oINI.GetString(sHeader, "PointPerEmitter", "70")))
    '    ml_POINT_PER_MIN2 = 110 'CInt(Val(oINI.GetString(sHeader, "PointPerDetection", "110")))
    '    ml_POINT_PER_MIN3 = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerCollection", "100")))
    '    ml_POINT_PER_MIN4 = 120 'CInt(Val(oINI.GetString(sHeader, "PointPerCasing", "120")))

    '    'oINI.WriteString(sHeader, "PointPerHull", ml_POINT_PER_HULL.ToString)
    '    'oINI.WriteString(sHeader, "PointPerPower", ml_POINT_PER_POWER.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResTime", ml_POINT_PER_RES_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResCost", ml_POINT_PER_RES_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdTime", ml_POINT_PER_PROD_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdCost", ml_POINT_PER_PROD_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerColonist", ml_POINT_PER_COLONIST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerEnlisted", ml_POINT_PER_ENLISTED.ToString)
    '    'oINI.WriteString(sHeader, "PointPerOfficer", ml_POINT_PER_OFFICER.ToString)

    '    'oINI.WriteString(sHeader, "PointPerEmitter", ml_POINT_PER_MIN1.ToString)
    '    'oINI.WriteString(sHeader, "PointPerDetection", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCollection", ml_POINT_PER_MIN3.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCasing", ml_POINT_PER_MIN4.ToString)

    '    'oINI = Nothing
    'End Sub

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return

        Dim yTemp(,) As Byte = {{0, 5, 5, 6, 0, 5, 0, 4, 0, 4, 0, 4, 0, 0}, _
            {0, 5, 5, 6, 0, 5, 0, 4, 0, 4, 0, 4, 0, 0}, _
            {0, 5, 5, 6, 0, 5, 0, 4, 0, 4, 0, 5, 0, 0}, _
            {0, 5, 5, 6, 0, 5, 0, 4, 0, 4, 0, 5, 0, 0}, _
            {0, 5, 5, 7, 0, 5, 0, 4, 0, 5, 0, 5, 0, 0}, _
            {0, 6, 6, 7, 0, 5, 0, 4, 0, 5, 0, 5, 0, 0}, _
            {0, 6, 6, 7, 0, 5, 0, 4, 0, 5, 0, 5, 0, 0}, _
            {1, 6, 6, 7, 1, 5, 0, 5, 1, 5, 1, 5, 1, 1}, _
            {1, 6, 6, 7, 1, 5, 0, 5, 1, 5, 1, 5, 1, 1}, _
            {1, 6, 6, 7, 1, 5, 0, 5, 1, 5, 1, 5, 1, 1}, _
            {1, 6, 6, 7, 1, 6, 0, 5, 1, 5, 1, 5, 1, 1}, _
            {1, 6, 6, 7, 1, 6, 1, 5, 1, 5, 1, 5, 1, 1}, _
            {1, 6, 6, 7, 1, 6, 1, 5, 1, 5, 1, 5, 1, 1}, _
            {1, 6, 6, 7, 1, 6, 1, 5, 1, 5, 1, 5, 1, 1}, _
            {1, 6, 6, 7, 1, 6, 1, 5, 1, 5, 1, 6, 1, 1}, _
            {1, 6, 6, 8, 1, 6, 1, 5, 1, 5, 1, 6, 1, 1}, _
            {1, 6, 6, 8, 1, 6, 1, 5, 1, 6, 1, 6, 1, 1}, _
            {1, 7, 7, 8, 1, 6, 1, 5, 1, 6, 1, 6, 1, 1}, _
            {2, 7, 7, 8, 2, 6, 1, 5, 2, 6, 2, 6, 2, 2}, _
            {2, 7, 7, 8, 2, 6, 1, 6, 2, 6, 2, 6, 2, 2}, _
            {2, 7, 7, 8, 2, 6, 1, 6, 2, 6, 2, 6, 2, 2}}


        ReDim muPropReqs(13)
        '4,15,14, 12,13,1,15,14, 4,3,15, 14,15,10
        Dim lPropID() As Int32 = {4, 15, 14, 14, 15, 10, 4, 3, 15, 12, 13, 1, 15, 14}

        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)
            If X < 9 Then muPropReqs(X).lMinID = X \ 3 Else muPropReqs(X).lMinID = 3
        Next X
    End Sub

    'Protected Overrides Function CalculateBasePower() As Int32
    '    Dim dPow As Double = 1.5
    '    Dim decTemp As Decimal = CDec(Math.Pow(blWepAcc, dPow) + Math.Pow(blScanRes, dPow) + Math.Pow(blOptRng, dPow) + Math.Pow(blMaxRng, dPow) + Math.Pow(blDisRes, dPow) + Math.Pow(blJamStrength, dPow))

    '    If blJamTargets = -1 Then
    '        decTemp += CDec(Math.Pow(255, dPow))
    '    Else
    '        decTemp += CDec(Math.Pow(blJamTargets, dPow))
    '    End If

    '    Dim lBaseHull As Int32 = CalculateBaseHull()
    '    If lBaseHull > 60 Then decTemp *= 2

    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResTime() As Int64
    '    Dim decTemp As Decimal = blWepAcc * 1000000D
    '    decTemp += blScanRes
    '    decTemp += blOptRng * 1000000D
    '    decTemp += blMaxRng * 1000000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResCost() As Int64
    '    Dim decTemp As Decimal = blWepAcc * 32000D
    '    decTemp += blScanRes * 10000D
    '    decTemp += blOptRng * 23000D
    '    decTemp += blMaxRng * 21000D
    '    decTemp += blDisRes * 10000D
    '    decTemp += blJamStrength * 21000D

    '    Dim lBaseHull As Int32 = CalculateBaseHull()
    '    If lBaseHull > 60 Then decTemp *= 10

    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdTime() As Int64
    '    Dim decTemp As Decimal = blWepAcc * 45300D
    '    decTemp += blScanRes * 54300D
    '    decTemp += blOptRng * 43000D
    '    decTemp += blMaxRng * 13000D
    '    decTemp += blDisRes * 40000D
    '    decTemp += blJamStrength * 76000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdCost() As Int64
    '    Dim decTemp As Decimal = blWepAcc * 8670D
    '    decTemp += blScanRes * 4530D
    '    decTemp += blOptRng * 18430D
    '    decTemp += blMaxRng * 4300D
    '    decTemp += blDisRes * 8600D
    '    decTemp += blJamStrength * 7120D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseColonists() As Int32
    '    Dim decTemp As Decimal = blDisRes
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseEnlisted() As Int32
    '    Dim decTemp As Decimal = blDisRes * 0.1D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseOfficers() As Int32
    '    Dim decTemp As Decimal = (blOptRng * 0.1D) + (blMaxRng * 0.1D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
        If Me.lHullTypeID = -1 Then Return 0D
 

        'Dim decCoeffList(,) As Decimal = {{6.9D, 4.5D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 14D, 14D, 14D, 14D, 0D, 0D}, _
        ' {7.241D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 11D, 11.3D, 10.9D, 11.1D, 0D, 0D}, _
        ' {6D, 3.85D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 9.2D, 10.8D, 13D, 9.4D, 0D, 0D}, _
        ' {7.241D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 11D, 11.3D, 10.9D, 11.1D, 0D, 0D}, _
        ' {4.87D, 3.74D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.2D, 7.45D, 6.8D, 7.65D, 0D, 0D}, _
        ' {4.7D, 3.95D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 5.9D, 6.3D, 6.4D, 6D, 0D, 0D}, _
        ' {3.155D, 3.155D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 0D, 0D}, _
        ' {2.3D, 2.92D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 0D, 0D}, _
        ' {2.55D, 2.89D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.05D, 1.214D, 3D, 3D, 3D, 3D, 0D, 0D}, _
        ' {2.186D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 0D, 0D}, _
        ' {1.99D, 2.36D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 0D, 0D}, _
        ' {2.3D, 2.8D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.1D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 0D, 0D}, _
        ' {2.5D, 2.5D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.2D, 1.06D, 2.2D, 2.2D, 2.3D, 2.2D, 0D, 0D}, _
        ' {2.5D, 2.5D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.2D, 1.06D, 2.2D, 2.2D, 2.3D, 2.2D, 0D, 0D}, _
        ' {2.5D, 2.5D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.2D, 1.06D, 2.2D, 2.2D, 2.3D, 2.2D, 0D, 0D}, _
        ' {2.3D, 2.8D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.1D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 0D, 0D}, _
        ' {2.3D, 2.8D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.1D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 0D, 0D}, _
        ' {1.99D, 2.36D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 0D, 0D}, _
        ' {2.186D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 0D, 0D}, _
        ' {2.3D, 2.92D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 0D, 0D}, _
        ' {2.186D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 0D, 0D}, _
        ' {3.155D, 3.155D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 0D, 0D}, _
        ' {4.87D, 3.74D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.2D, 7.45D, 6.8D, 7.65D, 0D, 0D}}
        Dim decCoeffList(,) As Decimal = {{5.95D, 3.5D, 0D, 0D, 19D, 0.93D, 0.667D, 1.2D, 1.443D, 14D, 14D, 13D, 13D, 0D, 0D}, _
            {5.32D, 3.4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.15D, 1.35D, 8.5D, 8.5D, 8.5D, 8.5D, 0D, 0D}, _
            {5.3D, 3.27D, 0D, 0D, 19D, 0.93D, 0.667D, 1.15D, 1.2D, 8D, 9D, 9D, 8D, 0D, 0D}, _
            {5.32D, 3.4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.15D, 1.35D, 8.5D, 8.5D, 8.5D, 8.5D, 0D, 0D}, _
            {4.4D, 3.35D, 0D, 0D, 19D, 0.834D, 0.655D, 1.15D, 1.2D, 6.2D, 6.3D, 6.1D, 6.2D, 0D, 0D}, _
            {4.7D, 3.1D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.2D, 5.9D, 6.3D, 6.4D, 6D, 0D, 0D}, _
            {2.877D, 2.85D, 1D, 1D, 19D, 0.6D, 0.6D, 0.96D, 1.1D, 3.7D, 3.8D, 3.8D, 3.5D, 0D, 0D}, _
            {2.234D, 2.836D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.08D, 1.05D, 2.66D, 2.66D, 2.66D, 2.66D, 0D, 0D}, _
            {2.7D, 2.83D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.01D, 1.1D, 3.35D, 3.35D, 3.35D, 3.35D, 0D, 0D}, _
            {2.322D, 2.77D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.065D, 1.045D, 2.73D, 2.73D, 2.73D, 2.73D, 0D, 0D}, _
            {2.117D, 2.54D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.068D, 1.033D, 2.45D, 2.45D, 2.45D, 2.45D, 0D, 0D}, _
            {2.21D, 2.68D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.057D, 1.057D, 2.56D, 2.56D, 2.56D, 2.56D, 0D, 0D}, _
            {2.265D, 2.5D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.088D, 1.056D, 2.62D, 2.62D, 2.62D, 2.62D, 0D, 0D}, _
            {2.265D, 2.5D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.088D, 1.056D, 2.62D, 2.62D, 2.62D, 2.62D, 0D, 0D}, _
            {2.265D, 2.5D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.235D, 1.5D, 2.62D, 2.62D, 2.62D, 2.62D, 0D, 0D}, _
            {2.21D, 2.68D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.057D, 1.057D, 2.56D, 2.56D, 2.56D, 2.56D, 0D, 0D}, _
            {2.21D, 2.68D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.057D, 1.057D, 2.56D, 2.56D, 2.56D, 2.56D, 0D, 0D}, _
            {2.117D, 2.54D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.068D, 1.033D, 2.45D, 2.45D, 2.45D, 2.45D, 0D, 0D}, _
            {2.322D, 2.77D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.065D, 1.045D, 2.73D, 2.73D, 2.73D, 2.73D, 0D, 0D}, _
            {2.234D, 2.836D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.08D, 1.05D, 2.66D, 2.66D, 2.66D, 2.66D, 0D, 0D}, _
            {2.322D, 2.77D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.065D, 1.045D, 2.73D, 2.73D, 2.73D, 2.73D, 0D, 0D}, _
            {2.877D, 2.85D, 1D, 1D, 19D, 0.6D, 0.6D, 0.96D, 1.1D, 3.7D, 3.8D, 3.8D, 3.5D, 0D, 0D}, _
            {4.4D, 3.35D, 0D, 0D, 19D, 0.834D, 0.655D, 1.15D, 1.2D, 6.2D, 6.3D, 6.1D, 6.2D, 0D, 0D}}

        Dim ofrm As frmCoeff = CType(goUILib.GetWindow("frmCoeff"), frmCoeff)
        If ofrm Is Nothing = False Then
            Dim decTemp(decCoeffList.GetUpperBound(1)) As Decimal
            For X As Int32 = 0 To decCoeffList.GetUpperBound(1)
                decTemp(X) = decCoeffList(Me.lHullTypeID, X)
            Next X
            decTemp = ofrm.GetCoeff(decTemp)
            Return decTemp(lLookup)
        End If

        Return decCoeffList(Me.lHullTypeID, lLookup)
    End Function

    Protected Overrides Function GetTheBill() As Decimal
        Dim decTemp As Decimal = 0

        Dim decWepAcc As Decimal = blWepAcc
        If blWepAcc > 80 Then
            decWepAcc = CDec(Math.Pow(blWepAcc - 80, 2) + 80)
        End If
        Dim decScanRes As Decimal = blScanRes
        If blScanRes > 80 Then
            decScanRes = CDec(Math.Pow(blScanRes - 80, 2) + 80)
        End If
        Dim decOptRng As Decimal = blOptRng
        If blOptRng > 80 Then
            decOptRng = CDec(Math.Pow(blOptRng - 80, 2) + 80)
        End If
        Dim decMaxRng As Decimal = blMaxRng
        If blMaxRng > 80 Then
            decMaxRng = CDec(Math.Pow(blMaxRng - 80, 2) + 80)
        End If
        Dim decDisRes As Decimal = blDisRes
        If blDisRes > 80 Then
            decDisRes = CDec(Math.Pow(blDisRes - 80, 2) + 80)
        End If
        Dim decJamStrength As Decimal = blJamStrength
        If blJamStrength > 80 Then
            decJamStrength = CDec(Math.Pow(blJamStrength - 80, 2) + 80)
        End If

        decTemp = decWepAcc + decScanRes + decOptRng + decMaxRng + decDisRes + decJamStrength
        If lJamType > 0 AndAlso blJamStrength > 0 Then
            decTemp += CDec(Math.Pow(blJamTargets, 5))
        End If

        Return CDec(Math.Pow(decTemp, 1.1)) * 2700000 'Return CDec(Math.Pow(decTemp, 1.2)) * 2700000
    End Function
End Class
