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
    '    '(1+abs(Desired-Actual))*(WeaponAccuracy+ScanRes+OptRange+MaxRange+DisruptStrength+JammingStrength)/7
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

    'Protected Overrides Function CalculateBasePower() As Int32
    '    Dim dPow As Double = 1.5
    '    Dim decTemp As Decimal = CDec(Math.Pow(blWepAcc, dPow) + Math.Pow(blScanRes, dPow) + Math.Pow(blOptRng, dPow) + Math.Pow(blMaxRng, dPow) + Math.Pow(blDisRes, dPow) + Math.Pow(blJamStrength, dPow)) ' + Math.Pow(blJamTargets, dPow))
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
 

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return

        Dim yTemp(,) As Byte = {{10, 130, 130, 157, 10, 121, 1, 103, 10, 108, 10, 112, 10, 10}, _
            {12, 132, 132, 159, 12, 123, 3, 105, 12, 110, 12, 114, 12, 12}, _
            {14, 134, 134, 161, 14, 125, 5, 107, 14, 112, 14, 116, 14, 14}, _
            {16, 136, 136, 163, 16, 127, 7, 109, 16, 114, 16, 118, 16, 16}, _
            {18, 138, 138, 165, 18, 129, 9, 111, 18, 116, 18, 120, 18, 18}, _
            {20, 140, 140, 167, 20, 131, 11, 113, 20, 118, 20, 122, 20, 20}, _
            {22, 142, 142, 169, 22, 133, 13, 115, 22, 120, 22, 124, 22, 22}, _
            {24, 144, 144, 171, 24, 135, 15, 117, 24, 122, 24, 126, 24, 24}, _
            {26, 146, 146, 173, 26, 137, 17, 119, 26, 124, 26, 128, 26, 26}, _
            {28, 148, 148, 175, 28, 139, 19, 121, 28, 126, 28, 130, 28, 28}, _
            {30, 150, 150, 177, 30, 141, 21, 123, 30, 128, 30, 132, 30, 30}, _
            {32, 152, 152, 179, 32, 143, 23, 125, 32, 130, 32, 134, 32, 32}, _
            {34, 154, 154, 181, 34, 145, 25, 127, 34, 132, 34, 136, 34, 34}, _
            {36, 156, 156, 183, 36, 147, 27, 129, 36, 134, 36, 138, 36, 36}, _
            {38, 158, 158, 185, 38, 149, 29, 131, 38, 136, 38, 140, 38, 38}, _
            {40, 160, 160, 187, 40, 151, 31, 133, 40, 138, 40, 142, 40, 40}, _
            {42, 162, 162, 189, 42, 153, 33, 135, 42, 140, 42, 144, 42, 42}, _
            {44, 164, 164, 191, 44, 155, 35, 137, 44, 142, 44, 146, 44, 44}, _
            {46, 166, 166, 193, 46, 157, 37, 139, 46, 144, 46, 148, 46, 46}, _
            {48, 168, 168, 195, 48, 159, 39, 141, 48, 146, 48, 150, 48, 48}, _
            {50, 170, 170, 197, 50, 161, 41, 143, 50, 148, 50, 152, 50, 50}}


        ReDim muPropReqs(13)
        '4,15,14, 12,13,1,15,14, 4,3,15, 14,15,10
        Dim lPropID() As Int32 = {4, 15, 14, 14, 15, 10, 4, 3, 15, 12, 13, 1, 15, 14}

        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)
            If X < 9 Then muPropReqs(X).lMinID = X \ 3 Else muPropReqs(X).lMinID = 3
        Next X
    End Sub

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
        'Dim decCoeffList(,) As Decimal = {{6.9D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 16.7D, 16.7D, 16.7D, 16.7D, 0D, 0D}, _
        '    {6.241D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 13.2D, 13.2D, 13.2D, 13.2D, 0D, 0D}, _
        '    {6D, 3.85D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 11.8D, 11.8D, 11.8D, 11.8D, 0D, 0D}, _
        '    {3.32D, 5D, 0D, 0D, 0D, 0.93D, 0.667D, 1.3D, 1.443D, 4.329D, 4.329D, 4.329D, 4.329D, 0D, 0D}, _
        '    {4.87D, 3.74D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 0D, 0D}, _
        '    {4.3D, 3D, 1D, 1D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 6.5D, 6.5D, 6.5D, 6.5D, 0D, 0D}, _
        '    {3.155D, 3.155D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 0D, 0D}, _
        '    {2.3D, 2.92D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 0D, 0D}, _
        '    {2.55D, 2.89D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.03D, 1.214D, 3D, 3D, 3D, 3D, 0D, 0D}, _
        '    {2.186D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.95D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 0D, 0D}, _
        '    {1.99D, 2.36D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.868D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 0D, 0D}, _
        '    {1.84D, 2.2D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.76D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 0D, 0D}, _
        '    {1.64D, 1.9D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.71D, 0.96D, 1.75D, 1.75D, 1.75D, 1.75D, 0D, 0D}, _
        '    {1.46D, 2D, 1.4D, 1.83D, 2.01D, 0.477D, 0.445D, 0.76D, 1.015D, 1.55D, 1.55D, 1.55D, 1.55D, 0D, 0D}, _
        '    {2.3D, 2.3D, 0D, 0D, 0D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 0D, 0D}, _
        '    {1.84D, 2.165D, 2D, 2.56D, 3.04D, 0.48D, 0.45D, 0.76D, 1.01D, 2.03D, 2.03D, 2.03D, 2.03D, 0D, 0D}, _
        '    {1.84D, 2.165D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.76D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 0D, 0D}, _
        '    {2.01D, 2.345D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.868D, 1.128D, 2.25D, 2.25D, 2.25D, 2.25D, 0D, 0D}, _
        '    {2.186D, 2.6D, 2.6D, 3.27D, 4.4D, 0.6D, 0.5D, 0.95D, 1.19D, 2.48D, 2.48D, 2.48D, 2.48D, 0D, 0D}, _
        '    {2.16D, 2.54D, 3D, 3.2D, 4.24D, 0.6D, 0.5D, 1.1D, 1.27D, 2.455D, 2.455D, 2.455D, 2.455D, 0D, 0D}, _
        '    {2.16D, 2.54D, 3D, 3.2D, 4.24D, 0.6D, 0.5D, 1.1D, 1.27D, 2.455D, 2.455D, 2.455D, 2.455D, 0D, 0D}, _
        '    {2.55D, 2.55D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.03D, 1.214D, 3D, 3D, 3D, 3D, 0D, 0D}, _
        '    {3.32D, 4.58D, 0D, 0D, 0D, 0.93D, 0.667D, 1.3D, 1.443D, 4.329D, 4.329D, 4.329D, 4.329D, 0D, 0D}}

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

        Return CDec(Math.Pow(decTemp, 1.1)) * 2700000
    End Function
End Class
