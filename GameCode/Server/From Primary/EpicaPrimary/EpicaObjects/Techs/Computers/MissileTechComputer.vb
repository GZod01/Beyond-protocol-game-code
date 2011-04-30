Option Strict On

Public Class MissileTechComputer
    Inherits TechBuilderComputer

    Public blMaxDamage As Int64
    Public blHullSize As Int64
    Public blMaxSpeed As Int64
    Public blManeuver As Int64
    Public blROF As Int64
    Public blRange As Int64
    Public blHomingAccuracy As Int64
    Public blExplosionRadius As Int64
    Public blStructHP As Int64

    Public decDPS As Decimal

    Public yPayloadType As Byte

    'Protected Overrides Function CalculateBaseColonists() As Integer
    '    Dim decTemp As Decimal = blHomingAccuracy + blExplosionRadius
    '    If decTemp < 0 Then Return 0
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseEnlisted() As Integer
    '    Dim decTemp As Decimal = (CDec(blManeuver) * 0.06D) + (CDec(blHomingAccuracy) * 0.3D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseHull() As Integer
    '    Dim decTemp As Decimal = blHullSize
    '    Dim decROF As Decimal = blROF / 30D
    '    decTemp += ((100D / decROF) * (blHullSize + blStructHP))
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
    '    Dim decTemp As Decimal = (1D + CDec(lMaxDA)) * (500D / CDec(blHullSize) + decDPS + CDec(blMaxSpeed) + CDec(blManeuver) + CDec(blHomingAccuracy) + CDec(blExplosionRadius) + CDec(blStructHP)) / 30D
    '    decTemp *= lSumDAForMineral
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseOfficers() As Integer
    '    Dim decTemp As Decimal = CDec(blExplosionRadius) * 0.2D
    '    If decTemp < 0 Then Return 0
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateBasePower() As Integer
    '    Dim decTemp As Decimal = blMaxDamage * 10D
    '    Dim decROF As Decimal = blROF / 30D

    '    decTemp += CDec(blMaxSpeed) * 4D
    '    decTemp += CDec(blManeuver) * 20D
    '    decTemp += (3D / decROF) * 50D
    '    decTemp += blRange
    '    decTemp += blHomingAccuracy * 10D
    '    decTemp += blExplosionRadius * 8D
    '    decTemp += blStructHP
    '    decTemp /= 10D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseProdCost() As Long
    '    Dim decTemp As Decimal = CDec(blMaxDamage) * 10000D
    '    decTemp += (1D / blHullSize) * 10000000D
    '    decTemp += CDec(blMaxDamage) * 10000D
    '    decTemp += CDec(blManeuver) * 100000D
    '    decTemp += CDec(blRange) * 100000D
    '    decTemp += CDec(blExplosionRadius) * 100000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseProdTime() As Long
    '    Dim decROF As Decimal = blROF / 30D
    '    Dim decTemp As Decimal = (1D / blROF) * 10000000D
    '    decTemp += CDec(blManeuver) * 100000D
    '    decTemp += CDec(blRange) * 100000D
    '    decTemp += CDec(blStructHP) * 100000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseResCost() As Long
    '    Dim decROF As Decimal = blROF / 30D
    '    Dim decTemp As Decimal = (1D / blROF) * 10000000D
    '    decTemp += CDec(blManeuver) * 1000000D
    '    decTemp += CDec(blMaxSpeed) * 10000D
    '    decTemp += CDec(blMaxDamage) * 10000D
    '    decTemp += CDec(blRange) * 12000D
    '    decTemp += CDec(blHomingAccuracy) * 300000D
    '    decTemp += CDec(blExplosionRadius) * 400000D
    '    decTemp += CDec(blStructHP)
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function

    'Protected Overrides Function CalculateBaseResTime() As Long
    '    Dim decROF As Decimal = blROF / 30D
    '    Dim decTemp As Decimal = (1D / blROF) * 10000000D
    '    decTemp += CDec(blManeuver) * 100000D
    '    decTemp += CDec(blMaxSpeed) * 100000D
    '    decTemp += CDec(blMaxDamage) * 10000D
    '    decTemp += CDec(blRange) * 100000D
    '    decTemp += CDec(blHomingAccuracy) * 1000000D
    '    decTemp += CDec(blExplosionRadius) * 1000000D
    '    decTemp += CDec(blStructHP)
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function

    'Protected Overrides Function CalculateThresholdValue() As Decimal
    '    Return decDPS
    'End Function

    'Protected Overrides Sub LoadPointVals()
    '    'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
    '    'If sPath.EndsWith("\") = False Then sPath &= "\"
    '    'sPath &= "TECH.dat"
    '    'Dim oINI As New InitFile(sPath)
    '    'Dim sHeader As String = "MISSILES"
    '    ml_POINT_PER_HULL = 600000 'CInt(Val(oINI.GetString(sHeader, "PointPerHull", "600000")))
    '    ml_POINT_PER_POWER = 600000 'CInt(Val(oINI.GetString(sHeader, "PointPerPower", "600000")))
    '    ml_POINT_PER_RES_TIME = 50 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "5")))
    '    ml_POINT_PER_RES_COST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "1")))
    '    ml_POINT_PER_PROD_TIME = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "1")))
    '    ml_POINT_PER_PROD_COST = 72 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "72")))
    '    ml_POINT_PER_COLONIST = 6 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "6")))
    '    ml_POINT_PER_ENLISTED = 60 ' CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "60")))
    '    ml_POINT_PER_OFFICER = 240 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "240")))

    '    ml_POINT_PER_MIN1 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerCoil", "600")))
    '    ml_POINT_PER_MIN2 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerAcc", "600")))
    '    ml_POINT_PER_MIN3 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerCasing", "600")))
    '    ml_POINT_PER_MIN4 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerFocuser", "600")))
    '    ml_POINT_PER_MIN5 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerCompChmbr", "600")))
    '    ml_POINT_PER_MIN6 = 0

    '    'oINI.WriteString(sHeader, "PointPerHull", ml_POINT_PER_HULL.ToString)
    '    'oINI.WriteString(sHeader, "PointPerPower", ml_POINT_PER_POWER.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResTime", ml_POINT_PER_RES_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResCost", ml_POINT_PER_RES_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdTime", ml_POINT_PER_PROD_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdCost", ml_POINT_PER_PROD_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerColonist", ml_POINT_PER_COLONIST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerEnlisted", ml_POINT_PER_ENLISTED.ToString)
    '    'oINI.WriteString(sHeader, "PointPerOfficer", ml_POINT_PER_OFFICER.ToString)

    '    'oINI.WriteString(sHeader, "PointPerBody", ml_POINT_PER_MIN1.ToString)
    '    'oINI.WriteString(sHeader, "PointPerTip", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerFlaps", ml_POINT_PER_MIN3.ToString)
    '    'oINI.WriteString(sHeader, "PointPerFuel", ml_POINT_PER_MIN4.ToString)
    '    'oINI.WriteString(sHeader, "PointPerPayload", ml_POINT_PER_MIN5.ToString)
    '    'oINI = Nothing
    'End Sub

    Protected Overrides Sub SetMinReqProps() 
        Dim yTemp(,) As Byte = {{123, 43, 74, 167, 1, 64, 85, 134, 124, 132, 3, 13}, _
            {124, 44, 75, 169, 2, 65, 86, 136, 125, 133, 4, 14}, _
            {125, 45, 76, 171, 3, 66, 87, 138, 126, 134, 5, 15}, _
            {126, 46, 77, 173, 4, 67, 88, 140, 127, 135, 6, 16}, _
            {127, 47, 78, 175, 5, 68, 89, 142, 128, 136, 7, 17}, _
            {128, 48, 79, 177, 6, 69, 90, 144, 129, 137, 8, 18}, _
            {129, 49, 80, 179, 7, 70, 91, 146, 130, 138, 9, 19}, _
            {130, 50, 81, 181, 8, 71, 92, 148, 131, 139, 10, 20}, _
            {131, 51, 82, 183, 9, 72, 93, 150, 132, 140, 11, 21}, _
            {132, 52, 83, 185, 10, 73, 94, 152, 133, 141, 12, 22}, _
            {133, 53, 84, 187, 11, 74, 95, 154, 134, 142, 13, 23}, _
            {134, 54, 85, 189, 12, 75, 96, 156, 135, 143, 14, 24}, _
            {135, 55, 86, 191, 13, 76, 97, 158, 136, 144, 15, 25}, _
            {136, 56, 87, 193, 14, 77, 98, 160, 137, 145, 16, 26}, _
            {137, 57, 88, 195, 15, 78, 99, 162, 138, 146, 17, 27}, _
            {138, 58, 89, 197, 16, 79, 100, 164, 139, 147, 18, 28}, _
            {139, 59, 90, 199, 17, 80, 101, 166, 140, 148, 19, 29}, _
            {140, 60, 91, 201, 18, 81, 102, 168, 141, 149, 20, 30}, _
            {141, 61, 92, 203, 19, 82, 103, 170, 142, 150, 21, 31}, _
            {142, 62, 93, 205, 20, 83, 104, 172, 143, 151, 22, 32}, _
            {143, 63, 94, 207, 21, 84, 105, 174, 144, 152, 23, 33}}


        ReDim muPropReqs(11)
        Dim lPayloadPropID As Int32 = 0

        If yPayloadType = 0 Then
            lPayloadPropID = 19
        Else
            lPayloadPropID = 18
        End If

        Dim lPropID() As Int32 = {2, 1, 3, 2, 5, 1, 13, 11, 18, 9, 1, lPayloadPropID}
        Dim lMinID() As Int32 = {0, 0, 0, 1, 1, 2, 2, 2, 3, 3, 3, 4}

        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)
            muPropReqs(X).lMinID = lMinID(X)
        Next X


    End Sub
    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
        'Dim decCoeffList(,) As Decimal = {{6.9D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 16.7D, 16.7D, 16.7D, 16.7D, 16.7D, 0D}, _
        '    {6.241D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 13.2D, 13.2D, 13.2D, 13.2D, 13.2D, 0D}, _
        '    {6D, 3.85D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 11.8D, 11.8D, 11.8D, 11.8D, 11.8D, 0D}, _
        '    {3.32D, 5D, 0D, 0D, 0D, 0.93D, 0.667D, 1.3D, 1.443D, 4.329D, 4.329D, 4.329D, 4.329D, 4.329D, 0D}, _
        '    {4.87D, 3.74D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D, 0D}, _
        '    {4.3D, 3D, 1D, 1D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 6.5D, 6.5D, 6.5D, 6.5D, 6.5D, 0D}, _
        '    {3.155D, 3.155D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 4.048D, 0D}, _
        '    {2.3D, 2.92D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        '    {2.55D, 2.89D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.03D, 1.214D, 3D, 3D, 3D, 3D, 3D, 0D}, _
        '    {2.186D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.95D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        '    {1.99D, 2.36D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.868D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        '    {1.84D, 2.2D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.76D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D, 0D}, _
        '    {1.64D, 1.9D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.71D, 0.96D, 1.75D, 1.75D, 1.75D, 1.75D, 1.75D, 0D}, _
        '    {1.46D, 2D, 1.4D, 1.83D, 2.01D, 0.477D, 0.445D, 0.76D, 1.015D, 1.55D, 1.55D, 1.55D, 1.55D, 1.55D, 0D}, _
        '    {2.3D, 1.7D, 0D, 0D, 0D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        '    {1.84D, 2.165D, 2D, 2.56D, 3.04D, 0.48D, 0.45D, 0.76D, 1.01D, 2.03D, 2.03D, 2.03D, 2.03D, 2.03D, 0D}, _
        '    {1.84D, 2.165D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.76D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D, 0D}, _
        '    {2.01D, 2.345D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.868D, 1.128D, 2.25D, 2.25D, 2.25D, 2.25D, 2.25D, 0D}, _
        '    {2.186D, 2.6D, 2.6D, 3.27D, 4.4D, 0.6D, 0.5D, 0.95D, 1.19D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        '    {2.16D, 2.54D, 3D, 3.2D, 4.24D, 0.6D, 0.5D, 1.1D, 1.27D, 2.455D, 2.455D, 2.455D, 2.455D, 2.455D, 0D}, _
        '    {2.16D, 2.54D, 3D, 3.2D, 4.24D, 0.6D, 0.5D, 1.1D, 1.27D, 2.455D, 2.455D, 2.455D, 2.455D, 2.455D, 0D}, _
        '    {2.55D, 2.55D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.03D, 1.214D, 3D, 3D, 3D, 3D, 3D, 0D}, _
        '    {3.32D, 4.58D, 0D, 0D, 0D, 0.93D, 0.667D, 1.3D, 1.443D, 4.329D, 4.329D, 4.329D, 4.329D, 4.329D, 0D}}
        Dim decCoeffList(,) As Decimal = {{6.5D, 3.73D, 0D, 0D, 19D, 0.93D, 0.667D, 1.2D, 1.3D, 13D, 13D, 13D, 13D, 13D, 0D}, _
            {5.9D, 4D, 1D, 1D, 19D, 0.82D, 0.6D, 1.25D, 1.36D, 9D, 9D, 9D, 9D, 9D, 0D}, _
            {5.6D, 3.6D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 9.5D, 9.2D, 9.1D, 10D, 9D, 0D}, _
            {5.9D, 4D, 1D, 1D, 19D, 0.82D, 0.6D, 1.25D, 1.36D, 9D, 9D, 9D, 9D, 9D, 0D}, _
            {4.9D, 3.7D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.2D, 7.8D, 7.8D, 7.8D, 7.8D, 7.8D, 0D}, _
            {3.9D, 3.08D, 0D, 0D, 19D, 0.82D, 0.6D, 1.1D, 1.2D, 5.5D, 5.6D, 5.6D, 5.6D, 5.6D, 0D}, _
            {2.99D, 3.8D, 1D, 1D, 19D, 0.6D, 0.6D, 1.09D, 1.23D, 3.8D, 4D, 3.8D, 4D, 4.048D, 0D}, _
            {2.569D, 3.36D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.006D, 1.075D, 3.2D, 3.2D, 3.2D, 3.2D, 3.2D, 0D}, _
            {2.47D, 3.85D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.04D, 1.1D, 3D, 3.1D, 3.1D, 3.1D, 3.1D, 0D}, _
            {2.58D, 3.13D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.987D, 1D, 3.14D, 3.14D, 3.14D, 3.14D, 3.14D, 0D}, _
            {2.32D, 2.65D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.966D, 0.975D, 2.8D, 2.8D, 2.8D, 2.8D, 2.8D, 0D}, _
            {1.948D, 2.24D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.965D, 0.972D, 2.285D, 2.285D, 2.285D, 2.285D, 2.285D, 0D}, _
            {1.705D, 2.09D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.947D, 0.952D, 1.96D, 1.96D, 1.96D, 1.96D, 1.96D, 0D}, _
            {1.705D, 2.09D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.947D, 0.952D, 1.96D, 1.96D, 1.96D, 1.96D, 1.96D, 0D}, _
            {1.9D, 1.7D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.13D, 1.1D, 2.3D, 2.2D, 2.2D, 2.3D, 2.2D, 0D}, _
            {1.948D, 2.24D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.965D, 0.972D, 2.285D, 2.285D, 2.285D, 2.285D, 2.285D, 0D}, _
            {1.948D, 2.24D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.965D, 0.972D, 2.285D, 2.285D, 2.285D, 2.285D, 2.285D, 0D}, _
            {2.32D, 2.65D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.966D, 0.975D, 2.8D, 2.8D, 2.8D, 2.8D, 2.8D, 0D}, _
            {2.58D, 3.13D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.987D, 1D, 3.14D, 3.14D, 3.14D, 3.14D, 3.14D, 0D}, _
            {2.569D, 3.36D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.006D, 1.075D, 3.2D, 3.2D, 3.2D, 3.2D, 3.2D, 0D}, _
            {2.58D, 3.13D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.987D, 1D, 3.14D, 3.14D, 3.14D, 3.14D, 3.14D, 0D}, _
            {2.99D, 3.8D, 1D, 1D, 19D, 0.6D, 0.6D, 1.09D, 1.23D, 3.8D, 4D, 3.8D, 4D, 4.048D, 0D}, _
            {4.9D, 3.7D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.2D, 7.8D, 7.8D, 7.8D, 7.8D, 7.8D, 0D}}

        'Old Facility; {1.705D, 1.7D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.947D, 0.952D, 1.96D, 1.96D, 1.96D, 1.96D, 1.96D, 0D}, _
        Return decCoeffList(Me.lHullTypeID, lLookup)
    End Function
    Protected Overrides Function GetTheBill() As Decimal
        Dim decNorm As Decimal = 0D
        Dim lGuns As Int32 = 0
        Dim lDPS As Int32 = 0
        Dim lMaxHull As Int32 = 0
        Dim lHullAvail As Int32 = 0
        Dim lPower As Int32 = 0

        GetTypeValues(Me.lHullTypeID, decNorm, lGuns, lDPS, lMaxHull, lHullAvail, lPower)

        Dim decTemp As Decimal
        Dim decCalcDPS As Decimal
        If decDPS > lDPS * 100 Then
            decCalcDPS = CDec(Math.Pow(decDPS - (lDPS * 100), 2)) + decDPS
        Else
            decCalcDPS = decDPS
        End If
        decTemp = CDec(Math.Pow((((decCalcDPS / 55D) / CDec(lDPS))), 1.2) * 540000000)

        Dim decRangeMod As Decimal = blRange
        Dim decSpeedMod As Decimal = blMaxSpeed
        Dim decStructMod As Decimal = blStructHP
        Dim decHomingMod As Decimal = blHomingAccuracy
        Dim decHullSizeMod As Decimal = Math.Max(10, blHullSize)

        If blRange > 150 Then
            decRangeMod = CDec(Math.Pow(blRange - 150, 2)) + blRange
        End If
        decTemp += CDec(decRangeMod * 529411)

        If blMaxSpeed > 50 Then
            decSpeedMod = CDec(Math.Pow(blMaxSpeed - 50, 2)) + blMaxSpeed
        End If
        decTemp += (decSpeedMod * 529411)

        decTemp += CDec(Math.Pow((blExplosionRadius * decCalcDPS), 2) * 1500)

        If blStructHP > 5 Then
            decStructMod = CDec(Math.Pow(blStructHP - 5, 2)) + blStructHP
        End If
        decTemp += (decStructMod * 90000000)

        If blHomingAccuracy > 40 Then
            decHomingMod = CDec(Math.Pow(blHomingAccuracy - 40, 2)) + blHomingAccuracy
        End If
        decTemp += (decHomingMod * 5200000)

        decTemp += (11 - decHullSizeMod) * 10000000000

        Return decTemp
    End Function

End Class
