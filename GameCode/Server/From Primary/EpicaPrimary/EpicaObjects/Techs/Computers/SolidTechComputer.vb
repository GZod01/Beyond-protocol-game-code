Option Strict On

Public Class SolidTechComputer
    Inherits TechBuilderComputer

    Public decDPS As Decimal
    Public blMaxRng As Int64
    Public blAccuracy As Int64
    Public blMaxDmg As Int64


    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
    '    '(1+abs(Desired-Actual))*DPS*(MaxRange+TargetAlign+1)/21
    '    Dim decTemp As Decimal = (1D + lMaxDA) * decDPS * (blMaxRng + blAccuracy + 1) / 21D
    '    decTemp /= 110D
    '    decTemp *= lSumDAForMineral
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateThresholdValue() As Decimal
    '    Return decDPS
    'End Function

    'Protected Overrides Sub LoadPointVals()
    '    'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
    '    'If sPath.EndsWith("\") = False Then sPath &= "\"
    '    'sPath &= "TECH.dat"
    '    'Dim oINI As New InitFile(sPath)
    '    'Dim sHeader As String = "SOLIDBEAM"
    '    ml_POINT_PER_HULL = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerHull", "1000")))
    '    ml_POINT_PER_POWER = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerPower", "1000")))
    '    ml_POINT_PER_RES_TIME = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "100")))
    '    ml_POINT_PER_RES_COST = 230 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "230")))
    '    ml_POINT_PER_PROD_TIME = 300 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "300")))
    '    ml_POINT_PER_PROD_COST = 180 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "180")))
    '    ml_POINT_PER_COLONIST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "1")))
    '    ml_POINT_PER_ENLISTED = 10 'CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "10")))
    '    ml_POINT_PER_OFFICER = 40 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "40")))

    '    ml_POINT_PER_MIN1 = 70 'CInt(Val(oINI.GetString(sHeader, "PointPerCoil", "70")))
    '    ml_POINT_PER_MIN2 = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerCoupler", "100")))
    '    ml_POINT_PER_MIN3 = 120 'CInt(Val(oINI.GetString(sHeader, "PointPerCasing", "120")))
    '    ml_POINT_PER_MIN4 = 110 'CInt(Val(oINI.GetString(sHeader, "PointPerFocuser", "110")))
    '    ml_POINT_PER_MIN5 = 90 'CInt(Val(oINI.GetString(sHeader, "PointPerMedium", "90")))
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

    '    'oINI.WriteString(sHeader, "PointPerCoil", ml_POINT_PER_MIN1.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCoupler", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCasing", ml_POINT_PER_MIN3.ToString)
    '    'oINI.WriteString(sHeader, "PointPerFocuser", ml_POINT_PER_MIN4.ToString)
    '    'oINI.WriteString(sHeader, "PointPerMedium", ml_POINT_PER_MIN5.ToString)
    '    'oINI = Nothing
    'End Sub

    'Protected Overrides Function CalculateBaseHull() As Int32
    '    Dim decTemp As Decimal = blMaxDmg * 5D
    '    decTemp += blMaxRng * 0.1D
    '    decTemp += blAccuracy * 0.1D
    '    decTemp += decDPS * 2D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBasePower() As Int32
    '    Dim decTemp As Decimal = blMaxRng * 3D
    '    decTemp += blAccuracy
    '    decTemp += decDPS
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResTime() As Int64
    '    Dim decTemp As Decimal = blMaxRng * 100000D
    '    decTemp += blAccuracy * 100000D
    '    decTemp += decDPS * 1000000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResCost() As Int64
    '    Dim decTemp As Decimal = blMaxRng * 100000D
    '    decTemp += blAccuracy * 100000D
    '    decTemp += decDPS * 984321D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdTime() As Int64
    '    Dim decTemp As Decimal = blMaxRng * 132000D
    '    decTemp += blAccuracy * 187000D
    '    decTemp += decDPS * 837022D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdCost() As Int64
    '    Dim decTemp As Decimal = blMaxRng * 132000D
    '    decTemp += blAccuracy * 187000D
    '    decTemp += decDPS * 23702D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseColonists() As Int32
    '    Dim decTemp As Decimal
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseEnlisted() As Int32
    '    Dim decTemp As Decimal
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseOfficers() As Int32
    '    Dim decTemp As Decimal
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
 

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return

        Dim yTemp(,) As Byte = {{156, 171, 5, 167, 128, 102, 5, 13, 142, 122, 4, 2, 163, 95, 10, 10, 108, 10, 142, 10}, _
            {158, 173, 7, 169, 130, 104, 7, 15, 144, 124, 6, 4, 165, 97, 12, 12, 110, 12, 144, 12}, _
            {160, 175, 9, 171, 132, 106, 9, 17, 146, 126, 8, 6, 167, 99, 14, 14, 112, 14, 146, 14}, _
            {162, 177, 11, 173, 134, 108, 11, 19, 148, 128, 10, 8, 169, 101, 16, 16, 114, 16, 148, 16}, _
            {164, 179, 13, 175, 136, 110, 13, 21, 150, 130, 12, 10, 171, 103, 18, 18, 116, 18, 150, 18}, _
            {166, 181, 15, 177, 138, 112, 15, 23, 152, 132, 14, 12, 173, 105, 20, 20, 118, 20, 152, 20}, _
            {168, 183, 17, 179, 140, 114, 17, 25, 154, 134, 16, 14, 175, 107, 22, 22, 120, 22, 154, 22}, _
            {170, 185, 19, 181, 142, 116, 19, 27, 156, 136, 18, 16, 177, 109, 24, 24, 122, 24, 156, 24}, _
            {172, 187, 21, 183, 144, 118, 21, 29, 158, 138, 20, 18, 179, 111, 26, 26, 124, 26, 158, 26}, _
            {174, 189, 23, 185, 146, 120, 23, 31, 160, 140, 22, 20, 181, 113, 28, 28, 126, 28, 160, 28}, _
            {176, 191, 25, 187, 148, 122, 25, 33, 162, 142, 24, 22, 183, 115, 30, 30, 128, 30, 162, 30}, _
            {178, 193, 27, 189, 150, 124, 27, 35, 164, 144, 26, 24, 185, 117, 32, 32, 130, 32, 164, 32}, _
            {180, 195, 29, 191, 152, 126, 29, 37, 166, 146, 28, 26, 187, 119, 34, 34, 132, 34, 166, 34}, _
            {182, 197, 31, 193, 154, 128, 31, 39, 168, 148, 30, 28, 189, 121, 36, 36, 134, 36, 168, 36}, _
            {184, 199, 33, 195, 156, 130, 33, 41, 170, 150, 32, 30, 191, 123, 38, 38, 136, 38, 170, 38}, _
            {186, 201, 35, 197, 158, 132, 35, 43, 172, 152, 34, 32, 193, 125, 40, 40, 138, 40, 172, 40}, _
            {188, 203, 37, 199, 160, 134, 37, 45, 174, 154, 36, 34, 195, 127, 42, 42, 140, 42, 174, 42}, _
            {190, 205, 39, 201, 162, 136, 39, 47, 176, 156, 38, 36, 197, 129, 44, 44, 142, 44, 176, 44}, _
            {192, 207, 41, 203, 164, 138, 41, 49, 178, 158, 40, 38, 199, 131, 46, 46, 144, 46, 178, 46}, _
            {194, 209, 43, 205, 166, 140, 43, 51, 180, 160, 42, 40, 201, 133, 48, 48, 146, 48, 180, 48}, _
            {196, 211, 45, 207, 168, 142, 45, 53, 182, 162, 44, 42, 203, 135, 50, 50, 148, 50, 182, 50}}

        ReDim muPropReqs(19)
        Dim lPropID() As Int32 = {1, 10, 15, 14, 6, 21, 13, 11, 1, 12, 11, 3, 7, 21, 11, 13, 21, 6, 9, 7}

        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(0, X)
            muPropReqs(X).lMinID = X \ 4
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
        Dim decCoeffList(,) As Decimal = {{5.5D, 3.09D, 0D, 0D, 19D, 0.93D, 0.667D, 1.01D, 1.1D, 11D, 11D, 11D, 11D, 11D, 0D}, _
            {5.65D, 3.83D, 1D, 1D, 19D, 0.93D, 0.667D, 1.2D, 1.1D, 10D, 10D, 10D, 10D, 10D, 0D}, _
            {5.38D, 3.45D, 0D, 0D, 19D, 0.93D, 0.667D, 1.1D, 1.1D, 9D, 9D, 9D, 9D, 9D, 0D}, _
            {5.65D, 3.83D, 1D, 1D, 19D, 0.93D, 0.667D, 1.2D, 1.1D, 10D, 10D, 10D, 10D, 10D, 0D}, _
            {4.7D, 3.58D, 0D, 0D, 19D, 0.834D, 0.655D, 1.15D, 1.2D, 7.2D, 7.2D, 7.2D, 7.2D, 7.2D, 0D}, _
            {3.68D, 2.88D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.1D, 5.2D, 5.2D, 5.2D, 5.2D, 5.2D, 0D}, _
            {2.944D, 3.73D, 1D, 1D, 19D, 0.6D, 0.6D, 1.072D, 1.207D, 3.9D, 3.9D, 3.7D, 3.8D, 3.8D, 0D}, _
            {2.595D, 3.4D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.016D, 1.091D, 3.23D, 3.23D, 3.23D, 3.23D, 3.23D, 0D}, _
            {2.51D, 3.91D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.04D, 1.214D, 3.1D, 3.1D, 3.1D, 3.1D, 3.1D, 0D}, _
            {2.64D, 3.21D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.994D, 1.015D, 3.2D, 3.2D, 3.2D, 3.2D, 3.2D, 0D}, _
            {2.319D, 2.64D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.977D, 0.974D, 2.79D, 2.79D, 2.79D, 2.79D, 2.79D, 0D}, _
            {1.942D, 2.23D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.962D, 0.9427D, 2.281D, 2.281D, 2.281D, 2.281D, 2.281D, 0D}, _
            {1.69D, 2.07D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.9377D, 0.9433D, 1.945D, 1.945D, 1.945D, 1.945D, 1.945D, 0D}, _
            {1.69D, 2.07D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.9377D, 0.9433D, 1.945D, 1.945D, 1.945D, 1.945D, 1.945D, 0D}, _
            {1.69D, 1.5D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.9377D, 0.9433D, 1.945D, 1.945D, 1.945D, 1.945D, 1.945D, 0D}, _
            {1.942D, 2.23D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.962D, 0.9427D, 2.281D, 2.281D, 2.281D, 2.281D, 2.281D, 0D}, _
            {1.942D, 2.23D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.962D, 0.9427D, 2.281D, 2.281D, 2.281D, 2.281D, 2.281D, 0D}, _
            {2.319D, 2.64D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.977D, 0.974D, 2.79D, 2.79D, 2.79D, 2.79D, 2.79D, 0D}, _
            {2.64D, 3.21D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.994D, 1.015D, 3.2D, 3.2D, 3.2D, 3.2D, 3.2D, 0D}, _
            {2.595D, 3.4D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.016D, 1.091D, 3.23D, 3.23D, 3.23D, 3.23D, 3.23D, 0D}, _
            {2.64D, 3.21D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.994D, 1.015D, 3.2D, 3.2D, 3.2D, 3.2D, 3.2D, 0D}, _
            {2.944D, 3.73D, 1D, 1D, 19D, 0.6D, 0.6D, 1.072D, 1.207D, 3.9D, 3.9D, 3.7D, 3.8D, 3.8D, 0D}, _
            {4.7D, 3.58D, 0D, 0D, 19D, 0.834D, 0.655D, 1.15D, 1.2D, 7.2D, 7.2D, 7.2D, 7.2D, 7.2D, 0D}}

        Return decCoeffList(Me.lHullTypeID, lLookup)
    End Function

    Protected Overrides Function GetTheBill() As Decimal

        Dim decNorm As Decimal = 0D
        Dim lGuns As Int32 = 0
        Dim lDPS As Int32 = 0
        Dim lMaxHull As Int32 = 0
        Dim lHullAvail As Int32 = 0
        Dim lPower As Int32 = 0
        TechBuilderComputer.GetTypeValues(Me.lHullTypeID, decNorm, lGuns, lDPS, lMaxHull, lHullAvail, lPower)

        Dim decTemp As Decimal = decDPS / CDec(lDPS)
        decTemp = CDec(Math.Pow(decTemp, 1.01)) * 210000000
        decTemp *= (blMaxRng / 50D)
        'decTemp += (blAccuracy * 1588235)
        Dim decAccuracyMod As Decimal = blAccuracy
        If blAccuracy > 80 Then
            decAccuracyMod = CDec(Math.Pow(blAccuracy - 80, 2) + blAccuracy)
        End If
        decTemp += (decAccuracyMod * 1588235)

        Return decTemp
    End Function
End Class
