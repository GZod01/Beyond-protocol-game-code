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

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return
        Dim yTemp(,) As Byte = {{6, 7, 0, 7, 5, 4, 0, 0, 6, 5, 0, 0, 6, 4, 0, 0, 4, 0, 6, 0}, _
            {6, 7, 0, 7, 5, 4, 0, 0, 6, 5, 0, 0, 7, 4, 0, 0, 4, 0, 6, 0}, _
            {6, 7, 0, 7, 5, 4, 0, 0, 6, 5, 0, 0, 7, 4, 0, 0, 4, 0, 6, 0}, _
            {6, 7, 0, 7, 5, 4, 0, 0, 6, 5, 0, 0, 7, 4, 0, 0, 4, 0, 6, 0}, _
            {7, 7, 0, 7, 5, 4, 0, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 6, 0}, _
            {7, 7, 0, 7, 5, 4, 0, 1, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 6, 0}, _
            {7, 7, 0, 7, 6, 4, 0, 1, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 6, 0}, _
            {7, 7, 0, 7, 6, 5, 0, 1, 6, 5, 0, 0, 7, 4, 1, 1, 5, 1, 6, 1}, _
            {7, 8, 0, 7, 6, 5, 0, 1, 6, 5, 0, 0, 7, 4, 1, 1, 5, 1, 6, 1}, _
            {7, 8, 1, 7, 6, 5, 1, 1, 6, 6, 0, 0, 7, 4, 1, 1, 5, 1, 6, 1}, _
            {7, 8, 1, 8, 6, 5, 1, 1, 6, 6, 1, 0, 7, 4, 1, 1, 5, 1, 6, 1}, _
            {7, 8, 1, 8, 6, 5, 1, 1, 7, 6, 1, 1, 7, 5, 1, 1, 5, 1, 7, 1}, _
            {7, 8, 1, 8, 6, 5, 1, 1, 7, 6, 1, 1, 8, 5, 1, 1, 5, 1, 7, 1}, _
            {7, 8, 1, 8, 6, 5, 1, 1, 7, 6, 1, 1, 8, 5, 1, 1, 5, 1, 7, 1}, _
            {7, 8, 1, 8, 6, 5, 1, 1, 7, 6, 1, 1, 8, 5, 1, 1, 5, 1, 7, 1}, _
            {7, 8, 1, 8, 6, 5, 1, 1, 7, 6, 1, 1, 8, 5, 1, 1, 5, 1, 7, 1}, _
            {8, 8, 1, 8, 6, 5, 1, 1, 7, 6, 1, 1, 8, 5, 1, 1, 6, 1, 7, 1}, _
            {8, 8, 1, 8, 6, 5, 1, 2, 7, 6, 1, 1, 8, 5, 1, 1, 6, 1, 7, 1}, _
            {8, 8, 1, 8, 7, 5, 1, 2, 7, 6, 1, 1, 8, 5, 2, 2, 6, 2, 7, 2}, _
            {8, 8, 1, 8, 7, 6, 1, 2, 7, 6, 1, 1, 8, 5, 2, 2, 6, 2, 7, 2}, _
            {8, 9, 1, 8, 7, 6, 1, 2, 7, 6, 1, 1, 8, 5, 2, 2, 6, 2, 7, 2}}


        ReDim muPropReqs(19)
        Dim lPropID() As Int32 = {1, 10, 15, 14, 6, 21, 13, 11, 1, 12, 11, 3, 7, 21, 11, 13, 21, 6, 9, 7}

        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(0, X)
            muPropReqs(X).lMinID = X \ 4
        Next X
    End Sub

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

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
 
        'Dim decCoeffList(,) As Decimal = {{6.9D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 12D, 13D, 14D, 16D, 13D, 0D}, _
        ' {8D, 4D, 1D, 1D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 11D, 11.5D, 11.5D, 11D, 11D, 0D}, _
        ' {6D, 3.85D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 12D, 10.2D, 11.8D, 10D, 10.6D, 0D}, _
        ' {8D, 4D, 1D, 1D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 11D, 11.5D, 11.5D, 11D, 11D, 0D}, _
        ' {6.12D, 4.92D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D, 0D}, _
        ' {4.3D, 3.95D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 6.2D, 6.5D, 6.4D, 6.4D, 6.2D, 0D}, _
        ' {3.2D, 4.5D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 4.048D, 0D}, _
        ' {2.95D, 3.9D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.15D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        ' {2.83D, 4.3D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.2D, 1.214D, 3D, 3D, 3D, 3D, 3D, 0D}, _
        ' {2.96D, 3.635D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.115D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {2.59D, 2.98D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.08D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {2.13D, 2.48D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.06D, 1.015D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {1.85D, 2.3D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1D, 1D, 2.1D, 2.1D, 2.1D, 2.1D, 2.1D, 0D}, _
        ' {1.85D, 2.3D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1D, 1D, 2.1D, 2.1D, 2.1D, 2.1D, 2.1D, 0D}, _
        ' {1.85D, 2.3D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1D, 1D, 2.1D, 2.1D, 2.1D, 2.1D, 2.1D, 0D}, _
        ' {2.13D, 2.48D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.06D, 1.015D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {2.13D, 2.48D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.06D, 1.015D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {2.59D, 2.98D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.08D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {2.96D, 3.635D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.115D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {2.95D, 3.9D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.15D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        ' {2.96D, 3.635D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.115D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {3.2D, 4.5D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 4.048D, 0D}, _
        ' {6.12D, 4.92D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D, 0D}}

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
