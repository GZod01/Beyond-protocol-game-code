Option Strict On

Public Class ShieldTechComputer
    Inherits TechBuilderComputer

    Public decHPS As Decimal
    Public blProjHull As Int64
    Public blMaxHP As Int64


    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
    '    '(1+abs(Desired-Actual))*(ProjHullSize+MaxHP)/1000+hpsec
    '    Dim decTemp As Decimal = (1D + lMaxDA) * (blProjHull + blMaxHP) / 1000D + decHPS
    '    decTemp *= lSumDAForMineral
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateThresholdValue() As Decimal
    '    Dim decResult As Decimal = 0D
    '    Dim decPHS2 As Decimal = blProjHull * 0.2D
    '    If blMaxHP > decPHS2 Then
    '        decResult = (blMaxHP - decPHS2)
    '        decResult *= decResult
    '        decResult *= 1000000
    '    End If
    '    If decHPS > blMaxHP Then
    '        decResult += (decHPS - blMaxHP) * 1000000
    '    End If
    '    Return decResult
    'End Function

    'Protected Overrides Sub LoadPointVals()
    '    'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
    '    'If sPath.EndsWith("\") = False Then sPath &= "\"
    '    'sPath &= "TECH.dat"
    '    'Dim oINI As New InitFile(sPath)
    '    'Dim sHeader As String = "SHIELD"
    '    ml_POINT_PER_HULL = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerHull", "1000")))
    '    ml_POINT_PER_POWER = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerPower", "1000")))
    '    ml_POINT_PER_RES_TIME = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "100")))
    '    ml_POINT_PER_RES_COST = 230 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "230")))
    '    ml_POINT_PER_PROD_TIME = 300 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "300")))
    '    ml_POINT_PER_PROD_COST = 180 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "180")))
    '    ml_POINT_PER_COLONIST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "1")))
    '    ml_POINT_PER_ENLISTED = 10 'CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "10")))
    '    ml_POINT_PER_OFFICER = 40 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "40")))

    '    ml_POINT_PER_MIN1 = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerCoil", "100")))
    '    ml_POINT_PER_MIN2 = 70 'CInt(Val(oINI.GetString(sHeader, "PointPerAcc", "70")))
    '    ml_POINT_PER_MIN3 = 120 'CInt(Val(oINI.GetString(sHeader, "PointPerCasing", "120")))

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
    '    'oINI.WriteString(sHeader, "PointPerAcc", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCasing", ml_POINT_PER_MIN3.ToString)
    '    'oINI = Nothing
    'End Sub

    'Protected Overrides Function CalculateBaseHull() As Int32
    '    Dim decTemp As Decimal = (blMaxHP * 0.2D) + (decHPS * 0.4D) + (blProjHull * 0.1D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBasePower() As Int32
    '    Dim decTemp As Decimal = (blMaxHP * 0.02D) + (decHPS * 0.1D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResTime() As Int64
    '    Dim decTemp As Decimal = (blMaxHP * 500D) + (decHPS * 870000D) + (blProjHull * 100D)
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResCost() As Int64
    '    Dim decTemp As Decimal = blMaxHP * 10000D
    '    decTemp += decHPS * 40000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdTime() As Int64
    '    Dim decTemp As Decimal = blMaxHP * 430D
    '    decTemp += decHPS * 721000D
    '    decTemp += blProjHull * 100D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdCost() As Int64
    '    Dim decTemp As Decimal = blMaxHP * 760D
    '    decTemp += decHPS * 320000D
    '    decTemp += blProjHull * 4200D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseColonists() As Int32
    '    Dim decTemp As Decimal = decHPS * 2D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseEnlisted() As Int32
    '    Dim decTemp As Decimal = decHPS * 0.2D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseOfficers() As Int32
    '    Dim decTemp As Decimal = decHPS * 0.02D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function


    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return

        Dim yTemp(,) As Byte = {{136, 20, 158, 121, 20, 156, 143, 112, 15}, _
            {139, 21, 161, 123, 20, 159, 146, 114, 16}, _
            {142, 22, 164, 125, 20, 162, 149, 116, 17}, _
            {145, 23, 167, 128, 20, 165, 152, 118, 18}, _
            {148, 24, 170, 131, 21, 168, 155, 120, 19}, _
            {151, 25, 173, 134, 22, 171, 158, 122, 20}, _
            {154, 26, 176, 137, 23, 174, 161, 124, 21}, _
            {157, 27, 180, 140, 24, 177, 164, 126, 22}, _
            {160, 28, 184, 143, 25, 181, 167, 129, 23}, _
            {163, 29, 188, 146, 26, 185, 170, 132, 24}, _
            {166, 30, 192, 149, 27, 189, 173, 135, 25}, _
            {169, 32, 196, 152, 28, 193, 176, 138, 26}, _
            {172, 34, 200, 155, 29, 197, 180, 141, 27}, _
            {175, 36, 204, 158, 30, 201, 184, 144, 28}, _
            {179, 38, 208, 161, 32, 205, 188, 147, 29}, _
            {183, 40, 119, 164, 34, 137, 192, 132, 30}, _
            {187, 42, 121, 167, 36, 140, 196, 135, 32}, _
            {191, 44, 123, 170, 38, 143, 175, 138, 34}, _
            {195, 46, 125, 173, 40, 146, 179, 141, 36}, _
            {199, 48, 128, 176, 42, 149, 183, 144, 38}, _
            {203, 50, 131, 180, 44, 152, 187, 147, 40}}


        ReDim muPropReqs(8)
        Dim lPropID() As Int32 = {1, 10, 15, 21, 10, 14, 1, 12, 11}

        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)
            muPropReqs(X).lMinID = X \ 3
        Next X
    End Sub
    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
        'Dim decCoeffList(,) As Decimal = {{6.9D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 16.7D, 16.7D, 16.7D, 0D, 0D, 0D}, _
        '    {6.241D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 13.2D, 13.2D, 13.2D, 0D, 0D, 0D}, _
        '    {6D, 3.85D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 11.8D, 11.8D, 11.8D, 0D, 0D, 0D}, _
        '    {3.32D, 5D, 0D, 0D, 0D, 0.93D, 0.667D, 1.3D, 1.443D, 4.329D, 4.329D, 4.329D, 0D, 0D, 0D}, _
        '    {4.87D, 3.74D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 0D, 0D, 0D}, _
        '    {4.3D, 3D, 1D, 1D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 6.5D, 6.5D, 6.5D, 0D, 0D, 0D}, _
        '    {3.155D, 3.155D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 0D, 0D, 0D}, _
        '    {2.3D, 2.92D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 0D, 0D, 0D}, _
        '    {2.55D, 2.89D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.03D, 1.214D, 3D, 3D, 3D, 0D, 0D, 0D}, _
        '    {2.186D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.95D, 1.2D, 2.48D, 2.48D, 2.48D, 0D, 0D, 0D}, _
        '    {1.99D, 2.36D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.868D, 1.128D, 2.216D, 2.216D, 2.216D, 0D, 0D, 0D}, _
        '    {1.84D, 2.2D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.76D, 1.015D, 2.01D, 2.01D, 2.01D, 0D, 0D, 0D}, _
        '    {1.64D, 1.9D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.71D, 0.96D, 1.75D, 1.75D, 1.75D, 0D, 0D, 0D}, _
        '    {1.46D, 2D, 1.4D, 1.83D, 2.01D, 0.477D, 0.445D, 0.76D, 1.015D, 1.55D, 1.55D, 1.55D, 0D, 0D, 0D}, _
        '    {2.3D, 2.3D, 0D, 0D, 0D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 0D, 0D, 0D}, _
        '    {1.84D, 2.165D, 2D, 2.56D, 3.04D, 0.48D, 0.45D, 0.76D, 1.01D, 2.03D, 2.03D, 2.03D, 0D, 0D, 0D}, _
        '    {1.84D, 2.165D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.76D, 1.015D, 2.01D, 2.01D, 2.01D, 0D, 0D, 0D}, _
        '    {2.01D, 2.345D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.868D, 1.128D, 2.25D, 2.25D, 2.25D, 0D, 0D, 0D}, _
        '    {2.186D, 2.6D, 2.6D, 3.27D, 4.4D, 0.6D, 0.5D, 0.95D, 1.19D, 2.48D, 2.48D, 2.48D, 0D, 0D, 0D}, _
        '    {2.16D, 2.54D, 3D, 3.2D, 4.24D, 0.6D, 0.5D, 1.1D, 1.27D, 2.455D, 2.455D, 2.455D, 0D, 0D, 0D}, _
        '    {2.16D, 2.54D, 3D, 3.2D, 4.24D, 0.6D, 0.5D, 1.1D, 1.27D, 2.455D, 2.455D, 2.455D, 0D, 0D, 0D}, _
        '    {2.55D, 2.55D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.03D, 1.214D, 3D, 3D, 3D, 0D, 0D, 0D}, _
        '    {3.32D, 4.58D, 0D, 0D, 0D, 0.93D, 0.667D, 1.3D, 1.443D, 4.329D, 4.329D, 4.329D, 0D, 0D, 0D}}

        Dim decCoeffList(,) As Decimal = {{5.9D, 7.5D, 0D, 0D, 19D, 0.99D, 0.6D, 1.15D, 1.2D, 9D, 9D, 8.5D, 0D, 0D, 0D}, _
            {4.8D, 4.7D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.2D, 5.9D, 6.3D, 6.4D, 0D, 0D, 0D}, _
            {5.1D, 5.5D, 0D, 0D, 19D, 0.93D, 0.667D, 1.1D, 1.1D, 7.1D, 7.1D, 7.2D, 0D, 0D, 0D}, _
            {4.8D, 4.7D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.2D, 5.9D, 6.3D, 6.4D, 0D, 0D, 0D}, _
            {4.26D, 4.18D, 0D, 0D, 19D, 0.834D, 0.655D, 1.1D, 1.1D, 5.6D, 5.7D, 5.8D, 0D, 0D, 0D}, _
            {3.86D, 2.75D, 0D, 0D, 19D, 0.82D, 0.6D, 1.1D, 1.1D, 5D, 5D, 5.1D, 0D, 0D, 0D}, _
            {2.745D, 3.65D, 1D, 1D, 19D, 0.6D, 0.6D, 0.987D, 1.14D, 3.2D, 3.4D, 3.32D, 0D, 0D, 0D}, _
            {2.066D, 2.59D, 3D, 3.55D, 5D, 0.6D, 0.5D, 0.935D, 0.993D, 2.4D, 2.4D, 2.4D, 0D, 0D, 0D}, _
            {2.202D, 2.455D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 0.95D, 1D, 2.58D, 2.58D, 2.58D, 0D, 0D, 0D}, _
            {1.985D, 2.26D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.877D, 0.91D, 2.3D, 2.3D, 2.3D, 0D, 0D, 0D}, _
            {1.845D, 2.13D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.877D, 0.897D, 2.12D, 2.12D, 2.12D, 0D, 0D, 0D}, _
            {1.719D, 1.95D, 2D, 2.5D, 3.04D, 0.486D, 0.45D, 0.877D, 0.876D, 1.95D, 1.95D, 1.95D, 0D, 0D, 0D}, _
            {1.537D, 1.606D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.853D, 0.849D, 1.73D, 1.73D, 1.73D, 0D, 0D, 0D}, _
            {1.537D, 1.25D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.853D, 0.849D, 1.73D, 1.73D, 1.73D, 0D, 0D, 0D}, _
            {2.265D, 1.4D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.8D, 0.9D, 2.62D, 2.62D, 2.62D, 0D, 0D, 0D}, _
            {1.719D, 1.95D, 2D, 2.5D, 3.04D, 0.486D, 0.45D, 0.877D, 0.876D, 1.95D, 1.95D, 1.95D, 0D, 0D, 0D}, _
            {1.719D, 1.95D, 2D, 2.5D, 3.04D, 0.486D, 0.45D, 0.877D, 0.876D, 1.95D, 1.95D, 1.95D, 0D, 0D, 0D}, _
            {1.845D, 2.13D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.877D, 0.897D, 2.12D, 2.12D, 2.12D, 0D, 0D, 0D}, _
            {1.985D, 2.26D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.877D, 0.91D, 2.3D, 2.3D, 2.3D, 0D, 0D, 0D}, _
            {2.066D, 2.59D, 3D, 3.55D, 5D, 0.6D, 0.5D, 0.935D, 0.993D, 2.4D, 2.4D, 2.4D, 0D, 0D, 0D}, _
            {1.985D, 2.26D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.877D, 0.91D, 2.3D, 2.3D, 2.3D, 0D, 0D, 0D}, _
            {2.745D, 3.65D, 1D, 1D, 19D, 0.6D, 0.6D, 0.987D, 1.14D, 3.2D, 3.4D, 3.32D, 0D, 0D, 0D}, _
            {4.26D, 4.18D, 0D, 0D, 19D, 0.834D, 0.655D, 1.1D, 1.1D, 5.6D, 5.7D, 5.8D, 0D, 0D, 0D}}


        Return decCoeffList(Me.lHullTypeID, lLookup)
    End Function

    Protected Overrides Function GetTheBill() As Decimal
        '30,000,000 ^ (1 + (MaxHP / ProjSize)) + (100,000,000 ^ (1 + (decHPS / MaxHP))
        Dim decTemp As Decimal = 1 + (decHPS / blMaxHP)
        decTemp = CDec(Math.Pow(100000000, decTemp))
        decTemp += CDec(Math.Pow(30000000, (1 + (blMaxHP / blProjHull))))

        Return decTemp
    End Function
End Class
