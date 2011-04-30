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

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return

        Dim yTemp(,) As Byte = {{5, 0, 6, 5, 0, 6, 6, 4, 0}, _
            {5, 0, 6, 5, 0, 6, 6, 4, 0}, _
            {6, 0, 7, 5, 0, 6, 6, 5, 0}, _
            {6, 1, 7, 5, 0, 7, 6, 5, 0}, _
            {6, 1, 7, 5, 0, 7, 6, 5, 0}, _
            {6, 1, 7, 5, 0, 7, 6, 5, 0}, _
            {6, 1, 7, 5, 1, 7, 6, 5, 0}, _
            {6, 1, 7, 6, 1, 7, 7, 5, 0}, _
            {6, 1, 7, 6, 1, 7, 7, 5, 1}, _
            {6, 1, 8, 6, 1, 7, 7, 5, 1}, _
            {7, 1, 8, 6, 1, 8, 7, 5, 1}, _
            {7, 1, 8, 6, 1, 8, 7, 5, 1}, _
            {7, 1, 8, 6, 1, 8, 7, 6, 1}, _
            {7, 1, 8, 6, 1, 8, 7, 6, 1}, _
            {7, 1, 8, 6, 1, 8, 8, 6, 1}, _
            {7, 1, 5, 7, 1, 5, 8, 5, 1}, _
            {8, 1, 5, 7, 1, 6, 8, 5, 1}, _
            {8, 1, 5, 7, 1, 6, 7, 5, 1}, _
            {8, 2, 5, 7, 1, 6, 7, 6, 1}, _
            {8, 2, 5, 7, 1, 6, 7, 6, 1}, _
            {8, 2, 5, 7, 1, 6, 8, 6, 1}}


        ReDim muPropReqs(8)
        Dim lPropID() As Int32 = {1, 10, 15, 21, 10, 14, 1, 12, 11}

        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)
            muPropReqs(X).lMinID = X \ 3
        Next X
    End Sub

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

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
 
        'Dim decCoeffList(,) As Decimal = {{6.9D, 9D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 12D, 11D, 10D, 0D, 0D, 0D}, _
        ' {7.2D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 11.2D, 11D, 10.9D, 0D, 0D, 0D}, _
        ' {6.2D, 7D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 8.4D, 9.5D, 9.8D, 0D, 0D, 0D}, _
        ' {7.2D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 11.2D, 11D, 10.9D, 0D, 0D, 0D}, _
        ' {4.87D, 5.2D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 5.99D, 6.73D, 6.21D, 0D, 0D, 0D}, _
        ' {4.7D, 3.95D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 5.9D, 6.3D, 6.4D, 0D, 0D, 0D}, _
        ' {3.155D, 4.234D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 3.8D, 3.9D, 3.82D, 0D, 0D, 0D}, _
        ' {2.3D, 2.92D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.03D, 1.27D, 2.66D, 2.66D, 2.66D, 0D, 0D, 0D}, _
        ' {2.55D, 2.89D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.05D, 1.214D, 3D, 3D, 3D, 0D, 0D, 0D}, _
        ' {2.25D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1D, 1.2D, 2.48D, 2.48D, 2.48D, 0D, 0D, 0D}, _
        ' {2.06D, 2.41D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1D, 1.128D, 2.216D, 2.216D, 2.216D, 0D, 0D, 0D}, _
        ' {1.94D, 2.24D, 2D, 2.5D, 3.04D, 0.486D, 0.45D, 1D, 1D, 2.01D, 2.01D, 2.01D, 0D, 0D, 0D}, _
        ' {1.7D, 1.79D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.96D, 1D, 2D, 2D, 2D, 0D, 0D, 0D}, _
        ' {1.7D, 1.79D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.96D, 1D, 2D, 2D, 2D, 0D, 0D, 0D}, _
        ' {1.7D, 1.79D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.96D, 1D, 2D, 2D, 2D, 0D, 0D, 0D}, _
        ' {1.94D, 2.24D, 2D, 2.5D, 3.04D, 0.486D, 0.45D, 1D, 1D, 2.01D, 2.01D, 2.01D, 0D, 0D, 0D}, _
        ' {1.94D, 2.24D, 2D, 2.5D, 3.04D, 0.486D, 0.45D, 1D, 1D, 2.01D, 2.01D, 2.01D, 0D, 0D, 0D}, _
        ' {2.06D, 2.41D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1D, 1.128D, 2.216D, 2.216D, 2.216D, 0D, 0D, 0D}, _
        ' {2.25D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1D, 1.2D, 2.48D, 2.48D, 2.48D, 0D, 0D, 0D}, _
        ' {2.3D, 2.92D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.03D, 1.27D, 2.66D, 2.66D, 2.66D, 0D, 0D, 0D}, _
        ' {2.25D, 2.59D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1D, 1.2D, 2.48D, 2.48D, 2.48D, 0D, 0D, 0D}, _
        ' {3.155D, 4.234D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 3.8D, 3.9D, 3.82D, 0D, 0D, 0D}, _
        ' {4.87D, 5.2D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 5.99D, 6.73D, 6.21D, 0D, 0D, 0D}}

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
        '30,000,000 ^ (1 + (MaxHP / ProjSize)) + (100,000,000 ^ (1 + (decHPS / MaxHP))
        Dim decTemp As Decimal = 1 + (decHPS / blMaxHP)
        decTemp = CDec(Math.Pow(100000000, decTemp))
        decTemp += CDec(Math.Pow(30000000, (1 + (blMaxHP / blProjHull))))

        Return decTemp
    End Function
End Class
