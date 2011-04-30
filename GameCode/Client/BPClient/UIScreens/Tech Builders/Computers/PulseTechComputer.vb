Option Strict On

Public Class PulseTechComputer
    Inherits TechBuilderComputer

    Public blCompress As Int64
    Public blInputEnergy As Int64
    Public decDPS As Decimal
    Public blMaxRange As Int64
    Public blScatterRadius As Int64
    Public blROF As Int64

    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
    '    '(1+abs(Desired-Actual))*DPS*(Range+Scatter+1)/19
    '    Dim decTemp As Decimal = (1D + lMaxDA) * decDPS * (blMaxRange + blScatterRadius + 1D) / 19D
    '    decTemp /= 290D
    '    decTemp *= lSumDAForMineral
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateThresholdValue() As Decimal
    '    Return decDPS
    'End Function

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return
        Dim yTemp(,) As Byte = {{5, 6, 0, 0, 4, 6, 6, 0, 5, 4, 0, 0, 6, 4, 0, 0, 4, 0, 0, 5}, _
            {5, 7, 0, 0, 4, 6, 6, 0, 5, 4, 0, 0, 6, 4, 0, 0, 4, 0, 0, 5}, _
            {6, 7, 0, 0, 4, 6, 6, 0, 5, 5, 0, 0, 6, 4, 0, 0, 4, 0, 0, 5}, _
            {6, 7, 0, 0, 4, 6, 6, 0, 6, 5, 0, 0, 6, 4, 0, 0, 5, 0, 0, 5}, _
            {6, 7, 0, 0, 5, 7, 6, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 0, 5}, _
            {6, 7, 0, 0, 5, 7, 6, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 0, 5}, _
            {6, 7, 0, 0, 5, 7, 6, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 0, 5}, _
            {6, 7, 0, 0, 5, 7, 7, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 0, 6}, _
            {6, 8, 0, 0, 5, 7, 7, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 0, 6}, _
            {6, 8, 0, 0, 5, 7, 7, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 0, 6}, _
            {7, 8, 0, 0, 5, 8, 7, 0, 6, 5, 0, 0, 7, 4, 0, 0, 5, 0, 0, 6}, _
            {7, 8, 0, 0, 5, 8, 7, 0, 7, 5, 0, 0, 8, 5, 0, 0, 5, 0, 0, 6}, _
            {7, 8, 0, 0, 5, 8, 7, 0, 7, 6, 0, 0, 8, 5, 0, 0, 5, 0, 0, 6}, _
            {7, 8, 0, 0, 6, 8, 7, 0, 7, 6, 0, 0, 8, 5, 0, 0, 6, 0, 0, 6}, _
            {7, 9, 0, 0, 6, 8, 8, 0, 7, 6, 0, 0, 8, 5, 0, 1, 6, 0, 0, 6}, _
            {5, 6, 0, 0, 4, 6, 6, 0, 5, 4, 0, 0, 6, 4, 0, 1, 4, 0, 0, 5}, _
            {5, 7, 0, 0, 4, 6, 6, 1, 5, 4, 0, 0, 6, 4, 0, 1, 4, 0, 0, 5}, _
            {6, 7, 0, 0, 4, 6, 6, 1, 5, 5, 1, 0, 6, 4, 0, 1, 4, 0, 0, 5}, _
            {6, 7, 1, 0, 4, 6, 6, 1, 6, 5, 1, 0, 6, 4, 1, 1, 5, 1, 0, 5}, _
            {6, 7, 1, 0, 5, 7, 6, 1, 6, 5, 1, 0, 7, 4, 1, 1, 5, 1, 0, 5}, _
            {6, 7, 1, 0, 5, 7, 6, 1, 6, 5, 1, 1, 7, 4, 1, 1, 5, 1, 0, 5}}

        ReDim muPropReqs(19)
        Dim lPropID() As Int32 = {1, 10, 15, 14, 21, 10, 14, 15, 1, 12, 11, 3, 7, 21, 11, 13, 1, 5, 3, 4}
        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)
            muPropReqs(X).lMinID = X \ 4
        Next X
    End Sub

    'Protected Overrides Function CalculateBaseHull() As Int32
    '    Dim decTemp As Decimal = blCompress + (decDPS * 4D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBasePower() As Int32
    '    Dim decTemp As Decimal = ((blInputEnergy * 10) + (blScatterRadius * 20)) * decDPS * 0.05D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResTime() As Int64
    '    Dim decTemp As Decimal = blInputEnergy * 2000000D
    '    decTemp += blCompress * 1000000D
    '    Const fMinROF As Single = 1 / 30.0F
    '    Dim fROF As Single = Math.Max(fMinROF, blROF / 30.0F)
    '    decTemp += CDec(1.0F / fROF) * 10000000D
    '    decTemp += blMaxRange * 700000D
    '    decTemp += blScatterRadius * 10000000
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResCost() As Int64
    '    Dim decTemp As Decimal = decDPS * 540000D
    '    decTemp += blMaxRange * 31000D
    '    decTemp += blScatterRadius * 720000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdTime() As Int64
    '    Dim decTemp As Decimal = decDPS * 1400000D
    '    decTemp += blScatterRadius * 500000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdCost() As Int64
    '    Dim decTemp As Decimal = decDPS * 300000D
    '    decTemp += blMaxRange * 21000D
    '    decTemp += blScatterRadius * 121000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseColonists() As Int32
    '    Dim decTemp As Decimal = blMaxRange / 2D
    '    decTemp += blScatterRadius
    '    If decTemp < 0 Then Return 0
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseEnlisted() As Int32
    '    Dim decTemp As Decimal = (blInputEnergy / 1000D) + 1D
    '    Const fMinROF As Single = 1 / 30.0F
    '    Dim fROF As Single = Math.Max(fMinROF, blROF / 30.0F)
    '    decTemp += CDec(10D / fROF)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseOfficers() As Int32
    '    Dim decTemp As Decimal = blCompress / 10D
    '    decTemp *= 1.5D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Sub LoadPointVals()
    '    'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
    '    'If sPath.EndsWith("\") = False Then sPath &= "\"
    '    'sPath &= "TECH.dat"
    '    'Dim oINI As New InitFile(sPath)
    '    'Dim sHeader As String = "PULSE"
    '    ml_POINT_PER_HULL = 600000 'CInt(Val(oINI.GetString(sHeader, "PointPerHull", "600000")))
    '    ml_POINT_PER_POWER = 600000 'CInt(Val(oINI.GetString(sHeader, "PointPerPower", "600000")))
    '    ml_POINT_PER_RES_TIME = 1 '5 '0 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "5")))
    '    ml_POINT_PER_RES_COST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "1")))
    '    ml_POINT_PER_PROD_TIME = 90 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "1")))
    '    ml_POINT_PER_PROD_COST = 72 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "72")))
    '    ml_POINT_PER_COLONIST = 6 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "6")))
    '    ml_POINT_PER_ENLISTED = 60 'CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "60")))
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

    '    'oINI.WriteString(sHeader, "PointPerCoil", ml_POINT_PER_MIN1.ToString)
    '    'oINI.WriteString(sHeader, "PointPerAcc", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCasing", ml_POINT_PER_MIN3.ToString)
    '    'oINI.WriteString(sHeader, "PointPerFocuser", ml_POINT_PER_MIN4.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCompChmbr", ml_POINT_PER_MIN5.ToString)
    '    'oINI = Nothing
    'End Sub

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
 
        'Dim decCoeffList(,) As Decimal = {{6.9D, 4.5D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 14D, 14D, 14D, 14D, 14D, 0D}, _
        ' {8.83D, 5.8D, 1D, 1D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 13.2D, 13.2D, 13.2D, 13.2D, 13.2D, 0D}, _
        ' {6.3D, 4.2D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 12D, 10.2D, 11.8D, 10D, 10.6D, 0D}, _
        ' {8.83D, 5.8D, 1D, 1D, 19D, 0.93D, 0.667D, 1.5D, 1.443D, 13.2D, 13.2D, 13.2D, 13.2D, 13.2D, 0D}, _
        ' {6.12D, 4.76D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D, 0D}, _
        ' {4.3D, 3.95D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 6.2D, 6.5D, 6.4D, 6.4D, 6.2D, 0D}, _
        ' {3.8D, 5.25D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 4.048D, 0D}, _
        ' {3.26D, 4.33D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.3D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        ' {3.19D, 4.9D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.28D, 1.214D, 3D, 3D, 3D, 3D, 3D, 0D}, _
        ' {3.32D, 4.12D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.255D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {2.88D, 3.35D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.2D, 1.128D, 2.3D, 2.3D, 2.3D, 2.3D, 2.3D, 0D}, _
        ' {2.4D, 2.8D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.195D, 1.07D, 2.3D, 2.3D, 2.3D, 2.3D, 2.3D, 0D}, _
        ' {2D, 2.53D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.1D, 1.02D, 2.2D, 2.2D, 2.2D, 2.2D, 2.2D, 0D}, _
        ' {2D, 2.53D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.1D, 1.02D, 2.2D, 2.2D, 2.2D, 2.2D, 2.2D, 0D}, _
        ' {2D, 2.53D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1.1D, 1.02D, 2.2D, 2.2D, 2.2D, 2.2D, 2.2D, 0D}, _
        ' {2.4D, 2.8D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.195D, 1.07D, 2.3D, 2.3D, 2.3D, 2.3D, 2.3D, 0D}, _
        ' {2.4D, 2.8D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.195D, 1.07D, 2.3D, 2.3D, 2.3D, 2.3D, 2.3D, 0D}, _
        ' {2.88D, 3.35D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.2D, 1.128D, 2.3D, 2.3D, 2.3D, 2.3D, 2.3D, 0D}, _
        ' {3.32D, 4.12D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.255D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {3.26D, 4.33D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.3D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        ' {3.32D, 4.12D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.255D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {3.8D, 5.25D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 4.048D, 0D}, _
        ' {6.12D, 4.76D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D, 0D}}
        Dim decCoeffList(,) As Decimal = {{6.9D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.2D, 11D, 11D, 11D, 11D, 11D, 0D}, _
            {6.4D, 4.46D, 1D, 1D, 19D, 0.93D, 0.667D, 1.3D, 1.2D, 11D, 11D, 11D, 11D, 11D, 0D}, _
            {6.17D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.2D, 1.3D, 10.3D, 10.2D, 10.3D, 10D, 10.6D, 0D}, _
            {6.4D, 4.46D, 1D, 1D, 19D, 0.93D, 0.667D, 1.3D, 1.2D, 11D, 11D, 11D, 11D, 11D, 0D}, _
            {5.37D, 4.13D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.2D, 8.4D, 8.4D, 8.4D, 8.4D, 8.4D, 0D}, _
            {4.15D, 3.3D, 0D, 0D, 19D, 0.82D, 0.6D, 1.23D, 1.3D, 5.9D, 5.9D, 5.9D, 5.9D, 5.9D, 0D}, _
            {3.373D, 4.34D, 1D, 1D, 19D, 0.6D, 0.6D, 1.233D, 1.374D, 4.4D, 4.4D, 4.3D, 4.4D, 4.4D, 0D}, _
            {2.977D, 3.96D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.17D, 1.243D, 3.7D, 3.7D, 3.7D, 3.7D, 3.7D, 0D}, _
            {2.88D, 4.57D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.2D, 1.214D, 3.55D, 3.55D, 3.55D, 3.55D, 3.55D, 0D}, _
            {3.01D, 3.7D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.137D, 1.14D, 3.74D, 3.74D, 3.74D, 3.74D, 3.74D, 0D}, _
            {2.607D, 3.01D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.067D, 1.083D, 3.14D, 3.14D, 3.14D, 3.14D, 3.14D, 0D}, _
            {2.103D, 2.437D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.045D, 1.046D, 2.46D, 2.46D, 2.46D, 2.46D, 2.46D, 0D}, _
            {1.74D, 2.144D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.966D, 0.97D, 2D, 2D, 2D, 2D, 2D, 0D}, _
            {1.74D, 2.144D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.966D, 0.97D, 2D, 2D, 2D, 2D, 2D, 0D}, _
            {1.74D, 1.45D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.95D, 0.9D, 2D, 2D, 2D, 2D, 2D, 0D}, _
            {2.103D, 2.437D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.045D, 1.046D, 2.46D, 2.46D, 2.46D, 2.46D, 2.46D, 0D}, _
            {2.103D, 2.437D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.045D, 1.046D, 2.46D, 2.46D, 2.46D, 2.46D, 2.46D, 0D}, _
            {2.607D, 3.01D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.067D, 1.083D, 3.14D, 3.14D, 3.14D, 3.14D, 3.14D, 0D}, _
            {3.01D, 3.7D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.137D, 1.14D, 3.74D, 3.74D, 3.74D, 3.74D, 3.74D, 0D}, _
            {2.977D, 3.96D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.17D, 1.243D, 3.7D, 3.7D, 3.7D, 3.7D, 3.7D, 0D}, _
            {3.01D, 3.7D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.137D, 1.14D, 3.74D, 3.74D, 3.74D, 3.74D, 3.74D, 0D}, _
            {3.373D, 4.34D, 1D, 1D, 19D, 0.6D, 0.6D, 1.233D, 1.374D, 4.4D, 4.4D, 4.3D, 4.4D, 4.4D, 0D}, _
            {5.37D, 4.13D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.2D, 8.4D, 8.4D, 8.4D, 8.4D, 8.4D, 0D}}

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

        GetTypeValues(Me.lHullTypeID, decNorm, lGuns, lDPS, lMaxHull, lHullAvail, lPower)

        Dim decTemp As Decimal = CDec(Math.Pow(((decDPS / CDec(lDPS))), 1.5)) * 675000000D
        'decTemp += (CDec(Math.Pow(blMaxRange, 1.1)) * 529411D)
        decTemp *= blMaxRange / 20D
        decTemp += CDec(Math.Pow((blScatterRadius * decDPS / 10D), 1.01D)) * 2100000D
        Return decTemp
    End Function
End Class
