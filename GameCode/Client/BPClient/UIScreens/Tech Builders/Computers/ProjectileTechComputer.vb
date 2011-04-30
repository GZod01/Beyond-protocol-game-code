Option Strict On

Public Class ProjectileTechComputer
    Inherits TechBuilderComputer

    Public yPayloadType As Byte
    Public yProjectionType As Byte
    Public decDPS As Decimal
    Public blPierceRatio As Int64
    Public blMaxRange As Int64
    Public blExplosionRadius As Int64
    Public blCartridgeSize As Int64
    Public blROF As Int64

    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
    '    '(1+abs(Desired-Actual))xDPSx(Pierce+Range+ExplosionRadius+1)/30
    '    Dim blPierceRatioMult As Int64 = (50 - Math.Abs(blPierceRatio - 50)) * 2
    '    Dim decTemp As Decimal = (1D + lMaxDA) * decDPS * (blPierceRatioMult + blMaxRange + blExplosionRadius + 1) / 30D
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
    '    'Dim sHeader As String = "PROJECTILE"
    '    ml_POINT_PER_HULL = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerHull", "1000")))
    '    ml_POINT_PER_POWER = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerPower", "1000")))
    '    ml_POINT_PER_RES_TIME = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "100")))
    '    ml_POINT_PER_RES_COST = 230 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "230")))
    '    ml_POINT_PER_PROD_TIME = 300 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "300")))
    '    ml_POINT_PER_PROD_COST = 180 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "180")))
    '    ml_POINT_PER_COLONIST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "1")))
    '    ml_POINT_PER_ENLISTED = 10 ' CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "10")))
    '    ml_POINT_PER_OFFICER = 40 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "40")))

    '    ml_POINT_PER_MIN1 = 70 'CInt(Val(oINI.GetString(sHeader, "PointPerBarrel", "70")))
    '    ml_POINT_PER_MIN2 = 120 'CInt(Val(oINI.GetString(sHeader, "PointPerCasing", "120")))
    '    ml_POINT_PER_MIN3 = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerPayload1", "100")))
    '    ml_POINT_PER_MIN4 = 90 'CInt(Val(oINI.GetString(sHeader, "PointPerPayload2", "90")))
    '    ml_POINT_PER_MIN5 = 110 'CInt(Val(oINI.GetString(sHeader, "PointPerProj", "110")))
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

    '    'oINI.WriteString(sHeader, "PointPerBarrel", ml_POINT_PER_MIN1.ToString)
    '    'oINI.WriteString(sHeader, "PointPerCasing", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerPayload1", ml_POINT_PER_MIN3.ToString)
    '    'oINI.WriteString(sHeader, "PointPerPayload2", ml_POINT_PER_MIN4.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProj", ml_POINT_PER_MIN5.ToString)

    '    'oINI = Nothing
    'End Sub

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return


        Dim lPayload2PropID As Int32 = 1
        Select Case yPayloadType
            Case 1
                lPayload2PropID = 19
            Case 2
                lPayload2PropID = 18
            Case 3
                lPayload2PropID = 15
        End Select
        Dim lProjectionPropID As Int32 = 19
        If yProjectionType = 1 Then
            lProjectionPropID = 15
        End If

        'need projection

        Dim yTemp(,) As Byte = {{0, 0, 4, 1, 5, 0, 6, 4, 0, 6, 5}, _
            {0, 0, 4, 1, 5, 0, 6, 4, 0, 6, 5}, _
            {1, 0, 4, 1, 6, 0, 6, 4, 0, 6, 6}, _
            {1, 0, 4, 1, 6, 0, 6, 4, 0, 6, 6}, _
            {1, 0, 4, 1, 6, 0, 6, 4, 0, 6, 6}, _
            {1, 0, 4, 1, 6, 0, 6, 4, 0, 7, 6}, _
            {1, 0, 4, 1, 6, 0, 6, 4, 0, 7, 6}, _
            {1, 0, 5, 1, 6, 0, 7, 5, 0, 7, 6}, _
            {1, 0, 5, 1, 6, 0, 7, 5, 0, 7, 6}, _
            {2, 0, 5, 1, 6, 1, 7, 5, 1, 7, 6}, _
            {2, 1, 5, 1, 7, 1, 7, 5, 1, 7, 7}, _
            {2, 1, 5, 1, 7, 1, 7, 5, 1, 7, 7}, _
            {2, 1, 5, 1, 7, 1, 7, 5, 1, 8, 7}, _
            {3, 1, 5, 1, 7, 1, 7, 5, 1, 8, 7}, _
            {3, 1, 5, 1, 7, 1, 8, 5, 1, 8, 7}, _
            {3, 1, 5, 1, 7, 1, 8, 5, 1, 8, 7}, _
            {4, 1, 5, 2, 7, 1, 8, 6, 1, 8, 8}, _
            {4, 1, 6, 2, 8, 2, 8, 6, 2, 8, 8}, _
            {4, 2, 6, 2, 8, 2, 8, 6, 2, 9, 8}, _
            {5, 2, 6, 2, 8, 2, 8, 6, 2, 9, 8}, _
            {5, 2, 6, 2, 8, 2, 9, 6, 2, 9, 8}}

        ReDim muPropReqs(10)
        Dim lPropID() As Int32 = {11, 13, 12, 11, 1, 12, 3, 2, 1, lPayload2PropID, lProjectionPropID}
        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)

            If X < 3 Then
                muPropReqs(X).lMinID = 0
            ElseIf X < 6 Then
                muPropReqs(X).lMinID = 1
            ElseIf X < 9 Then
                muPropReqs(X).lMinID = 2
            ElseIf X < 10 Then
                muPropReqs(X).lMinID = 3
            Else
                muPropReqs(X).lMinID = 4
            End If
        Next X
    End Sub

    'Protected Overrides Function CalculateBaseHull() As Int32
    '    Dim decTemp As Decimal = 0D

    '    Dim blPierceRatioMult As Int64 = (50 - Math.Abs(blPierceRatio - 50)) * 2

    '    If yProjectionType = 0 Then
    '        decTemp += blCartridgeSize * 4D
    '        decTemp += blPierceRatioMult * 2D
    '    Else
    '        decTemp += blCartridgeSize * 2D
    '        decTemp += blPierceRatioMult
    '    End If
    '    decTemp += decDPS * 4D
    '    If yPayloadType <> 0 Then decTemp += CDec(yPayloadType) * 2 * decDPS

    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBasePower() As Int32
    '    Dim decTemp As Decimal = 0D
    '    Dim blPierceRatioMult As Int64 = (50 - Math.Abs(blPierceRatio - 50)) * 2
    '    If yProjectionType = 1 Then
    '        decTemp += blCartridgeSize * 10D
    '        decTemp += blPierceRatioMult * 10D
    '    Else
    '        decTemp += blCartridgeSize * 2D
    '        decTemp += blPierceRatioMult * 2D
    '    End If
    '    decTemp += (decDPS * 2D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResTime() As Int64
    '    Dim blPierceRatioMult As Int64 = (50 - Math.Abs(blPierceRatio - 50)) * 2
    '    Dim decTemp As Decimal = blPierceRatioMult * 1000000D
    '    If yProjectionType = 0 Then
    '        decTemp += blMaxRange * 5000000D
    '        decTemp += blExplosionRadius * 1000000D
    '    Else
    '        decTemp += blMaxRange * 2000000D
    '        decTemp += blExplosionRadius * 10000000D
    '    End If
    '    If yPayloadType <> 0 Then
    '        decTemp += CDec(yPayloadType) * 1000000D * decDPS
    '    End If
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResCost() As Int64
    '    Dim blPierceRatioMult As Int64 = (50 - Math.Abs(blPierceRatio - 50)) * 2
    '    Dim decTemp As Decimal = blCartridgeSize * 1000000
    '    decTemp += blPierceRatioMult * 800000D
    '    Const dec_MIN_ROF As Decimal = 1 / 30D
    '    Dim decROF As Decimal = Math.Max(dec_MIN_ROF, blROF / 30D)
    '    decTemp += (1D / decROF) * 10000000D
    '    If decTemp < 0 Then Return 0
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdTime() As Int64
    '    Dim blPierceRatioMult As Int64 = (50 - Math.Abs(blPierceRatio - 50)) * 2
    '    Dim decTemp As Decimal = decDPS * 1200000D
    '    If yProjectionType = 1 Then
    '        decTemp += blCartridgeSize * 1500000D
    '        decTemp += blPierceRatioMult * 2000000D
    '    Else
    '        decTemp += blCartridgeSize * 1000000D
    '        decTemp += blPierceRatioMult * 1000000D
    '    End If
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdCost() As Int64
    '    Dim blPierceRatioMult As Int64 = (50 - Math.Abs(blPierceRatio - 50)) * 2
    '    Dim decTemp As Decimal = 0
    '    If blROF < 1 Then blROF = 1
    '    Dim decROF As Decimal = 30D / blROF

    '    If yProjectionType > 1 Then
    '        decTemp += blCartridgeSize * 1500000D
    '        decTemp += blPierceRatioMult * 1500000D
    '        decTemp += decROF * 1500000D
    '        decTemp += blExplosionRadius * 1500000D
    '    Else
    '        decTemp += blCartridgeSize * 1000000D
    '        decTemp += blPierceRatioMult * 1000000D
    '        decTemp += decROF * 1000000D
    '        decTemp += blExplosionRadius * 1000000D
    '    End If
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseColonists() As Int32
    '    Dim decTemp As Decimal = blExplosionRadius * 2D
    '    If yProjectionType = 1 Then
    '        decTemp += blCartridgeSize * 0.5D
    '    Else
    '        decTemp += blCartridgeSize
    '    End If
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseEnlisted() As Int32
    '    Dim decTemp As Decimal = (blCartridgeSize * 0.2D) + (blExplosionRadius * 0.1D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseOfficers() As Int32
    '    Dim decTemp As Decimal = (blCartridgeSize * 0.05D) + (blExplosionRadius * 0.5D)
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
 
        'Dim decCoeffList(,) As Decimal = {{6.9D, 4D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 14D, 14D, 14D, 14D, 14D, 0D}, _
        ' {7.8D, 4.6D, 1D, 1D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 5.9D, 6.3D, 6.4D, 6D, 6D, 0D}, _
        ' {6.1D, 3.85D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 11.8D, 11.8D, 11.8D, 11.8D, 11.8D, 0D}, _
        ' {7.8D, 4.6D, 1D, 1D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 5.9D, 6.3D, 6.4D, 6D, 6D, 0D}, _
        ' {4.87D, 3.74D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 6.8D, 6.9D, 7.1D, 7.5D, 7.2D, 0D}, _
        ' {4.3D, 3D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 6.5D, 6.5D, 6.4D, 6.4D, 6.2D, 0D}, _
        ' {3.155D, 4.8D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 4.048D, 0D}, _
        ' {2.93D, 4D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.16D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        ' {2.85D, 4.5D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.2D, 1.214D, 3D, 3D, 3D, 3D, 3D, 0D}, _
        ' {3.02D, 3.72D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.139D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {2.59D, 2.98D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.08D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {2.13D, 2.48D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.06D, 1.015D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {1.8D, 2.2D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1D, 1D, 2D, 2D, 2D, 2D, 2D, 0D}, _
        ' {1.8D, 2.2D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1D, 1D, 2D, 2D, 2D, 2D, 2D, 0D}, _
        ' {1.8D, 2.2D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1D, 1D, 2D, 2D, 2D, 2D, 2D, 0D}, _
        ' {2.13D, 2.48D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.06D, 1.015D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {2.13D, 2.48D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.06D, 1.015D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {2.59D, 2.98D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.08D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 0D}, _
        ' {3.02D, 3.72D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.139D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {2.93D, 4D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.16D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}, _
        ' {3.02D, 3.72D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.139D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 0D}, _
        ' {3.155D, 4.8D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 4.048D, 4.048D, 4.048D, 4.048D, 4.048D, 0D}, _
        ' {4.87D, 3.74D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 6.8D, 6.9D, 7.1D, 7.5D, 7.2D, 0D}}
        Dim decCoeffList(,) As Decimal = {{6.4D, 3.67D, 0D, 0D, 19D, 0.93D, 0.667D, 1.2D, 1.39D, 12D, 12D, 12D, 12D, 12D, 0D}, _
            {6D, 4.1D, 1D, 1D, 19D, 0.82D, 0.6D, 1.25D, 1.36D, 9D, 9D, 9D, 9D, 9D, 0D}, _
            {5.65D, 3.66D, 0D, 0D, 19D, 0.93D, 0.667D, 1.2D, 1.2D, 9.5D, 9.5D, 9.5D, 9.5D, 9.5D, 0D}, _
            {6D, 4.1D, 1D, 1D, 19D, 0.82D, 0.6D, 1.25D, 1.36D, 9D, 9D, 9D, 9D, 9D, 0D}, _
            {4.96D, 3.8D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.5D, 7.5D, 7.7D, 7.5D, 7.7D, 0D}, _
            {3.85D, 3.03D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 5.5D, 5.5D, 5.5D, 5.5D, 5.5D, 0D}, _
            {3.127D, 3.99D, 1D, 1D, 19D, 0.6D, 0.6D, 1.14D, 1.28D, 4.1D, 3.95D, 4.1D, 4.1D, 4D, 0D}, _
            {2.762D, 3.64D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.084D, 1.158D, 3.43D, 3.43D, 3.43D, 3.43D, 3.43D, 0D}, _
            {2.68D, 4.2D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.13D, 1.15D, 3.3D, 3.3D, 3.3D, 3.3D, 3.3D, 0D}, _
            {2.8D, 3.42D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.053D, 1.07D, 3.45D, 3.45D, 3.45D, 3.45D, 3.45D, 0D}, _
            {2.42D, 2.81D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.009D, 1.038D, 2.92D, 2.92D, 2.92D, 2.92D, 2.92D, 0D}, _
            {1.954D, 2.245D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.968D, 0.976D, 2.295D, 2.295D, 2.295D, 2.295D, 2.295D, 0D}, _
            {1.618D, 1.975D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.896D, 0.904D, 1.87D, 1.87D, 1.87D, 1.87D, 1.87D, 0D}, _
            {1.618D, 1.975D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.896D, 0.904D, 1.87D, 1.87D, 1.87D, 1.87D, 1.87D, 0D}, _
            {2.2D, 2D, 1.6D, 2.1D, 2.4D, 0.46D, 0.43D, 1D, 1.27D, 2.7D, 2.7D, 2.7D, 2.7D, 2.7D, 0D}, _
            {1.954D, 2.245D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.968D, 0.976D, 2.295D, 2.295D, 2.295D, 2.295D, 2.295D, 0D}, _
            {1.954D, 2.245D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.968D, 0.976D, 2.295D, 2.295D, 2.295D, 2.295D, 2.295D, 0D}, _
            {2.42D, 2.81D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.009D, 1.038D, 2.92D, 2.92D, 2.92D, 2.92D, 2.92D, 0D}, _
            {2.8D, 3.42D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.053D, 1.07D, 3.45D, 3.45D, 3.45D, 3.45D, 3.45D, 0D}, _
            {2.762D, 3.64D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.084D, 1.158D, 3.43D, 3.43D, 3.43D, 3.43D, 3.43D, 0D}, _
            {2.8D, 3.42D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.053D, 1.07D, 3.45D, 3.45D, 3.45D, 3.45D, 3.45D, 0D}, _
            {3.127D, 3.99D, 1D, 1D, 19D, 0.6D, 0.6D, 1.14D, 1.28D, 4.1D, 3.95D, 4.1D, 4.1D, 4D, 0D}, _
            {4.96D, 3.8D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.5D, 7.5D, 7.7D, 7.5D, 7.7D, 0D}}

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
        Dim lHullDPS As Int32 = 0
        Dim lMaxHull As Int32 = 0
        Dim lHullAvail As Int32 = 0
        Dim lPower As Int32 = 0

        GetTypeValues(Me.lHullTypeID, decNorm, lGuns, lHullDPS, lMaxHull, lHullAvail, lPower)

        Dim decROF As Decimal = blROF / 30D
        'Dim decTemp As Decimal = CDec(Math.Pow(((blCartridgeSize / decROF) / (lDPS * lGuns)), 1.4)) * 675000000D
        Dim decTemp As Decimal = CDec(Math.Pow((decDPS / CDec(lHullDPS)), 1.4)) * 675000000D
        'decTemp += (blMaxRange * 529411)
        decTemp *= blMaxRange / 18D
        decTemp += CDec(Math.Pow(blCartridgeSize * blExplosionRadius, 1.2) * 2100000D)
        Return decTemp
    End Function
End Class
