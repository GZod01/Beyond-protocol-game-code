Option Strict On

Public Class BombTechComputer
    Inherits TechBuilderComputer

    Public yPayloadType As Byte
    Public decDPS As Decimal
    Public blAOE As Int64
    Public blGuidance As Int64
    Public blRange As Int64
    Public blPayloadSize As Int64
    Public blROF As Int64

    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
    '    '(1+(abs(Desired-Actual))xDPSx(SelfGuidancce+Range+1)/19
    '    Dim decTemp As Decimal = (1D + lMaxDA) * decDPS * (blGuidance + blRange + 1) / 19D
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
    '    'Dim sHeader As String = "BOMB"
    '    ml_POINT_PER_HULL = 1000 ' CInt(Val(oINI.GetString(sHeader, "PointPerHull", "1000")))
    '    ml_POINT_PER_POWER = 1000 'CInt(Val(oINI.GetString(sHeader, "PointPerPower", "1000")))
    '    ml_POINT_PER_RES_TIME = 100 ' CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "100")))
    '    ml_POINT_PER_RES_COST = 230 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "230")))
    '    ml_POINT_PER_PROD_TIME = 300 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "300")))
    '    ml_POINT_PER_PROD_COST = 180 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "180")))
    '    ml_POINT_PER_COLONIST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "1")))
    '    ml_POINT_PER_ENLISTED = 10 'CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "10")))
    '    ml_POINT_PER_OFFICER = 40 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "40")))

    '    ml_POINT_PER_MIN1 = 100 'CInt(Val(oINI.GetString(sHeader, "PointPerCasing", "100")))
    '    ml_POINT_PER_MIN2 = 70 'CInt(Val(oINI.GetString(sHeader, "PointPerGuidance", "70")))
    '    ml_POINT_PER_MIN3 = 120 'CInt(Val(oINI.GetString(sHeader, "PointPerPayload", "120")))


    '    'oINI.WriteString(sHeader, "PointPerHull", ml_POINT_PER_HULL.ToString)
    '    'oINI.WriteString(sHeader, "PointPerPower", ml_POINT_PER_POWER.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResTime", ml_POINT_PER_RES_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResCost", ml_POINT_PER_RES_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdTime", ml_POINT_PER_PROD_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdCost", ml_POINT_PER_PROD_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerColonist", ml_POINT_PER_COLONIST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerEnlisted", ml_POINT_PER_ENLISTED.ToString)
    '    'oINI.WriteString(sHeader, "PointPerOfficer", ml_POINT_PER_OFFICER.ToString)

    '    'oINI.WriteString(sHeader, "PointPerCasing", ml_POINT_PER_MIN1.ToString)
    '    'oINI.WriteString(sHeader, "PointPerGuidance", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerPayload", ml_POINT_PER_MIN3.ToString)

    '    'oINI = Nothing
    'End Sub


    'Protected Overrides Function CalculateBaseHull() As Int32
    '    Dim decTemp As Decimal = blAOE * 2D
    '    decTemp += blGuidance * 10D
    '    decTemp += blRange * 3D
    '    decTemp += blPayloadSize * 9D '1000D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBasePower() As Int32
    '    Dim decTemp As Decimal = blGuidance * 2D
    '    Dim decROF As Decimal = Math.Max(1, blROF) / 30D
    '    decTemp += (1D / decROF) * 1000D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResTime() As Int64
    '    Dim decTemp As Decimal = blAOE * 300000D
    '    decTemp += blGuidance * 990000D
    '    decTemp += blRange * 200000D
    '    decTemp += blPayloadSize * 740000D
    '    decTemp += decDPS * 100000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResCost() As Int64
    '    Dim decTemp As Decimal = blGuidance * 132000D
    '    decTemp += blRange * 134000D
    '    decTemp += decDPS * 120000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdTime() As Int64
    '    Dim decTemp As Decimal = blAOE * 100000D
    '    decTemp += blRange * 43000D
    '    decTemp += blPayloadSize * 100000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdCost() As Int64
    '    Dim decTemp As Decimal = blAOE * 100000D
    '    decTemp += blPayloadSize * 10000D
    '    decTemp += decDPS * 100000D
    '    If decTemp > Int64.MaxValue Then Return Int64.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseColonists() As Int32
    '    Dim decTemp As Decimal = blPayloadSize * 5D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseEnlisted() As Int32
    '    Dim decTemp As Decimal = decDPS * 0.5D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseOfficers() As Int32
    '    Dim decTemp As Decimal = decDPS * 0.1D
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return
        Dim lPayloadTypePropID As Int32 = 19    'combust
        Select Case yPayloadType
            Case 1
                lPayloadTypePropID = 16 'toxic
            Case 2
                lPayloadTypePropID = 13 'thermexp
            Case 3
                lPayloadTypePropID = 15 'mag prod
        End Select

        Dim yTemp(,) As Byte = {{203, 5, 10, 10, 5, 1, 150, 10, 183, 15, 201, 235}}

        ReDim muPropReqs(11)
        Dim lPropID() As Int32 = {1, 5, 12, 13, 19, 4, 3, 15, 10, 1, 5, lPayloadTypePropID}
        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(0, X)

            If X < 5 Then
                muPropReqs(X).lMinID = 0
            ElseIf X < 9 Then
                muPropReqs(X).lMinID = 1
            Else
                muPropReqs(X).lMinID = 2
            End If
        Next X
    End Sub

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
        'Dim decCoeffList() As Decimal = {2.4D, 3D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 0D}
        Dim decCoeffList() As Decimal = {2.637D, 3.46D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.033D, 1.12D, 3.1D, 3.1D, 3.1D, 0D, 0D, 0D}
        Return decCoeffList(lLookup)
    End Function

    Protected Overrides Function GetTheBill() As Decimal
        Dim decROF As Decimal = blROF / 30D
        Dim decNorm As Decimal = 0D
        Dim lGuns As Int32 = 0
        Dim lDPS As Int32 = 0
        Dim lMaxHull As Int32 = 0
        Dim lHullAvail As Int32 = 0
        Dim lPower As Int32 = 0
        TechBuilderComputer.GetTypeValues(Me.lHullTypeID, decNorm, lGuns, lDPS, lMaxHull, lHullAvail, lPower)

        'Dim decTemp As Decimal = CDec(Math.Pow((((blPayloadSize / decROF) / 50D) / (lGuns * lDPS)), 1.1))
        Dim decTemp As Decimal = CDec(Math.Pow(((decDPS / 50D) / CDec(lDPS)), 1.1))
        decTemp *= 675000000
        decTemp += (blRange * 629411)
        decTemp += CDec(Math.Pow((blAOE * blPayloadSize), 1.1)) * 1600
        decTemp += (blGuidance * 500000)

        Return decTemp
    End Function
End Class
