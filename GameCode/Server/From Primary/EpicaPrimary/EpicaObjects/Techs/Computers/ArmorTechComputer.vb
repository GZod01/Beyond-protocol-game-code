Option Strict On

Public Class ArmorTechComputer
    Inherits TechBuilderComputer

    Public blImpact As Int64
    Public blBeam As Int64
    Public blPiercing As Int64
    Public blMagnetic As Int64
    Public blToxic As Int64
    Public blBurn As Int64
    Public blHullUsagePerPlate As Int64
    Public blHPPerPlate As Int64

    Public lHPRatioSpecial As Int32 = 10

    Protected ml_POINT_PER_HULL As Int32
    Protected ml_POINT_PER_POWER As Int32
    Protected ml_POINT_PER_RES_TIME As Int32
    Protected ml_POINT_PER_RES_COST As Int32
    Protected ml_POINT_PER_PROD_TIME As Int32
    Protected ml_POINT_PER_PROD_COST As Int32
    Protected ml_POINT_PER_COLONIST As Int32
    Protected ml_POINT_PER_ENLISTED As Int32
    Protected ml_POINT_PER_OFFICER As Int32

    Protected ml_POINT_PER_MIN1 As Int32
    Protected ml_POINT_PER_MIN2 As Int32
    Protected ml_POINT_PER_MIN3 As Int32
    Protected ml_POINT_PER_MIN4 As Int32
    Protected ml_POINT_PER_MIN5 As Int32
    Protected ml_POINT_PER_MIN6 As Int32

    Protected Function CalculateBaseColonists() As Integer
        Return 0
    End Function

    Protected Function CalculateBaseEnlisted() As Integer
        Return 0
    End Function

    Protected Function CalculateBaseHull() As Integer
        Return 0
    End Function

    Protected Function CalculateBaseMineralCosts(ByVal lMaxDA As Integer, ByVal lSumDAForMineral As Int32) As Integer
        'Dim decSumResist As Decimal = blImpact + blBeam + blPiercing + blMagnetic + blToxic + blBurn
        'Dim decTemp As Decimal = (1D + lMaxDA) * (decSumResist + blHPPerPlate + 2) * blHullUsagePerPlate
        Dim decTemp As Decimal = blHullUsagePerPlate '(1D + lMaxDA) * blHullUsagePerPlate
        If decTemp > Int32.MaxValue Then Return Int32.MaxValue
        If decTemp < 0 Then Return 0
        Return CInt(decTemp)
    End Function

    Protected Function CalculateBaseOfficers() As Integer
        Return 0
    End Function

    Protected Function CalculateBasePower() As Integer
        Return 0
    End Function

    Protected Function CalculateBaseProdCost() As Long
        Dim decTemp As Decimal = blImpact
        decTemp += blBeam
        decTemp += blPiercing
        decTemp += blMagnetic
        decTemp += blToxic
        decTemp += blBurn
        decTemp += blHPPerPlate
        decTemp /= CDec(blHullUsagePerPlate)

        Dim decHPPerHull As Decimal = CDec(blHPPerPlate / blHullUsagePerPlate)
        If decHPPerHull > lHPRatioSpecial Then
            decTemp += (decHPPerHull - lHPRatioSpecial) * 10000
        End If

        If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        If decTemp < 0 Then Return 0
        Return CLng(decTemp)
    End Function

    Protected Function CalculateBaseProdTime() As Long
        Dim decTemp As Decimal = blImpact
        decTemp += blBeam
        decTemp += blPiercing
        decTemp += blMagnetic
        decTemp += blToxic
        decTemp += blBurn
        decTemp += blHPPerPlate
        decTemp /= CDec(blHullUsagePerPlate)
        decTemp *= 1000D

        Dim decHPPerHull As Decimal = CDec(blHPPerPlate / blHullUsagePerPlate)
        If decHPPerHull > lHPRatioSpecial Then
            decTemp += (180000 * (blHullUsagePerPlate + 1))
        End If

        If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        If decTemp < 0 Then Return 0
        Return CLng(decTemp)
    End Function

    Protected Function CalculateBaseResCost() As Long
        Dim decTemp As Decimal = 0
        decTemp += CDec(blHPPerPlate / blHullUsagePerPlate) * 1000000D
        decTemp += blBeam * 200000D
        decTemp += blImpact * 50000D
        decTemp += blPiercing * 100000D
        decTemp += blMagnetic * 1000000D
        decTemp += blToxic * 500000D
        decTemp += blBurn * 300000D
        If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        If decTemp < 0 Then Return 0
        Return CLng(decTemp)
    End Function

    Protected Function CalculateBaseResTime() As Long
        Dim decTemp As Decimal = 0
        decTemp += CDec(blHPPerPlate / blHullUsagePerPlate) * 10000000D
        decTemp += blBeam * 1000000D
        decTemp += blImpact * 1000000D
        decTemp += blPiercing * 1000000D
        decTemp += blMagnetic * 2000000D
        decTemp += blToxic * 3000000D
        decTemp += blBurn * 500000D
        If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        If decTemp < 0 Then Return 0
        Return CLng(decTemp)
    End Function

    Protected Function CalculateThresholdValue() As Decimal
        Dim decTemp As Decimal = 0 'blImpact
        'decTemp += blBeam
        'decTemp += blPiercing
        'decTemp += blMagnetic
        'decTemp += blToxic
        'decTemp += blBurn
        'decTemp = Math.Max(1, decTemp - 100)
        'decTemp = (blHPPerPlate + 1) / CDec(blHullUsagePerPlate) * (decTemp / 15 + 1)
        decTemp = ((blHPPerPlate) / CDec(blHullUsagePerPlate)) * 5
        Return decTemp
    End Function

    Protected Sub LoadPointVals()
        'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        'If sPath.EndsWith("\") = False Then sPath &= "\"
        'sPath &= "TECH.dat"
        'Dim oINI As New InitFile(sPath)
        'Dim sHeader As String = "ARMOR"
        ml_POINT_PER_HULL = 0 'CInt(Val(oINI.GetString(sHeader, "PointPerHull", "600000")))
        ml_POINT_PER_POWER = 0
        ml_POINT_PER_RES_TIME = 5 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "5")))
        ml_POINT_PER_RES_COST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "1")))
        ml_POINT_PER_PROD_TIME = 100000 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "100000")))
        ml_POINT_PER_PROD_COST = 720000 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "720000")))
        ml_POINT_PER_COLONIST = 0 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "6")))
        ml_POINT_PER_ENLISTED = 0 'CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "0")))
        ml_POINT_PER_OFFICER = 0 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "0")))

        ml_POINT_PER_MIN1 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerOuter", "600")))
        ml_POINT_PER_MIN2 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerMiddle", "600")))
        ml_POINT_PER_MIN3 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerInner", "600")))
        ml_POINT_PER_MIN4 = 0
        ml_POINT_PER_MIN5 = 0
        ml_POINT_PER_MIN6 = 0

        'oINI.WriteString(sHeader, "PointPerResTime", ml_POINT_PER_RES_TIME.ToString)
        'oINI.WriteString(sHeader, "PointPerResCost", ml_POINT_PER_RES_COST.ToString)
        'oINI.WriteString(sHeader, "PointPerProdTime", ml_POINT_PER_PROD_TIME.ToString)
        'oINI.WriteString(sHeader, "PointPerProdCost", ml_POINT_PER_PROD_COST.ToString)
        'oINI.WriteString(sHeader, "PointPerOuter", ml_POINT_PER_MIN1.ToString)
        'oINI.WriteString(sHeader, "PointPerMiddle", ml_POINT_PER_MIN2.ToString)
        'oINI.WriteString(sHeader, "PointPerInner", ml_POINT_PER_MIN3.ToString)
        'oINI = Nothing
    End Sub

    Protected Overrides Sub SetMinReqProps()

        ReDim muPropReqs(36)
        Dim lIdx As Int32 = -1

        If blImpact > 0 Then
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 5 : .lPropVal = 83 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 1 : .lPropVal = 100 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 2 : .lPropVal = 220 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 14 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 1 : .lPropVal = 173 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 2 : .lPropID = 3 : .lPropVal = 11 : End With
        End If
        If blPiercing > 0 Then
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 5 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 1 : .lPropVal = 123 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 2 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 3 : .lPropVal = 34 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 5 : .lPropVal = 11 : End With
        End If
        If blMagnetic > 0 Then
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 1 : .lPropVal = 34 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 15 : .lPropVal = 173 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 14 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 3 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 15 : .lPropVal = 123 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 2 : .lPropID = 1 : .lPropVal = 11 : End With
        End If
        If blToxic > 0 Then
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 18 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 1 : .lPropVal = 173 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 18 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 1 : .lPropVal = 152 : End With
        End If
        If blBurn > 0 Then
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 19 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 1 : .lPropVal = 100 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 19 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 1 : .lPropVal = 173 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 12 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 13 : .lPropVal = 100 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 2 : .lPropID = 19 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 2 : .lPropID = 12 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 2 : .lPropID = 13 : .lPropVal = 11 : End With
        End If
        If blBeam > 0 OrElse lIdx = -1 Then
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 15 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 6 : .lPropVal = 173 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 0 : .lPropID = 7 : .lPropVal = 152 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 12 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 1 : .lPropID = 13 : .lPropVal = 11 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 2 : .lPropID = 4 : .lPropVal = 83 : End With
            lIdx += 1
            With muPropReqs(lIdx) : .lMinID = 2 : .lPropID = 13 : .lPropVal = 11 : End With
        End If

        ReDim Preserve muPropReqs(lIdx)
    End Sub

    Public Function GetThresholdModifier() As Decimal
        Dim decTemp As Decimal = 0 ' blImpact
        'decTemp += blBeam
        'decTemp += blPiercing
        'decTemp += blMagnetic
        'decTemp += blToxic
        'decTemp += blBurn
        'decTemp = (blHPPerPlate + 1) / CDec(blHullUsagePerPlate) * (decTemp + 1) - 2000
        decTemp = ((blHPPerPlate) / CDec(blHullUsagePerPlate)) * 5
        decTemp *= decTemp
        Return decTemp
    End Function

    Protected Shared Function InvalidHalfBaseCheck(ByVal blBaseValue As Int64, ByVal blLockedValue As Int64, ByVal yType As Byte) As Boolean
        Dim blMinVal As Int64 = 0
        If yType = 0 Then
            blMinVal = blBaseValue \ 2
            If blMinVal < 1 Then blMinVal = 1
        ElseIf yType = 1 Then
            blMinVal = CLng(blBaseValue * 0.1F)
            If blMinVal < 0 Then blMinVal = 0
        ElseIf yType = 2 Then
            blMinVal = CLng(Math.Ceiling(blBaseValue * 0.25F))
            If blMinVal < 1 Then blMinVal = 1
        Else
            blMinVal = CLng(Math.Min(1, blBaseValue))
        End If

        If blLockedValue < blMinVal Then
            Return True
        End If

        Return False
        'If yType = 0 Then
        '    If blLockedValue < blBaseValue / 2 Then
        '        Return True
        '    End If
        'ElseIf yType = 1 Then
        '    If blLockedValue < blBaseValue * 0.01F Then
        '        Return True
        '    End If
        'ElseIf yType = 2 Then
        '    If blLockedValue < blBaseValue * 0.25F Then
        '        Return True
        '    End If
        'Else
        '    If blLockedValue < Math.Min(1, blBaseValue) Then
        '        Return True
        '    End If
        'End If

        'Return False
    End Function

    Protected Shared Function DoLockedToBaseCheck(ByRef lLockedValue As Int32, ByRef lBaseValue As Int32) As Int32
        Dim lResult As Int32 = 0
        If lLockedValue = -1 Then
            'value is not locked, set the value so we can set it later...
            'lLockedValue = lBaseValue
            'and clear the lBaseHull value
            'lBaseValue = 0
            'lLockedValue = 0
            lResult = 0
        Else
            'value is locked, determine our difference
            lResult = lBaseValue - lLockedValue
            'If lResult > 0 Then
            '    Dim decTemp As Decimal = CDec(lResult) * CDec(lResult)
            '    lResult = CInt(Math.Max(Int32.MinValue, Math.Min(Int32.MaxValue, decTemp)))
            'End If
            'Dim decTemp As Decimal = CDec(lResult) * CDec(lResult)
            'If lResult < 0 Then
            '    decTemp = -decTemp
            'End If
            'lResult = CInt(Math.Max(Int32.MinValue, Math.Min(Int32.MaxValue, decTemp)))

            'Dim decTemp As Decimal = lBaseValue - lLockedValue
            'Dim lMult As Int32 = 1
            'If decTemp < 1 Then lMult = -1
            'decTemp = CDec(Math.Pow(Math.Abs(decTemp), 1.2)) * lMult
            'lBaseValue = CInt(Math.Max(Int32.MinValue, Math.Min(Int32.MaxValue, decTemp)))

            ''Trying this on for size...
            ''  (BaseValue - DesiredValue) ^ (1 + (1 - (DesiredValue / BaseValue)))
            'Dim decTemp As Decimal = (lBaseValue - lLockedValue)
            'Dim lTemp As Int32 = Math.Max(1, lBaseValue)
            'decTemp = CDec(Math.Pow(decTemp, 1 + (1 - (lLockedValue / lTemp))))
            'If decTemp > Int32.MaxValue Then
            '    lBaseValue = Int32.MaxValue
            'ElseIf decTemp < Int32.MinValue Then
            '    lBaseValue = Int32.MinValue
            'Else : lBaseValue = CInt(decTemp)
            'End If
        End If
        'AddToTestString("L2B Result: " & lResult.ToString)
        Return lResult
    End Function
    Protected Shared Function DoLockedToBaseCheck(ByRef lLockedValue As Int64, ByRef lBaseValue As Int64) As Int64
        Dim blResult As Int64 = 0
        If lLockedValue = -1 Then
            'value is not locked, set the value so we can set it later...
            'lLockedValue = lBaseValue
            'and clear the lBaseHull value
            'lBaseValue = 0
            'lLockedValue = 0
            blResult = 0
        Else
            'value is locked, determine our difference
            blResult = lBaseValue - lLockedValue
            'If blResult > 0 Then
            '    Dim decTemp As Decimal = CDec(blResult) * CDec(blResult)
            '    blResult = CLng(Math.Max(Int64.MinValue, Math.Min(Int64.MaxValue, decTemp)))
            'End If
            'Dim decTemp As Decimal = CDec(blResult) * CDec(blResult)
            'If blResult < 0 Then
            '    decTemp = -decTemp
            'End If
            'blResult = CLng(Math.Max(Int64.MinValue, Math.Min(Int64.MaxValue, decTemp)))

            'Dim decTemp As Decimal = lBaseValue - lLockedValue
            'Dim lMult As Int32 = 1
            'If decTemp < 1 Then lMult = -1
            'decTemp = CDec(Math.Pow(Math.Abs(decTemp), 1.2)) * lMult
            'lBaseValue = CLng(Math.Max(Int64.MinValue, Math.Min(Int64.MaxValue, decTemp)))

            ''trying this on for size...
            ''  (BaseValue - DesiredValue) ^ (1 + (1 - (DesiredValue / BaseValue)))
            'Dim decTemp As Decimal = (lBaseValue - lLockedValue)
            'Dim lTemp As Int64 = Math.Max(1, lBaseValue)
            'decTemp = CDec(Math.Pow(decTemp, 1 + (1 - (lLockedValue / lTemp))))
            'If decTemp > Int64.MaxValue Then
            '    lBaseValue = Int64.MaxValue
            'ElseIf decTemp < Int64.MinValue Then
            '    lBaseValue = Int64.MinValue
            'Else : lBaseValue = CLng(decTemp)
            'End If
        End If
        'AddToTestString("L2B Result: " & blResult.ToString)
        Return blResult
    End Function

    'Protected Shared Sub AddToTestString(ByVal sLine As String)
    '    If moSB Is Nothing Then moSB = New System.Text.StringBuilder
    '    moSB.AppendLine(sLine)
    'End Sub

    Protected Shared Sub AffectValueByPoints(ByRef blPoints As Int64, ByVal lPointPerValue As Int32, ByRef lValue As Int32, ByVal lModByVal As Int32)
        If blPoints < lPointPerValue Then Return
        If lValue = Int32.MaxValue Then Return
        lValue += lModByVal
        blPoints -= lPointPerValue
    End Sub
    Protected Shared Sub AffectValueByPoints(ByRef blPoints As Int64, ByVal lPointPerValue As Int32, ByRef blValue As Int64, ByVal lModByVal As Int32)
        If blPoints < lPointPerValue Then Return
        If blValue = Int64.MaxValue Then Return
        blValue += lModByVal
        blPoints -= lPointPerValue
    End Sub

    Private Function GetPointValue(ByVal blVal As Int64, ByVal sName As String, ByVal blPerPointMult As Int64) As Int64
        Dim blResult As Int64 = 0
        If blVal < 0 Then
            blVal = Math.Abs(blVal * blPerPointMult)

            'If sName.ToUpper = "RESTIME" Then
            '    blVal \= 180000
            'ElseIf sName.ToUpper = "PRODTIME" Then
            '    blVal \= 900000
            'End If

            blResult = -CLng((1D / Math.Log10(blVal)) * blVal)  '-CLng((1 / Math.Log10(blVal)) * blVal)
            blVal *= blPerPointMult

            'If sName.ToUpper = "RESTIME" Then
            '    blResult *= 180000
            'ElseIf sName.ToUpper = "PRODTIME" Then
            '    blResult *= 900000
            'End If
            'AddToTestString(sName & " grants " & blResult.ToString("#,##0"))
        Else
            blResult = blVal * blPerPointMult
        End If
        Return blResult
    End Function

    Public Function ArmorBuilderCostValueChange(ByVal iTechTypeID As Int16, ByVal decMaxThreshold As Decimal, ByVal lPopIntel As Int32) As Boolean
        'If moSB Is Nothing = False Then moSB.Length = 0

        Try
            LoadPointVals()
            SetMinReqProps()

            Dim decNormalize As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1

            If lHullTypeID = -1 AndAlso iTechTypeID <> ObjectType.eArmorTech Then
                'lblDesignFlaw.Caption = "Invalid Design: a Hull Type must be selected!"
                'mbIgnoreValueChange = False
                Return False
            End If

            GetTypeValues(lHullTypeID, decNormalize, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

            'AddToTestString("N: " & decNormalize.ToString & ", MG: " & lMaxGuns.ToString & ", MDPS: " & lMaxDPS & ", MH: " & lMaxHullSize & ", HA: " & lHullAvail & ", MP: " & lMaxPower)

            'Ok, something has changed, let's calculate our values
            'This occurs in the following steps/order
            ' 1) Determine the required values lookups for the mineral properties
            ' 2) Calculate all base values for the costs
            '   a) Hull
            Dim lBaseHull As Int32 = CalculateBaseHull()
            'AddToTestString("Basehull: " & lBaseHull.ToString)
            '   b) Power - we do not use power here
            Dim lBasePower As Int32 = CalculateBasePower()
            'AddToTestString("lBasePower: " & lBasePower.ToString)
            '   c) Colonists
            Dim lBaseColonists As Int32 = CalculateBaseColonists()
            'AddToTestString("BaseColonists: " & lBaseColonists.ToString)
            '   d) Enlisted
            Dim lBaseEnlisted As Int32 = CalculateBaseEnlisted()
            'AddToTestString("BaseEnlisted: " & lBaseEnlisted.ToString)
            '   e) Officers
            Dim lBaseOfficers As Int32 = CalculateBaseOfficers()
            'AddToTestString("BaseOfficers: " & lBaseOfficers.ToString)
            '   f) Research Cost
            Dim blBaseResCost As Int64 = CalculateBaseResCost()
            'AddToTestString("BaseResCost: " & blBaseResCost.ToString)
            '   g) Research Time
            Dim blBaseResTime As Int64 = CalculateBaseResTime()
            'AddToTestString("BaseResTime: " & blBaseResTime.ToString)
            '   h) Production Cost
            Dim blBaseProdCost As Int64 = CalculateBaseProdCost()
            'AddToTestString("BaseProdCost: " & blBaseProdCost.ToString)
            '   i) Production Time
            Dim blBaseProdTime As Int64 = CalculateBaseProdTime()
            'AddToTestString("BaseProdTime: " & blBaseProdTime.ToString)
            '   j) Material Costs

            'MyBase.mfrmBuilderCost.SetMinValsFromCalcs(lBaseHull, 0, blBaseResCost, blBaseProdCost, lBaseColonists, lBaseEnlisted, lBaseOfficers)
            lSumAllDA = 0
            Dim lBaseMin1 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral1ID, 0), GetSumClientDA(lMineral1ID, 0))
            'AddToTestString("BaseMin1: " & lBaseMin1.ToString)
            Dim lBaseMin2 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral2ID, 1), GetSumClientDA(lMineral2ID, 1))
            'AddToTestString("BaseMin2: " & lBaseMin2.ToString)
            Dim lBaseMin3 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral3ID, 2), GetSumClientDA(lMineral3ID, 2))
            'AddToTestString("BaseMin3: " & lBaseMin3.ToString)
            Dim lBaseMin4 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral4ID, 3), GetSumClientDA(lMineral4ID, 3))
            'AddToTestString("BaseMin4: " & lBaseMin4.ToString)
            Dim lBaseMin5 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral5ID, 4), GetSumClientDA(lMineral5ID, 4))
            'AddToTestString("BaseMin5: " & lBaseMin5.ToString)
            Dim lBaseMin6 As Int32 = CalculateBaseMineralCosts(GetMaxDA(lMineral6ID, 5), GetSumClientDA(lMineral6ID, 5))
            'AddToTestString("BaseMin6: " & lBaseMin6.ToString)

            'tpStructBody.MinValue = lBaseStructBody \ 2
            'tpStructFrame.MinValue = lBaseStructFrame \ 2
            'tpStructMeld.MinValue = lBaseStructMeld \ 2
            'tpDriveBody.MinValue = lBaseDriveBody \ 2
            'tpDriveFrame.MinValue = lBaseDriveFrame \ 2
            'tpDriveMeld.MinValue = lBaseDriveMeld \ 2

            ' 3) If the value is locked for the cost, determine the difference between the locked value and the calculated value
            Dim bMin1Locked, bMin2Locked, bMin3Locked, bMin4Locked, bMin5Locked, bMin6Locked As Boolean
            bMin1Locked = lLockedMin1 <> -1
            bMin2Locked = lLockedMin2 <> -1
            bMin3Locked = lLockedMin3 <> -1
            bMin4Locked = lLockedMin4 <> -1
            bMin5Locked = lLockedMin5 <> -1
            bMin6Locked = lLockedMin6 <> -1

            If lMineral1ID < 1 Then
                lLockedMin1 = 0 : lBaseMin1 = 0
            End If
            If lMineral2ID < 1 Then
                lLockedMin2 = 0 : lBaseMin2 = 0
            End If
            If lMineral3ID < 1 Then
                lLockedMin3 = 0 : lBaseMin3 = 0
            End If
            If lMineral4ID < 1 Then
                lLockedMin4 = 0 : lBaseMin4 = 0
            End If
            If lMineral5ID < 1 Then
                lLockedMin5 = 0 : lBaseMin5 = 0
            End If
            If lMineral6ID < 1 Then
                lLockedMin6 = 0 : lBaseMin6 = 0
            End If

            Dim bHullLocked As Boolean = lLockedHull <> -1
            Dim bPowerLocked As Boolean = lLockedPower <> -1
            Dim bResCostLocked As Boolean = blLockedResCost <> -1
            Dim bResTimeLocked As Boolean = blLockedResTime <> -1
            Dim bProdCostLocked As Boolean = blLockedProdCost <> -1
            Dim bProdTimeLocked As Boolean = blLockedProdTime <> -1
            Dim bColonistLocked As Boolean = lLockedColonists <> -1
            Dim bEnlistedLocked As Boolean = lLockedEnlisted <> -1
            Dim bOfficersLocked As Boolean = lLockedOfficers <> -1

            Dim blOrigLockedResTime As Int64 = blLockedResTime
            Dim blOrigLockedProdTime As Int64 = blLockedProdTime

            'AddToTestString("ShiftedResTime: " & blBaseResTime.ToString)
            'AddToTestString("ShiftedProdTime: " & blBaseProdTime.ToString)

            lResultHull = 0 'lBaseHull
            If bHullLocked = True Then lResultHull = lLockedHull Else lResultHull = lBaseHull
            lResultPower = 0
            If bPowerLocked = True Then lResultPower = lLockedPower Else lResultPower = lBasePower
            blResultResCost = 0
            If bResCostLocked = True Then blResultResCost = blLockedResCost Else blResultResCost = blBaseResCost
            blResultResTime = 0
            If bResTimeLocked = True Then blResultResTime = blLockedResTime Else blResultResTime = blBaseResTime
            blResultProdCost = 0
            If bProdCostLocked = True Then blResultProdCost = blLockedProdCost Else blResultProdCost = blBaseProdCost
            blResultProdTime = 0
            If bProdTimeLocked = True Then blResultProdTime = blLockedProdTime Else blResultProdTime = blBaseProdTime
            lResultColonists = 0
            If bColonistLocked = True Then lResultColonists = lLockedColonists Else lResultColonists = lBaseColonists
            lResultEnlisted = 0
            If bEnlistedLocked = True Then lResultEnlisted = lLockedEnlisted Else lResultEnlisted = lBaseEnlisted
            lResultOfficers = 0
            If bOfficersLocked = True Then lResultOfficers = lLockedOfficers Else lResultOfficers = lBaseOfficers
            lResultMin1 = 0
            If bMin1Locked = True Then lResultMin1 = lLockedMin1 Else lResultMin1 = lBaseMin1
            lResultMin2 = 0
            If bMin2Locked = True Then lResultMin2 = lLockedMin2 Else lResultMin2 = lBaseMin2
            lResultMin3 = 0
            If bMin3Locked = True Then lResultMin3 = lLockedMin3 Else lResultMin3 = lBaseMin3
            lResultMin4 = 0
            If bMin4Locked = True Then lResultMin4 = lLockedMin4 Else lResultMin4 = lBaseMin4
            lResultMin5 = 0
            If bMin5Locked = True Then lResultMin5 = lLockedMin5 Else lResultMin5 = lBaseMin5
            lResultMin6 = 0
            If bMin6Locked = True Then lResultMin6 = lLockedMin6 Else lResultMin6 = lBaseMin6

            If bHullLocked = True AndAlso InvalidHalfBaseCheck(lBaseHull, lLockedHull, 2) = True Then Return False
            If bPowerLocked = True AndAlso InvalidHalfBaseCheck(lBasePower, lLockedPower, 2) = True Then Return False
            If bResCostLocked = True AndAlso InvalidHalfBaseCheck(blBaseResCost, blLockedResCost, 0) = True Then Return False
            If bProdCostLocked = True AndAlso InvalidHalfBaseCheck(blBaseProdCost, blLockedProdCost, 1) = True Then Return False
            If bColonistLocked = True AndAlso InvalidHalfBaseCheck(lBaseColonists, lLockedColonists, 3) = True Then Return False
            If bEnlistedLocked = True AndAlso InvalidHalfBaseCheck(lBaseEnlisted, lLockedEnlisted, 3) = True Then Return False
            If bOfficersLocked = True AndAlso InvalidHalfBaseCheck(lBaseOfficers, lLockedOfficers, 3) = True Then Return False
            If bMin1Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin1, lLockedMin1, 0) = True Then Return False
            If bMin2Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin2, lLockedMin2, 0) = True Then Return False
            If bMin3Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin3, lLockedMin3, 0) = True Then Return False
            If bMin4Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin4, lLockedMin4, 0) = True Then Return False
            If bMin5Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin5, lLockedMin5, 0) = True Then Return False
            If bMin6Locked = True AndAlso InvalidHalfBaseCheck(lBaseMin6, lLockedMin6, 0) = True Then Return False
            If (bResTimeLocked = True AndAlso blLockedResTime < 1) OrElse (bProdTimeLocked = True AndAlso blLockedProdTime < 1) Then
                'lblDesignFlaw.Caption = "Your scientists cannot get the time that low."
                'lblDesignFlaw.Visible = True
                'mbIgnoreValueChange = False
                Return False
            End If

            Dim lHullDiff As Int32 = DoLockedToBaseCheck(lLockedHull, lBaseHull)
            Dim lPowerDiff As Int32 = DoLockedToBaseCheck(lLockedPower, lBasePower)
            Dim blResCostDiff As Int64 = DoLockedToBaseCheck(blLockedResCost, blBaseResCost)
            Dim blResTimeDiff As Int64 = DoLockedToBaseCheck(blLockedResTime, blBaseResTime)
            Dim blProdCostDiff As Int64 = DoLockedToBaseCheck(blLockedProdCost, blBaseProdCost)
            Dim blProdTimeDiff As Int64 = DoLockedToBaseCheck(blLockedProdTime, blBaseProdTime)
            Dim lColonistDiff As Int32 = DoLockedToBaseCheck(lLockedColonists, lBaseColonists)
            Dim lEnlistedDiff As Int32 = DoLockedToBaseCheck(lLockedEnlisted, lBaseEnlisted)
            Dim lOfficerDiff As Int32 = DoLockedToBaseCheck(lLockedOfficers, lBaseOfficers)
            Dim lMin1Diff As Int32 = DoLockedToBaseCheck(lLockedMin1, lBaseMin1)
            Dim lMin2Diff As Int32 = DoLockedToBaseCheck(lLockedMin2, lBaseMin2)
            Dim lMin3Diff As Int32 = DoLockedToBaseCheck(lLockedMin3, lBaseMin3)
            Dim lMin4Diff As Int32 = DoLockedToBaseCheck(lLockedMin4, lBaseMin4)
            Dim lMin5Diff As Int32 = DoLockedToBaseCheck(lLockedMin5, lBaseMin5)
            Dim lMin6Diff As Int32 = DoLockedToBaseCheck(lLockedMin6, lBaseMin6)

            ' 4) Determine the point costs for each item 
            Dim blPoints As Int64 = 0
            'blPoints += CLng(lHullDiff) * ml_POINT_PER_HULL
            'blPoints += CLng(lPowerDiff) * ml_POINT_PER_POWER
            'blPoints += blResTimeDiff * ml_POINT_PER_RES_TIME
            'blPoints += blResCostDiff * ml_POINT_PER_RES_COST
            'blPoints += blProdTimeDiff * ml_POINT_PER_PROD_TIME
            'blPoints += blProdCostDiff * ml_POINT_PER_PROD_COST
            'blPoints += lColonistDiff * ml_POINT_PER_COLONIST
            'blPoints += lEnlistedDiff * ml_POINT_PER_ENLISTED
            'blPoints += lOfficerDiff * ml_POINT_PER_OFFICER
            'blPoints += lMin1Diff * ml_POINT_PER_MIN1
            'blPoints += lMin2Diff * ml_POINT_PER_MIN2
            'blPoints += lMin3Diff * ml_POINT_PER_MIN3
            'blPoints += lMin4Diff * ml_POINT_PER_MIN4
            'blPoints += lMin5Diff * ml_POINT_PER_MIN5
            'blPoints += lMin6Diff * ml_POINT_PER_MIN6
            blPoints += GetPointValue(lHullDiff, "Hull", ml_POINT_PER_HULL)
            blPoints += GetPointValue(lPowerDiff, "Power", ml_POINT_PER_POWER)
            blPoints += GetPointValue(blResTimeDiff, "ResTime", ml_POINT_PER_RES_TIME)
            blPoints += GetPointValue(blResCostDiff, "ResCost", ml_POINT_PER_RES_COST)
            blPoints += GetPointValue(blProdTimeDiff, "ProdTime", ml_POINT_PER_PROD_TIME)
            blPoints += GetPointValue(blProdCostDiff, "ProdCost", ml_POINT_PER_PROD_COST)
            blPoints += GetPointValue(lColonistDiff, "Colonist", ml_POINT_PER_COLONIST)
            blPoints += GetPointValue(lEnlistedDiff, "Enlisted", ml_POINT_PER_ENLISTED)
            blPoints += GetPointValue(lOfficerDiff, "Officer", ml_POINT_PER_OFFICER)
            blPoints += GetPointValue(lMin1Diff, "Min1", ml_POINT_PER_MIN1)
            blPoints += GetPointValue(lMin2Diff, "Min2", ml_POINT_PER_MIN2)
            blPoints += GetPointValue(lMin3Diff, "Min3", ml_POINT_PER_MIN3)
            blPoints += GetPointValue(lMin4Diff, "Min4", ml_POINT_PER_MIN4)
            blPoints += GetPointValue(lMin5Diff, "Min5", ml_POINT_PER_MIN5)
            blPoints += GetPointValue(lMin6Diff, "Min6", ml_POINT_PER_MIN6)

            'AddToTestString("BasePoints: " & blPoints.ToString)
            If blPoints < 0 Then Return False

            'Now, check for our additive points from being uber
            Dim decBA As Decimal = CDec(lMaxGuns) / decNormalize '/ CDec(lMaxGuns)
            Dim decThreshold As Decimal = CalculateThresholdValue()
            'IF Power*Thrust*Manuever*Speed*B/A > 1,350,000,000 THEN
            'AddToTestString("Threshold: " & decThreshold.ToString)
            If iTechTypeID = ObjectType.eShieldTech Then
                If CDec(blPoints + decThreshold) > Int64.MaxValue Then blPoints = Int64.MaxValue Else blPoints = CLng(blPoints + decThreshold)
            Else
                If decThreshold > decMaxThreshold Then
                    Dim blAllConstants As Int64 = ml_POINT_PER_HULL
                    blAllConstants += ml_POINT_PER_POWER
                    blAllConstants += ml_POINT_PER_RES_TIME
                    blAllConstants += ml_POINT_PER_RES_COST
                    blAllConstants += ml_POINT_PER_PROD_TIME
                    blAllConstants += ml_POINT_PER_PROD_COST
                    blAllConstants += ml_POINT_PER_COLONIST
                    blAllConstants += ml_POINT_PER_ENLISTED
                    blAllConstants += ml_POINT_PER_OFFICER
                    blAllConstants += ml_POINT_PER_MIN1
                    blAllConstants += ml_POINT_PER_MIN2
                    blAllConstants += ml_POINT_PER_MIN3
                    blAllConstants += ml_POINT_PER_MIN4
                    blAllConstants += ml_POINT_PER_MIN5
                    blAllConstants += ml_POINT_PER_MIN6

                    Dim decAdd As Decimal = 0
                    If iTechTypeID = ObjectType.eArmorTech Then
                        decAdd = CType(Me, ArmorTechComputer).GetThresholdModifier()
                    Else
                        decAdd = decThreshold - decMaxThreshold
                        If iTechTypeID = ObjectType.eWeaponTech Then decAdd *= 1000000D
                    End If
                    decAdd *= CDec(blAllConstants)
                    If CDec(blPoints + decAdd) > Int64.MaxValue Then blPoints = Int64.MaxValue Else blPoints = CLng(blPoints + decAdd)
                End If
            End If


            'AddToTestString("Points: " & blPoints.ToString)

            ' 5) Take the sum of the remaining points and divide it by 1 point of each type not locked summed together rounded down
            Dim blPointSum As Int64 = 0
            Dim lModByVal As Int32 = 0
            If blPoints < 0 Then

                Dim lAdjusters As Int32 = 0

                If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 AndAlso lResultHull > 0 Then
                    blPointSum += ml_POINT_PER_HULL
                    lAdjusters += 1
                End If
                If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 AndAlso lResultPower > 0 Then
                    blPointSum += ml_POINT_PER_POWER
                    lAdjusters += 1
                End If
                If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 AndAlso blResultResCost > 0 Then
                    blPointSum += ml_POINT_PER_RES_COST
                    lAdjusters += 1
                End If
                If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 AndAlso blResultResTime > 0 Then
                    blPointSum += ml_POINT_PER_RES_TIME
                    lAdjusters += 1
                End If
                If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 AndAlso blResultProdCost > 0 Then
                    blPointSum += ml_POINT_PER_PROD_COST
                    lAdjusters += 1
                End If
                If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 AndAlso blResultProdTime > 0 Then
                    blPointSum += ml_POINT_PER_PROD_TIME
                    lAdjusters += 1
                End If
                If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 AndAlso lResultColonists > 0 Then
                    blPointSum += ml_POINT_PER_COLONIST
                    lAdjusters += 1
                End If
                If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 AndAlso lResultEnlisted > 0 Then
                    blPointSum += ml_POINT_PER_ENLISTED
                    lAdjusters += 1
                End If
                If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 AndAlso lResultOfficers > 0 Then
                    blPointSum += ml_POINT_PER_OFFICER
                    lAdjusters += 1
                End If

                If bMin1Locked = False AndAlso lResultMin1 > 0 Then
                    blPointSum += ml_POINT_PER_MIN1
                    lAdjusters += 1
                End If
                If bMin2Locked = False AndAlso lResultMin2 > 0 Then
                    blPointSum += ml_POINT_PER_MIN2
                    lAdjusters += 1
                End If
                If bMin3Locked = False AndAlso lResultMin3 > 0 Then
                    blPointSum += ml_POINT_PER_MIN3
                    lAdjusters += 1
                End If
                If bMin4Locked = False AndAlso lResultMin4 > 0 Then
                    blPointSum += ml_POINT_PER_MIN4
                    lAdjusters += 1
                End If
                If bMin5Locked = False AndAlso lResultMin5 > 0 Then
                    blPointSum += ml_POINT_PER_MIN5
                    lAdjusters += 1
                End If
                If bMin6Locked = False AndAlso lResultMin6 > 0 Then
                    blPointSum += ml_POINT_PER_MIN6
                    lAdjusters += 1
                End If

                If blPointSum > 0 Then

                    'AddToTestString("PointSum: " & blPointSum.ToString)

                    ' 6) The result is applied to all non-locked values
                    'Dim blTemp As Int64 = Math.Abs(blPoints) \ blPointSum
                    Dim blTemp As Int64 = CLng(Math.Log10(blPoints) * blPoints) \ blPointSum 'CLng(Math.Pow(blPoints, 1 / 0.9)) \ blPointSum
                    Dim blBaseAdjust As Int64 = blPoints \ blPointSum

                    Dim lAddPts As Int32 = 0
                    If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                    If blTemp < 0 Then blTemp = 0
                    lAddPts = CInt(blTemp)
                    If lAddPts > 0 Then
                        If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 AndAlso lResultHull > lAddPts Then
                            lResultHull -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 AndAlso lResultPower > lAddPts Then
                            lAdjusters -= 1
                            lResultPower -= lAddPts
                        End If
                        If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 AndAlso blResultResCost > lAddPts Then
                            blResultResCost -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 AndAlso blResultResTime > lAddPts Then
                            blResultResTime -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 AndAlso blResultProdCost > lAddPts Then
                            blResultProdCost -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 AndAlso blResultProdTime > lAddPts Then
                            blResultProdTime -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 AndAlso lResultColonists > lAddPts Then
                            lResultColonists -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 AndAlso lResultEnlisted > lAddPts Then
                            lResultEnlisted -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 AndAlso lResultOfficers > lAddPts Then
                            lResultOfficers -= lAddPts
                            lAdjusters -= 1
                        End If

                        If bMin1Locked = False AndAlso lResultMin1 > lAddPts Then
                            lResultMin1 -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bMin2Locked = False AndAlso lResultMin2 > lAddPts Then
                            lResultMin2 -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bMin3Locked = False AndAlso lResultMin3 > lAddPts Then
                            lResultMin3 -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bMin4Locked = False AndAlso lResultMin4 > lAddPts Then
                            lResultMin4 -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bMin5Locked = False AndAlso lResultMin5 > lAddPts Then
                            lResultMin5 -= lAddPts
                            lAdjusters -= 1
                        End If
                        If bMin6Locked = False AndAlso lResultMin6 > lAddPts Then
                            lResultMin6 -= lAddPts
                            lAdjusters -= 1
                        End If

                        If lAdjusters <> 0 Then Return False

                        blPoints += blBaseAdjust * blPointSum
                    End If
                End If

                lModByVal = -1
            ElseIf blPoints > 0 Then
                Dim lAdjusters As Int32 = 0

                If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 Then
                    blPointSum += ml_POINT_PER_HULL
                    lAdjusters += 1
                End If
                If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 Then
                    blPointSum += ml_POINT_PER_POWER
                    lAdjusters += 1
                End If
                If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 Then
                    blPointSum += ml_POINT_PER_RES_COST
                    lAdjusters += 1
                End If
                If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 Then
                    blPointSum += ml_POINT_PER_RES_TIME
                    lAdjusters += 1
                End If
                If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 Then
                    blPointSum += ml_POINT_PER_PROD_COST
                    lAdjusters += 1
                End If
                If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 Then
                    blPointSum += ml_POINT_PER_PROD_TIME
                    lAdjusters += 1
                End If
                If ml_POINT_PER_MIN1 > 0 AndAlso bMin1Locked = False Then
                    blPointSum += ml_POINT_PER_MIN1
                    lAdjusters += 1
                End If
                If ml_POINT_PER_MIN2 > 0 AndAlso bMin2Locked = False Then
                    blPointSum += ml_POINT_PER_MIN2
                    lAdjusters += 1
                End If
                If ml_POINT_PER_MIN3 > 0 AndAlso bMin3Locked = False Then
                    blPointSum += ml_POINT_PER_MIN3
                    lAdjusters += 1
                End If
                If ml_POINT_PER_MIN4 > 0 AndAlso bMin4Locked = False Then
                    blPointSum += ml_POINT_PER_MIN4
                    lAdjusters += 1
                End If
                If ml_POINT_PER_MIN5 > 0 AndAlso bMin5Locked = False Then
                    blPointSum += ml_POINT_PER_MIN5
                    lAdjusters += 1
                End If
                If ml_POINT_PER_MIN6 > 0 AndAlso bMin6Locked = False Then
                    blPointSum += ml_POINT_PER_MIN6
                    lAdjusters += 1
                End If

                If blPointSum > 0 Then
                    ' AddToTestString("PointSum: " & blPointSum.ToString)

                    ' 6) The result is applied to all non-locked values
                    'Dim blTemp As Int64 = blPoints \ blPointSum
                    Dim blTemp As Int64 = CLng(Math.Log10(blPoints) * blPoints) \ blPointSum 'CLng(Math.Pow(blPoints, 1 / 0.9)) \ blPointSum
                    Dim blBaseAdjust As Int64 = blPoints \ blPointSum

                    Dim lAddPts As Int32 = 0
                    'If blTemp > Int32.MaxValue Then lAddPts = Int32.MaxValue
                    'If blTemp < 0 Then lAddPts = 0
                    If blTemp > 0 Then
                        If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 Then
                            Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_HULL
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultHull += lAddPts 'lBaseHull += lAddPts
                            lAdjusters -= 1
                        End If
                        If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 Then
                            Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_HULL
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultPower += lAddPts 'lBaseHull += lAddPts
                            lAdjusters -= 1
                        End If
                        If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 Then
                            blResultResCost += blTemp ' \ ml_POINT_PER_RES_COST 'blBaseResCost += blTemp \ ml_POINT_PER_RES_COST
                            lAdjusters -= 1
                        End If
                        If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 Then
                            blResultResTime += blTemp '\ ml_POINT_PER_RES_TIME
                            lAdjusters -= 1
                        End If
                        If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 Then
                            blResultProdCost += blTemp '\ ml_POINT_PER_PROD_COST
                            lAdjusters -= 1
                        End If
                        If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 Then
                            blResultProdTime += blTemp '\ ml_POINT_PER_PROD_TIME
                            lAdjusters -= 1
                        End If

                        If ml_POINT_PER_MIN1 > 0 AndAlso bMin1Locked = False AndAlso ml_POINT_PER_MIN1 > 0 Then
                            Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_DBody
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultMin1 += lAddPts
                            lAdjusters -= 1
                        End If
                        If ml_POINT_PER_MIN2 > 0 AndAlso bMin2Locked = False AndAlso ml_POINT_PER_MIN2 > 0 Then
                            Dim blAdd As Int64 = blTemp ' \ ml_POINT_PER_DFrame
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultMin2 += lAddPts
                            lAdjusters -= 1
                        End If
                        If ml_POINT_PER_MIN3 > 0 AndAlso bMin3Locked = False AndAlso ml_POINT_PER_MIN3 > 0 Then
                            Dim blAdd As Int64 = blTemp ' \ ml_POINT_PER_DMeld
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultMin3 += lAddPts
                            lAdjusters -= 1
                        End If
                        If ml_POINT_PER_MIN4 > 0 AndAlso bMin4Locked = False AndAlso ml_POINT_PER_MIN4 > 0 Then
                            Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_SBody
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultMin4 += lAddPts
                            lAdjusters -= 1
                        End If
                        If ml_POINT_PER_MIN5 > 0 AndAlso bMin5Locked = False AndAlso ml_POINT_PER_MIN5 > 0 Then
                            Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_SFrame
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultMin5 += lAddPts
                            lAdjusters -= 1
                        End If
                        If ml_POINT_PER_MIN6 > 0 AndAlso bMin6Locked = False AndAlso ml_POINT_PER_MIN6 > 0 Then
                            Dim blAdd As Int64 = blTemp ' \ ml_POINT_PER_SMeld
                            If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
                            lResultMin6 += lAddPts
                            lAdjusters -= 1
                        End If

                        If lAdjusters <> 0 Then Return False

                        blPoints -= blBaseAdjust * blPointSum
                    End If
                End If

                lModByVal = 1
            End If

            If blPoints > 0 Then
                blPointSum = 0
                If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 Then blPointSum += ml_POINT_PER_COLONIST
                If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 Then blPointSum += ml_POINT_PER_ENLISTED
                If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 Then blPointSum += ml_POINT_PER_OFFICER

                If blPointSum > 0 Then
                    Dim blTemp As Int64 = blPoints \ blPointSum
                    Dim lAddPts As Int32 = 0
                    If blTemp > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blTemp)
                    If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 Then lResultColonists += lAddPts
                    If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 Then lResultEnlisted += lAddPts
                    If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 Then lResultOfficers += lAddPts

                    blPoints -= CLng(blTemp) * blPointSum
                End If
            End If

            ' 7) The remainder points are applied to the non-locked items until the remainder < 0
            Dim lCnt As Int32 = 0
            While blPoints > 0
                Dim blPrevPoints As Int64 = blPoints
                If bHullLocked = False AndAlso ml_POINT_PER_HULL > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_HULL, lResultHull, lModByVal)
                If bPowerLocked = False AndAlso ml_POINT_PER_POWER > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_POWER, lResultPower, lModByVal)
                If bResCostLocked = False AndAlso ml_POINT_PER_RES_COST > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_RES_COST, blResultResCost, lModByVal)
                If bResTimeLocked = False AndAlso ml_POINT_PER_RES_TIME > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_RES_TIME, blResultResTime, lModByVal)
                If bProdCostLocked = False AndAlso ml_POINT_PER_PROD_COST > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_PROD_COST, blResultProdCost, lModByVal)
                If bProdTimeLocked = False AndAlso ml_POINT_PER_PROD_TIME > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_PROD_TIME, blResultProdTime, lModByVal)
                If bColonistLocked = False AndAlso ml_POINT_PER_COLONIST > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_COLONIST, lResultColonists, lModByVal)
                If bEnlistedLocked = False AndAlso ml_POINT_PER_ENLISTED > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_ENLISTED, lResultEnlisted, lModByVal)
                If bOfficersLocked = False AndAlso ml_POINT_PER_OFFICER > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_OFFICER, lResultOfficers, lModByVal)
                If bMin1Locked = False AndAlso ml_POINT_PER_MIN1 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN1, lResultMin1, lModByVal)
                If bMin2Locked = False AndAlso ml_POINT_PER_MIN2 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN2, lResultMin2, lModByVal)
                If bMin3Locked = False AndAlso ml_POINT_PER_MIN3 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN3, lResultMin3, lModByVal)
                If bMin4Locked = False AndAlso ml_POINT_PER_MIN4 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN4, lResultMin4, lModByVal)
                If bMin5Locked = False AndAlso ml_POINT_PER_MIN5 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN5, lResultMin5, lModByVal)
                If bMin6Locked = False AndAlso ml_POINT_PER_MIN6 > 0 Then AffectValueByPoints(blPoints, ml_POINT_PER_MIN6, lResultMin6, lModByVal)
                ' 8) If the remainder does not reach 0 after several attempts, break out
                If blPrevPoints = blPoints Then Exit While
                lCnt += 1
                If lCnt > 100 Then Exit While
            End While

            'frmBuildCost.SetBaseValues(lResultHull, lResultPower, blResultResCost, blResultResTime, blResultProdCost, blResultProdTime, lResultColonists, lResultEnlisted, lResultOfficers)


            'AddToTestString("RESULTS: ")
            'AddToTestString(" Remaining Points: " & blPoints.ToString)
            'AddToTestString(" Hull: " & lResultHull.ToString)
            'AddToTestString(" Power: " & lResultPower.ToString)
            'AddToTestString(" ResCost: " & blResultResCost.ToString)
            'AddToTestString(" ResTime: " & blResultResTime.ToString)
            'AddToTestString(" ProdCost: " & blResultProdCost.ToString)
            'AddToTestString(" ProdTime: " & blResultProdTime.ToString)
            'AddToTestString(" Colonist: " & lResultColonists.ToString)
            'AddToTestString(" Enlisted: " & lResultEnlisted.ToString)
            'AddToTestString(" Officer: " & lResultOfficers.ToString)
            'AddToTestString(" Min1: " & lResultMin1.ToString)
            'AddToTestString(" Min2: " & lResultMin2.ToString)
            'AddToTestString(" Min3: " & lResultMin3.ToString)
            'AddToTestString(" Min4: " & lResultMin4.ToString)
            'AddToTestString(" Min5: " & lResultMin5.ToString)
            'AddToTestString(" Min6: " & lResultMin6.ToString)
            Dim decIntelMult As Decimal = 160D / lPopIntel
            blResultResCost = CLng(blResultResCost * decIntelMult)
            blResultResTime = CLng(blResultResTime * decIntelMult)

            If Math.Abs(blPoints) > 1000 Then
                'lblDesignFlaw.Visible = True
                'lblDesignFlaw.Caption = "Your scientists believe this design to be possible."
                Return False
            End If

            Return True
        Catch
            'lblDesignFlaw.Visible = True
            'lblDesignFlaw.Caption = "This design is impossible to comprehend."
            Return False
        End Try

    End Function

    Public lSumAllDA As Int32
    Private Function GetMaxDA(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32

        Dim oMineral As Mineral = GetEpicaMineral(lMineralID)
        If oMineral Is Nothing Then Return 0

        Dim lTempMax As Int32 = 0
        If muPropReqs Is Nothing = False Then
            For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
                If muPropReqs(X).lMinID = lMineralIndex Then
                    Dim lA As Int32 = oMineral.GetPropertyValue(muPropReqs(X).lPropID) 'oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))

                    lTempMax = Math.Max(Math.Abs(muPropReqs(X).lPropVal - lA), lTempMax)

                    'Now, for our D-A Sum
                    Dim lClientD As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(muPropReqs(X).lPropVal)
                    Dim lClientA As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lA)
                    lSumAllDA += Math.Abs(lClientD - lClientA)

                End If
            Next X
        End If
        Dim lMaxDA As Int32 = lTempMax

        ''TODO: Remove me when ready!!!
        'If oMineral.ObjectID = 1 Then
        '    lMaxDA = 5 * 23
        'ElseIf oMineral.ObjectID = 5 Then
        '    lMaxDA = 23
        'ElseIf oMineral.ObjectID = 9 Then
        '    lMaxDA = 255
        'End If

        Return lMaxDA
    End Function

    Private Function GetSumClientDA(ByVal lMineralID As Int32, ByVal lMineralIndex As Int32) As Int32

        Dim oMineral As Mineral = GetEpicaMineral(lMineralID)
        If oMineral Is Nothing Then Return 0

        Dim lResult As Int32 = 0

        If muPropReqs Is Nothing = False Then
            For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
                If muPropReqs(X).lMinID = lMineralIndex Then
                    Dim lA As Int32 = oMineral.GetPropertyValue(muPropReqs(X).lPropID) 'oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))

                    'Now, for our D-A Sum
                    Dim lClientD As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(muPropReqs(X).lPropVal)
                    Dim lClientA As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(lA)
                    lResult += Math.Abs(lClientD - lClientA)

                End If
            Next X
        End If
        If lResult < 1 Then lResult = 1

        Return lResult
    End Function

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
        Return 0
    End Function

    Protected Overrides Function GetTheBill() As Decimal
        Return 0
    End Function
End Class
