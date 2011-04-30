Option Strict On

Public Class EngineTechComputer
    Inherits TechBuilderComputer

    Public blPowGen As Int64 = 0
    Public blThrustGen As Int64 = 0
    Public blMaxSpd As Int64 = 0
    Public blMan As Int64 = 0

    'Protected Overrides Function CalculateBaseHull() As Int32
    '    Dim decTemp As Decimal = CDec((blPowGen * 0.2F) + (blThrustGen * 0.18F))
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < Int32.MinValue Then Return Int32.MinValue
    '    Return CInt(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseColonists() As Int32
    '    Return CInt(blPowGen * 0.1F)
    'End Function
    'Protected Overrides Function CalculateBaseEnlisted() As Int32
    '    Return CInt((blPowGen * 0.01F) + (blThrustGen * 0.01F))
    'End Function
    'Protected Overrides Function CalculateBaseOfficers() As Int32
    '    Return CInt((blPowGen * 0.001F) + (blThrustGen * 0.001F))
    'End Function
    'Protected Overrides Function CalculateBaseResCost() As Int64
    '    Dim decTemp As Decimal = (CDec(blPowGen) * 1000D) + (CDec(blThrustGen) * 1000D)
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseResTime() As Int64
    '    Dim decTemp As Decimal = CDec(blPowGen) * 10000D
    '    decTemp += CDec(blThrustGen) * 32000D
    '    decTemp += CDec(blMaxSpd) * 2700000D
    '    decTemp += CDec(blMan) * 5D * CDec(blThrustGen)
    '    'decTemp /= 180000D
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdCost() As Int64
    '    Dim decTemp As Decimal = CDec(blPowGen) * 750D
    '    decTemp += CDec(blThrustGen) * 850D
    '    Return CLng(decTemp)
    'End Function
    'Protected Overrides Function CalculateBaseProdTime() As Int64
    '    Dim decTemp As Decimal = CDec(blPowGen) * 2000000D
    '    decTemp += CDec(blThrustGen) * 1000000D * CDec(blMan)
    '    ' decTemp /= 800000D
    '    Return CLng(decTemp)
    'End Function

    ''Private Function CalculateBaseStructBody() As Int32
    ''    Dim oMineral As Mineral = Nothing
    ''    If cboStructBody.ListIndex > -1 Then
    ''        Dim lID As Int32 = cboStructBody.ItemData(cboStructBody.ListIndex)
    ''        For X As Int32 = 0 To glMineralUB
    ''            If glMineralIdx(X) = lID Then
    ''                oMineral = goMinerals(X)
    ''                Exit For
    ''            End If
    ''        Next X
    ''    End If
    ''    If oMineral Is Nothing Then Return 0

    ''    Dim lTempMax As Int32 = 0
    ''    If muPropReqs Is Nothing = False Then
    ''        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    ''            If muPropReqs(X).lMinID = 0 Then
    ''                Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
    ''                lTempMax = Math.Max(GetDesiredActualRelationshipValue(muPropReqs(X).lPropVal, lA), lTempMax)
    ''            End If
    ''        Next X
    ''    End If
    ''    Dim lMaxDA As Int32 = lTempMax
    ''    'Dim lProp1Desired As Int32 = 0      'magprod
    ''    'Dim lProp2Desired As Int32 = 6      'density
    ''    'Dim lProp3Desired As Int32 = 6      'melting pt

    ''    'Dim lTemp As Int32 = oMineral.GetPropertyIndex(15)
    ''    'Dim lProp1Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(1)
    ''    'Dim lProp2Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(8)
    ''    'Dim lProp3Actual As Int32 = oMineral.MinPropValueScore(lTemp)

    ''    'lProp1Actual = Math.Abs(lProp1Desired - lProp1Actual)
    ''    'lProp2Actual = Math.Abs(lProp2Desired - lProp2Actual)
    ''    'lProp3Actual = Math.Abs(lProp3Desired - lProp3Actual)
    ''    'Dim lMaxDA As Int32 = Math.Max(Math.Max(lProp2Actual, lProp3Actual), lProp1Actual)

    ''    'TODO: Remove me when ready!!!
    ''    If oMineral.ObjectID = 1 Then
    ''        lMaxDA = 5 * 23
    ''    ElseIf oMineral.ObjectID = 5 Then
    ''        lMaxDA = 23
    ''    ElseIf oMineral.ObjectID = 9 Then
    ''        lMaxDA = 255
    ''    End If
    ''    Dim decTemp As Decimal = (1D + lMaxDA) * ((blPowGen + blThrustGen) / 2D + (CDec(blMaxSpd) + CDec(blMan) + 1D)) / 50D
    ''    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    ''    If decTemp < 0 Then Return 0
    ''    Return CInt(decTemp)
    ''End Function
    ''Private Function CalculateBaseStructFrame() As Int32
    ''    Dim oMineral As Mineral = Nothing
    ''    If cboStructFrame.ListIndex > -1 Then
    ''        Dim lID As Int32 = cboStructFrame.ItemData(cboStructFrame.ListIndex)
    ''        For X As Int32 = 0 To glMineralUB
    ''            If glMineralIdx(X) = lID Then
    ''                oMineral = goMinerals(X)
    ''                Exit For
    ''            End If
    ''        Next X
    ''    End If
    ''    If oMineral Is Nothing Then Return 0

    ''    Dim lTempMax As Int32 = 0
    ''    If muPropReqs Is Nothing = False Then
    ''        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    ''            If muPropReqs(X).lMinID = 1 Then
    ''                Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
    ''                lTempMax = Math.Max(GetDesiredActualRelationshipValue(muPropReqs(X).lPropVal, lA), lTempMax)
    ''            End If
    ''        Next X
    ''    End If
    ''    Dim lMaxDA As Int32 = lTempMax
    ''    'Dim lProp1Desired As Int32 = 6      'hardness
    ''    'Dim lProp2Desired As Int32 = 5      'malleable
    ''    'Dim lProp3Desired As Int32 = 0      'compress

    ''    'Dim lTemp As Int32 = oMineral.GetPropertyIndex(2)
    ''    'Dim lProp1Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(3)
    ''    'Dim lProp2Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(5)
    ''    'Dim lProp3Actual As Int32 = oMineral.MinPropValueScore(lTemp)

    ''    'lProp1Actual = Math.Abs(lProp1Desired - lProp1Actual)
    ''    'lProp2Actual = Math.Abs(lProp2Desired - lProp2Actual)
    ''    'lProp3Actual = Math.Abs(lProp3Desired - lProp3Actual)
    ''    'Dim lMaxDA As Int32 = Math.Max(Math.Max(lProp2Actual, lProp3Actual), lProp1Actual)

    ''    'TODO: Remove me when ready!!!
    ''    If oMineral.ObjectID = 1 Then
    ''        lMaxDA = 5 * 23
    ''    ElseIf oMineral.ObjectID = 5 Then
    ''        lMaxDA = 23
    ''    ElseIf oMineral.ObjectID = 9 Then
    ''        lMaxDA = 255
    ''    End If
    ''    Dim decTemp As Decimal = (1D + lMaxDA) * ((blPowGen + blThrustGen) / 2D + (CDec(blMaxSpd) + CDec(blMan) + 1D)) / 50D
    ''    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    ''    If decTemp < 0 Then Return 0
    ''    Return CInt(decTemp)
    ''End Function
    ''Private Function CalculateBaseStructMeld() As Int32
    ''    Dim oMineral As Mineral = Nothing
    ''    If cboStructMeld.ListIndex > -1 Then
    ''        Dim lID As Int32 = cboStructMeld.ItemData(cboStructMeld.ListIndex)
    ''        For X As Int32 = 0 To glMineralUB
    ''            If glMineralIdx(X) = lID Then
    ''                oMineral = goMinerals(X)
    ''                Exit For
    ''            End If
    ''        Next X
    ''    End If
    ''    If oMineral Is Nothing Then Return 0

    ''    Dim lTempMax As Int32 = 0
    ''    If muPropReqs Is Nothing = False Then
    ''        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    ''            If muPropReqs(X).lMinID = 2 Then
    ''                Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
    ''                lTempMax = Math.Max(GetDesiredActualRelationshipValue(muPropReqs(X).lPropVal, lA), lTempMax)
    ''            End If
    ''        Next X
    ''    End If
    ''    Dim lMaxDA As Int32 = lTempMax

    ''    'Dim lProp1Desired As Int32 = 6      'hardness
    ''    'Dim lProp2Desired As Int32 = 0      'malleable
    ''    'Dim lProp3Desired As Int32 = 6      'melting pt

    ''    'Dim lTemp As Int32 = oMineral.GetPropertyIndex(2)
    ''    'Dim lProp1Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(3)
    ''    'Dim lProp2Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(8)
    ''    'Dim lProp3Actual As Int32 = oMineral.MinPropValueScore(lTemp)

    ''    'lProp1Actual = Math.Abs(lProp1Desired - lProp1Actual)
    ''    'lProp2Actual = Math.Abs(lProp2Desired - lProp2Actual)
    ''    'lProp3Actual = Math.Abs(lProp3Desired - lProp3Actual)
    ''    'Dim lMaxDA As Int32 = Math.Max(Math.Max(lProp2Actual, lProp3Actual), lProp1Actual)

    ''    'TODO: Remove me when ready!!!
    ''    If oMineral.ObjectID = 1 Then
    ''        lMaxDA = 5 * 23
    ''    ElseIf oMineral.ObjectID = 5 Then
    ''        lMaxDA = 23
    ''    ElseIf oMineral.ObjectID = 9 Then
    ''        lMaxDA = 255
    ''    End If
    ''    Dim decTemp As Decimal = (1D + lMaxDA) * ((blPowGen + blThrustGen) / 2D + (CDec(blMaxSpd) + CDec(blMan) + 1D)) / 50D
    ''    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    ''    If decTemp < 0 Then Return 0
    ''    Return CInt(decTemp)
    ''End Function
    ''Private Function CalculateBaseDriveBody() As Int32
    ''    Dim oMineral As Mineral = Nothing
    ''    If cboDriveBody.ListIndex > -1 Then
    ''        Dim lID As Int32 = cboDriveBody.ItemData(cboDriveBody.ListIndex)
    ''        For X As Int32 = 0 To glMineralUB
    ''            If glMineralIdx(X) = lID Then
    ''                oMineral = goMinerals(X)
    ''                Exit For
    ''            End If
    ''        Next X
    ''    End If
    ''    If oMineral Is Nothing Then Return 0

    ''    Dim lTempMax As Int32 = 0
    ''    If muPropReqs Is Nothing = False Then
    ''        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    ''            If muPropReqs(X).lMinID = 3 Then
    ''                Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
    ''                lTempMax = Math.Max(GetDesiredActualRelationshipValue(muPropReqs(X).lPropVal, lA), lTempMax)
    ''            End If
    ''        Next X
    ''    End If
    ''    Dim lMaxDA As Int32 = lTempMax
    ''    'Dim lProp1Desired As Int32 = 0      'density
    ''    'Dim lProp2Desired As Int32 = 0      'compress
    ''    'Dim lProp3Desired As Int32 = 7      'magreact

    ''    'Dim lTemp As Int32 = oMineral.GetPropertyIndex(1)
    ''    'Dim lProp1Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(5)
    ''    'Dim lProp2Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(14)
    ''    'Dim lProp3Actual As Int32 = oMineral.MinPropValueScore(lTemp)

    ''    'lProp1Actual = Math.Abs(lProp1Desired - lProp1Actual)
    ''    'lProp2Actual = Math.Abs(lProp2Desired - lProp2Actual)
    ''    'lProp3Actual = Math.Abs(lProp3Desired - lProp3Actual)
    ''    'Dim lMaxDA As Int32 = Math.Max(Math.Max(lProp2Actual, lProp3Actual), lProp1Actual)

    ''    'TODO: Remove me when ready!!!
    ''    If oMineral.ObjectID = 1 Then
    ''        lMaxDA = 5 * 23
    ''    ElseIf oMineral.ObjectID = 5 Then
    ''        lMaxDA = 23
    ''    ElseIf oMineral.ObjectID = 9 Then
    ''        lMaxDA = 255
    ''    End If
    ''    Dim decTemp As Decimal = (1D + lMaxDA) * ((blPowGen + blThrustGen) / 2D + (CDec(blMaxSpd) + CDec(blMan) + 1D)) / 50D
    ''    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    ''    If decTemp < 0 Then Return 0
    ''    Return CInt(decTemp)
    ''End Function
    ''Private Function CalculateBaseDriveFrame() As Int32
    ''    Dim oMineral As Mineral = Nothing
    ''    If cboDriveFrame.ListIndex > -1 Then
    ''        Dim lID As Int32 = cboDriveFrame.ItemData(cboDriveFrame.ListIndex)
    ''        For X As Int32 = 0 To glMineralUB
    ''            If glMineralIdx(X) = lID Then
    ''                oMineral = goMinerals(X)
    ''                Exit For
    ''            End If
    ''        Next X
    ''    End If
    ''    If oMineral Is Nothing Then Return 0

    ''    Dim lTempMax As Int32 = 0
    ''    If muPropReqs Is Nothing = False Then
    ''        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    ''            If muPropReqs(X).lMinID = 4 Then
    ''                Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
    ''                lTempMax = Math.Max(GetDesiredActualRelationshipValue(muPropReqs(X).lPropVal, lA), lTempMax)
    ''            End If
    ''        Next X
    ''    End If
    ''    Dim lMaxDA As Int32 = lTempMax
    ''    'Dim lProp1Desired As Int32 = 7      'superc
    ''    'Dim lProp2Desired As Int32 = 0      'compress
    ''    'Dim lProp3Desired As Int32 = 6      'hardness

    ''    'Dim lTemp As Int32 = oMineral.GetPropertyIndex(10)
    ''    'Dim lProp1Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(5)
    ''    'Dim lProp2Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(2)
    ''    'Dim lProp3Actual As Int32 = oMineral.MinPropValueScore(lTemp)

    ''    'lProp1Actual = Math.Abs(lProp1Desired - lProp1Actual)
    ''    'lProp2Actual = Math.Abs(lProp2Desired - lProp2Actual)
    ''    'lProp3Actual = Math.Abs(lProp3Desired - lProp3Actual)
    ''    'Dim lMaxDA As Int32 = Math.Max(Math.Max(lProp2Actual, lProp3Actual), lProp1Actual)

    ''    'TODO: Remove me when ready!!!
    ''    If oMineral.ObjectID = 1 Then
    ''        lMaxDA = 5 * 23
    ''    ElseIf oMineral.ObjectID = 5 Then
    ''        lMaxDA = 23
    ''    ElseIf oMineral.ObjectID = 9 Then
    ''        lMaxDA = 255
    ''    End If
    ''    Dim decTemp As Decimal = (1D + lMaxDA) * ((blPowGen + blThrustGen) / 2D + (CDec(blMaxSpd) + CDec(blMan) + 1D)) / 50D
    ''    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    ''    If decTemp < 0 Then Return 0
    ''    Return CInt(decTemp)
    ''End Function
    ''Private Function CalculateBaseDriveMeld() As Int32
    ''    Dim oMineral As Mineral = Nothing
    ''    If cboDriveMeld.ListIndex > -1 Then
    ''        Dim lID As Int32 = cboDriveMeld.ItemData(cboDriveMeld.ListIndex)
    ''        For X As Int32 = 0 To glMineralUB
    ''            If glMineralIdx(X) = lID Then
    ''                oMineral = goMinerals(X)
    ''                Exit For
    ''            End If
    ''        Next X
    ''    End If
    ''    If oMineral Is Nothing Then Return 0

    ''    Dim lTempMax As Int32 = 0
    ''    If muPropReqs Is Nothing = False Then
    ''        For X As Int32 = 0 To muPropReqs.GetUpperBound(0)
    ''            If muPropReqs(X).lMinID = 5 Then
    ''                Dim lA As Int32 = oMineral.MinPropValueScore(oMineral.GetPropertyIndex(muPropReqs(X).lPropID))
    ''                lTempMax = Math.Max(GetDesiredActualRelationshipValue(muPropReqs(X).lPropVal, lA), lTempMax)
    ''            End If
    ''        Next X
    ''    End If
    ''    Dim lMaxDA As Int32 = lTempMax
    ''    'Dim lProp1Desired As Int32 = 6      'chemreact
    ''    'Dim lProp2Desired As Int32 = 0      'combust
    ''    'Dim lProp3Desired As Int32 = 0      'malleable

    ''    'Dim lTemp As Int32 = oMineral.GetPropertyIndex(18)
    ''    'Dim lProp1Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(19)
    ''    'Dim lProp2Actual As Int32 = oMineral.MinPropValueScore(lTemp)
    ''    'lTemp = oMineral.GetPropertyIndex(3)
    ''    'Dim lProp3Actual As Int32 = oMineral.MinPropValueScore(lTemp)

    ''    'lProp1Actual = Math.Abs(lProp1Desired - lProp1Actual)
    ''    'lProp2Actual = Math.Abs(lProp2Desired - lProp2Actual)
    ''    'lProp3Actual = Math.Abs(lProp3Desired - lProp3Actual)
    ''    'Dim lMaxDA As Int32 = Math.Max(Math.Max(lProp2Actual, lProp3Actual), lProp1Actual)

    ''    'TODO: Remove me when ready!!!
    ''    If oMineral.ObjectID = 1 Then
    ''        lMaxDA = 5 * 23
    ''    ElseIf oMineral.ObjectID = 5 Then
    ''        lMaxDA = 23
    ''    ElseIf oMineral.ObjectID = 9 Then
    ''        lMaxDA = 255
    ''    End If
    ''    Dim decTemp As Decimal = (1D + lMaxDA) * ((blPowGen + blThrustGen) / 2D + (CDec(blMaxSpd) + CDec(blMan) + 1D)) / 50D
    ''    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    ''    If decTemp < 0 Then Return 0
    ''    Return CInt(decTemp)
    ''End Function

    'Protected Overrides Function CalculateBaseMineralCosts(ByVal lMaxDA As Int32, ByVal lSumDAForMineral As Int32) As Integer
    '    Dim decTemp As Decimal = (1D + lMaxDA) * ((blPowGen + blThrustGen) / 2D + (CDec(blMaxSpd) + CDec(blMan) + 1D)) / 150D
    '    decTemp *= lSumDAForMineral
    '    If decTemp > Int32.MaxValue Then Return Int32.MaxValue
    '    If decTemp < 0 Then Return 0
    '    Return CInt(decTemp)
    'End Function

    'Protected Overrides Function CalculateBasePower() As Integer
    '    Return 0
    'End Function

    'Protected Overrides Function CalculateThresholdValue() As Decimal
    '    Dim lMaxGuns As Int32 = 1
    '    Dim decNormalize As Decimal
    '    GetTypeValues(lHullTypeID, decNormalize, lMaxGuns, 0, 0, 0, 0)
    '    Dim decBA As Decimal = CDec(lMaxGuns) / decNormalize '/ CDec(lMaxGuns)
    '    Dim decThreshold As Decimal = CDec(blPowGen) * CDec(blThrustGen) * CDec(blMan) * CDec(blMaxSpd) * decBA
    '    Return decThreshold
    'End Function

    Protected Overrides Sub SetMinReqProps()
        If lHullTypeID < 0 Then Return
        Dim yTemp(,) As Byte = {{0, 6, 6, 5, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 5, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 5, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {1, 6, 6, 6, 6, 1, 1, 6, 6, 7, 1, 1, 7, 6, 1, 6, 1, 1}, _
        {1, 6, 7, 6, 6, 1, 1, 6, 6, 7, 1, 1, 7, 6, 1, 6, 1, 1}, _
        {1, 7, 7, 6, 6, 1, 1, 7, 6, 7, 1, 1, 7, 6, 1, 7, 1, 1}, _
        {1, 7, 7, 6, 7, 1, 1, 7, 6, 8, 1, 1, 7, 6, 1, 7, 1, 1}, _
        {1, 7, 7, 6, 7, 1, 1, 7, 7, 8, 1, 1, 8, 7, 1, 7, 1, 1}, _
        {1, 7, 7, 6, 7, 1, 1, 7, 7, 8, 1, 1, 8, 7, 1, 7, 1, 1}, _
        {0, 6, 6, 5, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 5, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}, _
        {0, 6, 6, 6, 6, 0, 0, 6, 6, 7, 0, 0, 7, 6, 0, 6, 0, 0}}

        ReDim muPropReqs(17)
        Dim lPropID() As Int32 = {15, 1, 8, 3, 2, 5, 3, 2, 8, 14, 5, 1, 10, 2, 5, 18, 3, 19}
        For X As Int32 = 0 To lPropID.GetUpperBound(0)
            muPropReqs(X).lPropID = lPropID(X)
            muPropReqs(X).lPropVal = yTemp(lHullTypeID, X)
            muPropReqs(X).lMinID = X \ 3
        Next X
    End Sub

    'Protected Overrides Sub LoadPointVals()
    '    'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
    '    'If sPath.EndsWith("\") = False Then sPath &= "\"
    '    'sPath &= "TECH.dat"
    '    'Dim oINI As New InitFile(sPath)
    '    'Dim sHeader As String = "ENGINE"
    '    ml_POINT_PER_HULL = 600000 'CInt(Val(oINI.GetString(sHeader, "PointPerHull", "600000")))
    '    ml_POINT_PER_POWER = 0
    '    ml_POINT_PER_RES_TIME = 1 '50 'CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "5")))
    '    ml_POINT_PER_RES_COST = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "1")))
    '    ml_POINT_PER_PROD_TIME = 1 'CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "1")))
    '    ml_POINT_PER_PROD_COST = 72 'CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "72")))
    '    ml_POINT_PER_COLONIST = 6 'CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "6")))
    '    ml_POINT_PER_ENLISTED = 60 'CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "60")))
    '    ml_POINT_PER_OFFICER = 240 'CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "240")))

    '    ml_POINT_PER_MIN1 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerSBody", "600")))
    '    ml_POINT_PER_MIN2 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerSFrame", "600")))
    '    ml_POINT_PER_MIN3 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerSMeld", "600")))
    '    ml_POINT_PER_MIN4 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerDBody", "600")))
    '    ml_POINT_PER_MIN5 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerDFrame", "600")))
    '    ml_POINT_PER_MIN6 = 600 'CInt(Val(oINI.GetString(sHeader, "PointPerDMeld", "600")))

    '    'oINI.WriteString(sHeader, "PointPerHull", ml_POINT_PER_HULL.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResTime", ml_POINT_PER_RES_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerResCost", ml_POINT_PER_RES_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdTime", ml_POINT_PER_PROD_TIME.ToString)
    '    'oINI.WriteString(sHeader, "PointPerProdCost", ml_POINT_PER_PROD_COST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerColonist", ml_POINT_PER_COLONIST.ToString)
    '    'oINI.WriteString(sHeader, "PointPerEnlisted", ml_POINT_PER_ENLISTED.ToString)
    '    'oINI.WriteString(sHeader, "PointPerOfficer", ml_POINT_PER_OFFICER.ToString)

    '    'oINI.WriteString(sHeader, "PointPerSBody", ml_POINT_PER_MIN1.ToString)
    '    'oINI.WriteString(sHeader, "PointPerSFrame", ml_POINT_PER_MIN2.ToString)
    '    'oINI.WriteString(sHeader, "PointPerSMeld", ml_POINT_PER_MIN3.ToString)
    '    'oINI.WriteString(sHeader, "PointPerDBody", ml_POINT_PER_MIN4.ToString)
    '    'oINI.WriteString(sHeader, "PointPerDFrame", ml_POINT_PER_MIN5.ToString)
    '    'oINI.WriteString(sHeader, "PointPerDMeld", ml_POINT_PER_MIN6.ToString)
    '    'oINI = Nothing
    'End Sub

    Protected Overrides Function GetCoefficient(ByVal lLookup As TechBuilderComputer.elPropCoeffLookup) As Decimal
        'In HullType lookup enum order first dimension
        'the second dimension is based on the lLookup coeff
        'do not call directly, instead, call TechBuilderCompter.GetCoeffValue
 

        'Dim decCoeffList(,) As Decimal = {{6.9D, 0D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 14D, 14D, 14D, 14D, 14D, 14D}, _
        ' {7.8D, 0D, 1D, 1D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 5.9D, 6.3D, 6.4D, 6D, 6D, 6D}, _
        ' {6D, 0D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 11.8D, 11.8D, 11.8D, 11.8D, 11.8D, 11.8D}, _
        ' {7.8D, 0D, 1D, 1D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 5.9D, 6.3D, 6.4D, 6D, 6D, 6D}, _
        ' {4.87D, 0D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D}, _
        ' {4.3D, 0D, 0D, 0D, 19D, 0.82D, 0.6D, 1.19D, 1.36D, 7.12D, 7.23D, 7.8D, 7.1D, 7.3D, 7.6D}, _
        ' {4D, 0D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 6D, 6D, 6D, 6D, 6D, 6D}, _
        ' {2.3D, 0D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D}, _
        ' {2.55D, 0D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 1.03D, 1.214D, 3D, 3D, 3D, 3D, 3D, 3D}, _
        ' {2.186D, 0D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.02D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D}, _
        ' {1.99D, 0D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.01D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D}, _
        ' {1.84D, 0D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.005D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D}, _
        ' {1.64D, 0D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 1D, 1D, 1.75D, 1.75D, 1.75D, 1.75D, 1.75D, 1.75D}, _
        ' {1.46D, 0D, 1.4D, 1.83D, 2.01D, 0.477D, 0.445D, 1D, 1.015D, 1.55D, 1.55D, 1.55D, 1.55D, 1.55D, 1.55D}, _
        ' {2.3D, 0D, 0D, 0D, 0D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D}, _
        ' {1.84D, 0D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.005D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D}, _
        ' {1.84D, 0D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 1.005D, 1.015D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D, 2.01D}, _
        ' {1.99D, 0D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 1.01D, 1.128D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D, 2.216D}, _
        ' {2.186D, 0D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.02D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D}, _
        ' {2.3D, 0D, 3D, 3.55D, 5D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D}, _
        ' {2.186D, 0D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 1.02D, 1.2D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D, 2.48D}, _
        ' {4D, 0D, 1D, 1D, 19D, 0.6D, 0.6D, 1.16D, 1.27D, 6D, 6D, 6D, 6D, 6D, 6D}, _
        ' {4.87D, 0D, 0D, 0D, 19D, 0.834D, 0.655D, 1.205D, 1.443D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D, 7.98D}}
        Dim decCoeffList(,) As Decimal = {{6.3D, 0D, 0D, 0D, 19D, 0.93D, 0.667D, 1.3D, 1.443D, 14D, 14D, 14D, 14D, 14D, 14D}, _
            {7.1D, 0D, 1D, 1D, 19D, 0.82D, 0.6D, 1.23D, 1.2D, 5.9D, 5.9D, 5.9D, 6D, 6D, 6D}, _
            {5.4D, 0D, 1D, 1D, 19D, 0.82D, 0.6D, 1.07D, 1.1D, 5.9D, 5.9D, 5.9D, 6D, 6D, 6D}, _
            {2.5D, 0D, 1D, 1D, 19D, 0.82D, 0.6D, 1.1D, 1.1D, 2.8D, 2.8D, 2.8D, 2.8D, 2.8D, 2.8D}, _
            {4.46D, 0D, 0D, 0D, 19D, 0.834D, 0.655D, 1.05D, 1.1D, 7D, 7D, 7D, 7D, 7D, 7D}, _
            {4.12D, 0D, 0D, 0D, 19D, 0.82D, 0.6D, 1.2D, 1.36D, 6.5D, 6.5D, 6.5D, 6.5D, 6.5D, 6.5D}, _
            {3.651D, 0D, 1D, 1D, 19D, 0.6D, 0.6D, 0.98D, 1.12D, 5.4D, 5.4D, 5.4D, 5.4D, 5.4D, 5.4D}, _
            {2.082D, 0D, 3D, 3.55D, 5D, 0.6D, 0.5D, 0.928D, 0.967D, 2.58D, 2.58D, 2.58D, 2.58D, 2.58D, 2.58D}, _
            {2.55D, 0D, 3.6D, 4.21D, 6.6D, 0.76D, 0.55D, 0.934D, 1D, 3.3D, 3.3D, 3.3D, 3.3D, 3.3D, 3.3D}, _
            {1.937D, 0D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.895D, 0.889D, 2.375D, 2.375D, 2.375D, 2.375D, 2.375D, 2.375D}, _
            {1.765D, 0D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.8915D, 0.865D, 2.12D, 2.12D, 2.12D, 2.12D, 2.12D, 2.12D}, _
            {1.615D, 0D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.879D, 0.85D, 1.9D, 1.9D, 1.9D, 1.9D, 1.9D, 1.9D}, _
            {1.43D, 0D, 1.6D, 2.11D, 2.41D, 0.46D, 0.43D, 0.869D, 0.825D, 1.65D, 1.65D, 1.65D, 1.65D, 1.65D, 1.65D}, _
            {1.46D, 0D, 1.4D, 1.83D, 2.01D, 0.477D, 0.445D, 1D, 1.015D, 1.55D, 1.55D, 1.55D, 1.55D, 1.55D, 1.55D}, _
            {2.3D, 0D, 0D, 0D, 0D, 0.6D, 0.5D, 1.1D, 1.27D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D, 2.66D}, _
            {1.615D, 0D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.879D, 0.85D, 1.9D, 1.9D, 1.9D, 1.9D, 1.9D, 1.9D}, _
            {1.615D, 0D, 2D, 2.5D, 3.04D, 0.48D, 0.45D, 0.879D, 0.85D, 1.9D, 1.9D, 1.9D, 1.9D, 1.9D, 1.9D}, _
            {1.765D, 0D, 2.23D, 2.82D, 3.545D, 0.58D, 0.486D, 0.8915D, 0.865D, 2.12D, 2.12D, 2.12D, 2.12D, 2.12D, 2.12D}, _
            {1.937D, 0D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.895D, 0.889D, 2.375D, 2.375D, 2.375D, 2.375D, 2.375D, 2.375D}, _
            {2.082D, 0D, 3D, 3.55D, 5D, 0.6D, 0.5D, 0.928D, 0.967D, 2.58D, 2.58D, 2.58D, 2.58D, 2.58D, 2.58D}, _
            {1.937D, 0D, 2.6D, 3.27D, 4.8D, 0.6D, 0.5D, 0.895D, 0.889D, 2.375D, 2.375D, 2.375D, 2.375D, 2.375D, 2.375D}, _
            {3.651D, 0D, 1D, 1D, 19D, 0.6D, 0.6D, 0.98D, 1.12D, 5.4D, 5.4D, 5.4D, 5.4D, 5.4D, 5.4D}, _
            {4.46D, 0D, 0D, 0D, 19D, 0.834D, 0.655D, 1.05D, 1.1D, 7D, 7D, 7D, 7D, 7D, 7D}}

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
        Dim decNorm As Decimal = 0
        Dim lMaxGuns As Int32 = 0
        Dim lMaxDPS As Int32 = 0
        Dim lMaxHullSize As Int32 = 0
        Dim lHullAvail As Int32 = 0
        Dim lPower As Int32 = 0
        GetTypeValues(Me.lHullTypeID, decNorm, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lPower)

        Dim decPowRating As Decimal = blPowGen

        If blPowGen > lPower Then
            decPowRating = CDec(Math.Pow(blPowGen - lPower, 3)) + blPowGen
        End If

        If Me.lHullTypeID = Base_Tech.eyHullType.Facility OrElse Me.lHullTypeID = Base_Tech.eyHullType.SpaceStation Then
            Return CDec(Math.Pow(decPowRating, 2.183) * lMaxGuns) / decNorm
        Else
            Return CDec(blThrustGen * decPowRating * blMan * blMaxSpd * Math.Max(1, lMaxGuns)) / (lMaxHullSize / 90D) '(blThrustGen / 90D)
        End If

    End Function
End Class
