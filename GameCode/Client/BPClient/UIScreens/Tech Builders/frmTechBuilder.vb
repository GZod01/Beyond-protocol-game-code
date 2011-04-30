Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public MustInherit Class frmTechBuilder
    Inherits UIWindow

    Protected mfrmBuilderCost As frmTechBuilderCost

    Public MustOverride Sub ViewResults(ByRef oTech As Base_Tech, ByVal lProdFactor As Int32)
    Protected MustOverride Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)

    Protected ml_Min1DA As Int32 = 255
    Protected ml_Min2DA As Int32 = 255
    Protected ml_Min3DA As Int32 = 255
    Protected ml_Min4DA As Int32 = 255
    Protected ml_Min5DA As Int32 = 255
    Protected ml_Min6DA As Int32 = 255

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        mfrmBuilderCost = New frmTechBuilderCost(oUILib)
        mfrmBuilderCost.Visible = True
        AddHandler mfrmBuilderCost.ValueChanged, AddressOf FrmBuilderCost_ValueChanged
    End Sub

    Private Sub FrmBuilderCost_ValueChanged()
        BuilderCostValueChange(False)
    End Sub

    Private Sub frmTechBuilder_OnResize() Handles Me.OnResize
        If mfrmBuilderCost Is Nothing = False Then
            mfrmBuilderCost.Top = Me.Top + Me.Height + 5
            mfrmBuilderCost.Left = Me.Left
        End If
    End Sub
 
    Private Sub frmTechBuilder_WindowClosed() Handles Me.WindowClosed
        MyBase.moUILib.RemoveWindow("frmTechBuilderCost")
    End Sub

    Protected Sub ReshowTechBuilderCost()
        mfrmBuilderCost.Visible = True
        mfrmBuilderCost.ResetAll()
        MyBase.moUILib.AddWindow(mfrmBuilderCost)
    End Sub

    Public Shared Function DoLockedToBaseCheck(ByRef lLockedValue As Int32, ByRef lBaseValue As Int32) As Int32
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
        AddToTestString("L2B Result: " & lResult.ToString)
        Return lResult
    End Function
    Public Shared Function DoLockedToBaseCheck(ByRef lLockedValue As Int64, ByRef lBaseValue As Int64) As Int64
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
        AddToTestString("L2B Result: " & blResult.ToString)
        Return blResult
    End Function
    Public Shared Sub AffectValueByPoints(ByRef blPoints As Int64, ByVal lPointPerValue As Int32, ByRef lValue As Int32, ByVal lModByVal As Int32)
        If blPoints < lPointPerValue Then Return
        If lValue = Int32.MaxValue Then Return
        lValue += lModByVal
        blPoints -= lPointPerValue
    End Sub
    Public Shared Sub AffectValueByPoints(ByRef blPoints As Int64, ByVal lPointPerValue As Int32, ByRef blValue As Int64, ByVal lModByVal As Int32)
        If blPoints < lPointPerValue Then Return
        If blValue = Int64.MaxValue Then Return
        blValue += lModByVal
        blPoints -= lPointPerValue
    End Sub

    Protected Shared moSB As System.Text.StringBuilder
    Protected Shared Sub AddToTestString(ByVal sLine As String)
        If moSB Is Nothing Then moSB = New System.Text.StringBuilder
        moSB.AppendLine(sLine)
    End Sub

    Protected Shared Function GetDesiredActualRelationshipValue(ByVal lD As Int32, ByVal lA As Int32) As Int32

        Return (Math.Abs(lD - lA) + 1) * 23

        'If lD = lA Then Return 23
        'If lD < lA Then
        '    Select Case lD
        '        Case 0
        '        Case 1
        '        Case 2
        '        Case 3
        '        Case 4
        '        Case 5
        '        Case 6
        '        Case 7
        '        Case 8

        '    End Select
        'ElseIf lD > lA Then
        '    '
        'End If

        ''If lValue < 23 Then '11 Then
        ''    Return 0
        ''ElseIf lValue < 46 Then ' 21 Then
        ''    Return 1
        ''ElseIf lValue < 69 Then ' 31 Then
        ''    Return 2
        ''ElseIf lValue < 92 Then ' 41 Then
        ''    Return 3
        ''ElseIf lValue < 116 Then '51 Then
        ''    Return 4
        ''ElseIf lValue < 140 Then '81 Then
        ''    Return 5
        ''ElseIf lValue < 164 Then '101 Then
        ''    Return 6
        ''ElseIf lValue < 187 Then '126 Then
        ''    Return 7
        ''ElseIf lValue < 210 Then '151 Then
        ''    Return 8
        ''ElseIf lValue < 233 Then '201 Then
        ''    Return 9
        ''Else
        ''    Return 10
        ''End If
    End Function

    'Protected Shared Sub GetTypeValues(ByVal plHullType As Int32, ByRef pdecNormalizer As Decimal, ByRef plMaxGuns As Int32, ByRef plMaxDPS As Int32, ByRef plMaxHullSize As Int32, ByRef plHullAvail As Int32, ByRef plPower As Int32)
    '    Select Case CType(plHullType, EngineTech.eyHullType)
    '        Case EngineTech.eyHullType.BattleCruiser
    '            pdecNormalizer = 2777.77777777778D : plMaxGuns = 25 : plMaxDPS = 500 : plMaxHullSize = 250000 : plHullAvail = 200000 : plPower = 12500
    '        Case EngineTech.eyHullType.Battleship
    '            pdecNormalizer = 11111.1111111111D : plMaxGuns = 20 : plMaxDPS = 1900 : plMaxHullSize = 1000000 : plHullAvail = 800000 : plPower = 46900
    '        Case EngineTech.eyHullType.Corvette
    '            pdecNormalizer = 222.222222222222D : plMaxGuns = 15 : plMaxDPS = 60 : plMaxHullSize = 20000 : plHullAvail = 16000 : plPower = 1515
    '        Case EngineTech.eyHullType.Cruiser
    '            pdecNormalizer = 888.888888888889D : plMaxGuns = 20 : plMaxDPS = 125 : plMaxHullSize = 80000 : plHullAvail = 64000 : plPower = 6250
    '        Case EngineTech.eyHullType.Destroyer
    '            pdecNormalizer = 444.444444444444D : plMaxGuns = 20 : plMaxDPS = 250 : plMaxHullSize = 40000 : plHullAvail = 32000 : plPower = 3125
    '        Case EngineTech.eyHullType.Escort
    '            pdecNormalizer = 22.2222222222222D : plMaxGuns = 8 : plMaxDPS = 21 : plMaxHullSize = 2000 : plHullAvail = 1600 : plPower = 536
    '        Case EngineTech.eyHullType.Facility
    '            pdecNormalizer = 11111.1111111111D : plMaxGuns = 20 : plMaxDPS = 1900 : plMaxHullSize = 1000000 : plHullAvail = 800000 : plPower = 46900
    '        Case EngineTech.eyHullType.Frigate
    '            pdecNormalizer = 88.8888888888889D : plMaxGuns = 12 : plMaxDPS = 55 : plMaxHullSize = 8000 : plHullAvail = 6400 : plPower = 1389
    '        Case EngineTech.eyHullType.HeavyFighter
    '            pdecNormalizer = 3.33333333333333D : plMaxGuns = 4 : plMaxDPS = 14 : plMaxHullSize = 300 : plHullAvail = 240 : plPower = 347
    '        Case EngineTech.eyHullType.LightFighter
    '            pdecNormalizer = 1 : plMaxGuns = 2 : plMaxDPS = 10 : plMaxHullSize = 90 : plHullAvail = 72 : plPower = 250
    '        Case EngineTech.eyHullType.MediumFighter
    '            pdecNormalizer = 1.88888888888889D : plMaxGuns = 3 : plMaxDPS = 12 : plMaxHullSize = 170 : plHullAvail = 136 : plPower = 303
    '        Case EngineTech.eyHullType.NavalBattleship
    '            pdecNormalizer = 2666.66666666667D : plMaxGuns = 20 : plMaxDPS = 500 : plMaxHullSize = 240000 : plHullAvail = 192000 : plPower = 12500
    '        Case EngineTech.eyHullType.NavalCarrier
    '            pdecNormalizer = 2777.77777777778D : plMaxGuns = 20 : plMaxDPS = 500 : plMaxHullSize = 250000 : plHullAvail = 200000 : plPower = 12500
    '        Case EngineTech.eyHullType.NavalCruiser
    '            pdecNormalizer = 1111.11111111111D : plMaxGuns = 15 : plMaxDPS = 150 : plMaxHullSize = 100000 : plHullAvail = 80000 : plPower = 6500
    '        Case EngineTech.eyHullType.NavalDestroyer
    '            pdecNormalizer = 500D : plMaxGuns = 10 : plMaxDPS = 200 : plMaxHullSize = 45000 : plHullAvail = 36000 : plPower = 3000
    '        Case EngineTech.eyHullType.NavalFrigate
    '            pdecNormalizer = 222.222222222222D : plMaxGuns = 10 : plMaxDPS = 75 : plMaxHullSize = 20000 : plHullAvail = 16000 : plPower = 1500
    '        Case EngineTech.eyHullType.NavalSub
    '            pdecNormalizer = 555.555555555556D : plMaxGuns = 2 : plMaxDPS = 1000 : plMaxHullSize = 50000 : plHullAvail = 40000 : plPower = 3500
    '        Case EngineTech.eyHullType.SmallVehicle
    '            pdecNormalizer = 1.33333333333333D : plMaxGuns = 5 : plMaxDPS = 10 : plMaxHullSize = 120 : plHullAvail = 96 : plPower = 250
    '        Case EngineTech.eyHullType.SpaceStation
    '            pdecNormalizer = 44444.4444444444D : plMaxGuns = 40 : plMaxDPS = 15000 : plMaxHullSize = 4000000 : plHullAvail = 3200000 : plPower = 375000
    '        Case EngineTech.eyHullType.Tank
    '            pdecNormalizer = 5.55555555555556D : plMaxGuns = 12 : plMaxDPS = 40 : plMaxHullSize = 500 : plHullAvail = 400 : plPower = 1042
    '        Case EngineTech.eyHullType.Utility
    '            pdecNormalizer = 2.22222222222222D : plMaxGuns = 1 : plMaxDPS = 4 : plMaxHullSize = 200 : plHullAvail = 160 : plPower = 100
    '    End Select
    'End Sub

    Protected Shared Function InvalidHalfBaseCheck(ByVal blBaseValue As Int64, ByVal blLockedValue As Int64, ByVal sName As String, ByVal bPersonnel As Boolean, ByRef lbl As UILabel, ByRef bIgnoreValChange As Boolean) As Boolean
        If bPersonnel = False Then
            If blLockedValue < blBaseValue \ 2 Then
                lbl.Caption = "Your scientists cannot get the " & sName & " that low."
                lbl.Visible = True
                bIgnoreValChange = False
                Return True
            End If
        Else
            If blLockedValue < Math.Min(1, blBaseValue) Then
                lbl.Caption = "Your scientists cannot get the " & sName & " that low."
                lbl.Visible = True
                bIgnoreValChange = False
                Return True
            End If
        End If

        Return False
    End Function

    Public Sub DADataReceived(ByVal yData() As Byte)
        Dim lPos As Int32 = 2

        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim ySubTypeID As Byte = yData(lPos) : lPos += 1
        Dim yHullTypeID As Byte = yData(lPos) : lPos += 1
        Dim ySeedCoeff As Byte = yData(lPos) : lPos += 1

        Dim lSeedVal As Int32 = (CInt(iTypeID) + 1) * (CInt(ySubTypeID) + 1) * (CInt(yHullTypeID) + 1) * (CInt(ySeedCoeff) + 1)

        Dim lMinDA1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMinDA2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMinDA3 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMinDA4 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMinDA5 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMinDA6 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oRand As New Random(lSeedVal)
        Dim lMaxVal As Int32 = Int32.MaxValue \ 2
        ml_Min1DA = lMinDA1 - oRand.Next(0, lMaxVal)
        ml_Min2DA = lMinDA2 - oRand.Next(0, lMaxVal)
        ml_Min3DA = lMinDA3 - oRand.Next(0, lMaxVal)
        ml_Min4DA = lMinDA4 - oRand.Next(0, lMaxVal)
        ml_Min5DA = lMinDA5 - oRand.Next(0, lMaxVal)
        ml_Min6DA = lMinDA6 - oRand.Next(0, lMaxVal)
        oRand = Nothing

        'Now, we have the real values
        BuilderCostValueChange(False)
    End Sub
    Protected Sub RequestDAValues(ByVal iTypeID As Int16, ByVal ySubTypeID As Byte, ByVal yHullTypeID As Byte, ByVal lMin1ID As Int32, ByVal lMin2ID As Int32, ByVal lMin3ID As Int32, ByVal lMin4ID As Int32, ByVal lMin5ID As Int32, ByVal lMin6ID As Int32, ByVal yPayloadType As Byte, ByVal yProjType As Byte)
        Dim yMsg(31) As Byte
        Dim lPos As Int32 = 0

        'do request here for
        System.BitConverter.GetBytes(GlobalMessageCode.eGetDAValues).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
        yMsg(lPos) = ySubTypeID : lPos += 1
        yMsg(lPos) = yHullTypeID : lPos += 1
        System.BitConverter.GetBytes(lMin1ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin2ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin3ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin4ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin5ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMin6ID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = yPayloadType : lPos += 1
        yMsg(lPos) = yProjType : lPos += 1

        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub
End Class
