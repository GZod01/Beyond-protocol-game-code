Option Strict On

Public Enum elBuyOrderPropID As Integer
    eArmorBeamResist = -1
    eArmorImpactResist = -2
    eArmorPierceResist = -3
    eArmorMagneticResist = -4
    eArmorChemicalResist = -5
    eArmorBurnResist = -6
    eArmorRadarResist = -7
    eArmorHullUsagePerPlate = -8
    eArmorHPPerPlate = -9
    eArmorIntegrity = -10

    ePulseAccuracy = -21
    ePulseAOE = -22
    ePulseMaxDmg = -23
    ePulseMinDmg = -24
    ePulseRange = -25
    ePulseROF = -26
    ePulseHull = -27
    ePulsePower = -28

    eSolidAccuracy = -41
    eSolidRange = -42
    eSolidROF = -43
    eSolidMaxDmg = -44
    eSolidMinDmg = -45
    eSolidDmgType = -46
    eSolidHull = -47
    eSolidPower = -48

    eEngineThrust = -61
    eEngineManeuver = -62
    eEngineMaxSpeed = -63
    eEnginePowerProd = -64
    eEngineHull = -65

    eMissileROF = -81
    eMissileMaxDmg = -82
    eMissileMissileHullSize = -83
    eMissileMaxSpeed = -84
    eMissileManeuver = -85
    eMissileFlightTime = -86
    eMissileAccuracy = -87
    eMissilePayload = -88
    eMissileAOE = -89
    eMissileStructHP = -90
    eMissileHull = -91
    eMissilePower = -92

    eProjectileROF = -101
    eProjectileMaxRng = -102
    eProjectilePayload = -103
    eProjectileAOE = -104
    eProjectileAmmoSize = -105
    eProjectileMaxDmg = -106
    eProjectileMinDmg = -107
    eProjectilePierceRatio = -108
    eProjectileHull = -109
    eProjectilePower = -110

    eRadarAccuracy = -121
    eRadarScanRes = -122
    eRadarOptRng = -123
    eRadarMaxRng = -124
    eRadarDisRes = -125
    eRadarJamImm = -126
    eRadarJamStr = -127
    eRadarJamTargets = -128
    eRadarJamEffect = -129
    eRadarPower = -130
    eRadarHull = -131

    eShieldMaxHP = -141
    eShieldRechargeRate = -142
    eShieldRechargeFreq = -143
    eShieldProjectionHullSize = -144
    eShieldPower = -145
    eShieldHull = -146

    eSpecificMineralID = -150
End Enum
Public Enum BuyOrderCompareTypes As Byte
    eLessThan = 0
    eGreaterThan = 1
    eEqualTo = 2
    eLessThanEqualTo = 3
    eGreaterThanEqualTo = 4
    eNotEqualTo = 5
End Enum

Public Class BuyOrder
    Inherits TradeOrder

    Public yBuyOrderState As Byte
    Public lSpecificID As Int32 = -1
    Public iTradeTypeID As Int16
    Public lAcceptedByID As Int32 = -1
    Public lAcceptedOn As Int32
    Public lDeadline As Int32       'seconds remaining
    Public blEscrow As Int64 

	Private mdtReceivedDetails As Date = Date.MinValue
	Private mdtReceivedMsg As Date = Date.MinValue

    Private Structure BuyOrderProperty
        Public lPropertyID As Int32
        Public lPropertyValue As Int32
        Public yCompareType As Byte             '0 = EQUALS, 1 = LESS THAN, 2 = GREATER THAN
    End Structure
    Private muProperties() As BuyOrderProperty
    Private mlPropertyUB As Int32 = -1

    Public Overrides Function GetCurrentListText() As String
        '"Escrow               Payment                Deadline"
        Dim sResult As String = blEscrow.ToString("#,##0").PadRight(21, " "c)
        sResult &= blPrice.ToString("#,##0").PadRight(23, " "c)

        'Dim lSeconds As Int32 = lDeadline
        'Dim lMinutes As Int32 = lSeconds \ 60
        'lSeconds -= (lMinutes * 60)
        'Dim lHours As Int32 = lMinutes \ 60
        'lMinutes -= (lHours * 60)
        'Dim lDays As Int32 = lHours \ 24
        'lHours -= (lDays * 24)

        sResult &= GetDeadlineText() ' lDays.ToString("00") & ":" & lHours.ToString("00") & ":" & lMinutes.ToString("00") & ":" & lSeconds.ToString("00")
        Return sResult
    End Function

    Public Overrides Sub RequestDetails(ByVal lFromTPID As Integer)
        If mbDetailsRequested = True Then Return
        mbDetailsRequested = True

        Dim yMsg(16) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetOrderSpecifics).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, 2)
        yMsg(6) = 1         'buy order
        System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 7)
        System.BitConverter.GetBytes(-1S).CopyTo(yMsg, 11)
        System.BitConverter.GetBytes(lFromTPID).CopyTo(yMsg, 13)
        goUILib.SendMsgToPrimary(yMsg)
    End Sub

    Public Overrides Sub SetFromDetailMsg(ByRef yData() As Byte, ByVal lPos As Integer)
        mdtReceivedDetails = Now

        lDeliveryTime = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSellerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecificID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        iTradeTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        lAcceptedByID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lAcceptedOn = System.BitConverter.ToInt32(yData, lPos) : lPos += 4


        mlPropertyUB = yData(lPos) - 1 : lPos += 1
        ReDim muProperties(mlPropertyUB)
        For X As Int32 = 0 To mlPropertyUB
            With muProperties(X)
                .lPropertyID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lPropertyValue = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yCompareType = yData(lPos) : lPos += 1
            End With
        Next X
    End Sub

	Public Overrides Sub SetFromMsg(ByRef yData() As Byte, ByVal lPos As Integer)
		mdtReceivedMsg = Now
		yRouteType = yData(lPos) : lPos += 1
		lTradePostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yBuyOrderState = yData(lPos) : lPos += 1
		iTradeTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		lDeadline = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		blEscrow = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
		blQty = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
		blPrice = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
		lAcceptedByID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	End Sub

    Public Function GetItemDetailsText() As String
        Dim oSB As New System.Text.StringBuilder()
        For X As Int32 = 0 To mlPropertyUB

            Dim sLine As String = ""

            If muProperties(X).lPropertyID = elBuyOrderPropID.eSpecificMineralID Then
                Dim lMinID As Int32 = muProperties(X).lPropertyValue
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To glMineralUB
                    If glMineralIdx(Y) = lMinID Then
                        sLine = goMinerals(Y).MineralName
                        Exit For
                    End If
                Next Y
                If bFound = False Then sLine = "Unknown Mineral"
            Else
                sLine = GetBuyOrderPropertyText(muProperties(X).lPropertyID)
                sLine &= ": " & GetBuyOrderCompareTypeText(muProperties(X).yCompareType)

                Dim lPropID As Int32 = muProperties(X).lPropertyID
                Dim bTime As Boolean = lPropID = elBuyOrderPropID.ePulseROF OrElse lPropID = elBuyOrderPropID.eSolidROF OrElse _
                  lPropID = elBuyOrderPropID.eMissileROF OrElse lPropID = elBuyOrderPropID.eMissileFlightTime OrElse _
                  lPropID = elBuyOrderPropID.eProjectileROF OrElse lPropID = elBuyOrderPropID.eShieldRechargeFreq

                If bTime = True Then
                    sLine &= (muProperties(X).lPropertyValue / 30.0F).ToString("#,###.#0")
                Else
                    sLine &= muProperties(X).lPropertyValue
                End If
            End If
            oSB.AppendLine(sLine)
        Next X
        Return oSB.ToString
    End Function

    Public Function GetDeadlineText() As String
		If mdtReceivedMsg = Date.MinValue Then Return ""

		Dim lSecondsMod As Int32 = Math.Abs(CInt(Now.Subtract(mdtReceivedMsg).TotalSeconds))

        lSecondsMod = lDeadline - lSecondsMod
        If lSecondsMod < 0 Then Return "Eminent"

        Return GetDurationFromSeconds(lSecondsMod, False)
    End Function

    Public Function GetAdjustedDeadline() As String
		If mdtReceivedMsg = Date.MinValue Then Return ""

		Dim lAdjustedDeadline As Int32 = lDeadline - Math.Abs(CInt(Now.Subtract(mdtReceivedMsg).TotalSeconds))
        lAdjustedDeadline -= Me.lDeliveryTime

        'Ok, now adjusted deadline is our value
        If lAdjustedDeadline < 0 Then Return "Exceeded"
        Return GetDurationFromSeconds(lAdjustedDeadline, False)
    End Function

    Public Function GetAcceptedOnText() As String
		If mdtReceivedDetails = Date.MinValue Then Return ""
        Dim lSecondsMod As Int32 = Math.Abs(CInt(Now.Subtract(mdtReceivedDetails).TotalSeconds))
        lSecondsMod = lAcceptedOn + lSecondsMod

        Dim dtTemp As Date = Now.Add(New TimeSpan(0, 0, lSecondsMod))
        Return dtTemp.ToString("MM/dd")
    End Function

    Public Function GetAcceptRequest(ByVal lBuyerTradePostID As Int32) As Byte()
        Dim yMsg(13) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eAcceptBuyOrder).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lBuyerTradePostID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, lPos) : lPos += 4

        Return yMsg
    End Function

    Public Shared Function GetBuyOrderPropertyText(ByVal lPropertyID As Int32) As String
        If lPropertyID > 0 Then
            For X As Int32 = 0 To glMineralPropertyUB
                If glMineralPropertyIdx(X) = lPropertyID Then Return goMineralProperty(X).MineralPropertyName
            Next X
        Else
            Select Case CType(lPropertyID, elBuyOrderPropID)
                Case elBuyOrderPropID.eArmorBeamResist
                    Return "Beam Resist"
                Case elBuyOrderPropID.eArmorImpactResist
                    Return "Impact Resist"
                Case elBuyOrderPropID.eArmorPierceResist
                    Return "Pierce Resist"
                Case elBuyOrderPropID.eArmorMagneticResist
                    Return "Magnetic Resist"
                Case elBuyOrderPropID.eArmorChemicalResist
                    Return "Chemical Resist"
                Case elBuyOrderPropID.eArmorBurnResist
                    Return "Burn Resist"
                Case elBuyOrderPropID.eArmorRadarResist
                    Return "Radar Resist"
                Case elBuyOrderPropID.eArmorHullUsagePerPlate
                    Return "Hull Usage Per Plate"
                Case elBuyOrderPropID.eArmorHPPerPlate
                    Return "Hitpoints Per Plate"
                Case elBuyOrderPropID.eArmorIntegrity
                    Return "Integrity"
                Case elBuyOrderPropID.ePulseAccuracy
                    Return "Accuracy"
                Case elBuyOrderPropID.ePulseAOE
                    Return "Scatter Radius"
                Case elBuyOrderPropID.ePulseMaxDmg
                    Return "Maximum Damage"
                Case elBuyOrderPropID.ePulseMinDmg
                    Return "Minimum Damage"
                Case elBuyOrderPropID.ePulseRange
                    Return "Range"
                Case elBuyOrderPropID.ePulseROF
                    Return "ROF"
                Case elBuyOrderPropID.ePulseHull
                    Return "Hull Required"
                Case elBuyOrderPropID.ePulsePower
                    Return "Power Required"
                Case elBuyOrderPropID.eSolidAccuracy
                    Return "Accuracy"
                Case elBuyOrderPropID.eSolidRange
                    Return "Range"
                Case elBuyOrderPropID.eSolidROF
                    Return "ROF"
                Case elBuyOrderPropID.eSolidMaxDmg
                    Return "Maximum Damage"
                Case elBuyOrderPropID.eSolidMinDmg
                    Return "Minimum Damage"
                Case elBuyOrderPropID.eSolidDmgType
                    Return "Damage Type"
                Case elBuyOrderPropID.eSolidHull
                    Return "Hull Required"
                Case elBuyOrderPropID.eSolidPower
                    Return "Power Required"
                Case elBuyOrderPropID.eEngineThrust
                    Return "Thrust"
                Case elBuyOrderPropID.eEngineManeuver
                    Return "Maneuver"
                Case elBuyOrderPropID.eEngineMaxSpeed
                    Return "Max Speed"
                Case elBuyOrderPropID.eEnginePowerProd
                    Return "Power Production"
                Case elBuyOrderPropID.eEngineHull
                    Return "Hull Required"
                Case elBuyOrderPropID.eMissileROF
                    Return "ROF"
                Case elBuyOrderPropID.eMissileMaxDmg
                    Return "Maximum Damage"
                Case elBuyOrderPropID.eMissileMissileHullSize
                    Return "Missile Size"
                Case elBuyOrderPropID.eMissileMaxSpeed
                    Return "Max Speed"
                Case elBuyOrderPropID.eMissileManeuver
                    Return "Maneuver"
                Case elBuyOrderPropID.eMissileFlightTime
                    Return "Flight Time"
                Case elBuyOrderPropID.eMissileAccuracy
                    Return "Accuracy"
                Case elBuyOrderPropID.eMissilePayload
                    Return "Payload Type"
                Case elBuyOrderPropID.eMissileAOE
                    Return "Explosion Radius"
                Case elBuyOrderPropID.eMissileStructHP
                    Return "Structure Hitpoints"
                Case elBuyOrderPropID.eMissileHull
                    Return "Hull Required"
                Case elBuyOrderPropID.eMissilePower
                    Return "Power Required"
                Case elBuyOrderPropID.eProjectileROF
                    Return "ROF"
                Case elBuyOrderPropID.eProjectileMaxRng
                    Return "Range"
                Case elBuyOrderPropID.eProjectilePayload
                    Return "Payload Type"
                Case elBuyOrderPropID.eProjectileAOE
                    Return "Explosion Radius"
                Case elBuyOrderPropID.eProjectileAmmoSize
                    Return "Ammo Size"
                Case elBuyOrderPropID.eProjectileMaxDmg
                    Return "Maximum Damage"
                Case elBuyOrderPropID.eProjectileMinDmg
                    Return "Minimum Damage"
                Case elBuyOrderPropID.eProjectilePierceRatio
                    Return "Pierce Ratio"
                Case elBuyOrderPropID.eProjectileHull
                    Return "Hull Required"
                Case elBuyOrderPropID.eProjectilePower
                    Return "Power Required"
                Case elBuyOrderPropID.eRadarAccuracy
                    Return "Weapon Accuracy"
                Case elBuyOrderPropID.eRadarScanRes
                    Return "Scan Resolution"
                Case elBuyOrderPropID.eRadarOptRng
                    Return "Optimum Range"
                Case elBuyOrderPropID.eRadarMaxRng
                    Return "Maximum Range"
                Case elBuyOrderPropID.eRadarDisRes
                    Return "Disruption Resist"
                Case elBuyOrderPropID.eRadarJamImm
                    Return "Jamming Immunity"
                Case elBuyOrderPropID.eRadarJamStr
                    Return "Jamming Strength"
                Case elBuyOrderPropID.eRadarJamTargets
                    Return "Jam Targets"
                Case elBuyOrderPropID.eRadarJamEffect
                    Return "Jam Effect"
                Case elBuyOrderPropID.eRadarPower
                    Return "Power Required"
                Case elBuyOrderPropID.eRadarHull
                    Return "Hull Required"
                Case elBuyOrderPropID.eShieldMaxHP
                    Return "Max Hitpoints"
                Case elBuyOrderPropID.eShieldRechargeRate
                    Return "Recharge Rate"
                Case elBuyOrderPropID.eShieldRechargeFreq
                    Return "Recharge Frequency"
                Case elBuyOrderPropID.eShieldProjectionHullSize
                    Return "Projection Hull Size"
                Case elBuyOrderPropID.eShieldPower
                    Return "Power Required"
                Case elBuyOrderPropID.eShieldHull
                    Return "Hull Required"
            End Select
        End If


        Return "Unknown Property"
    End Function
    Public Shared Function GetBuyOrderCompareTypeText(ByVal yCompareType As Byte) As String
        Select Case CType(yCompareType, BuyOrderCompareTypes)
            Case BuyOrderCompareTypes.eEqualTo
                Return "="
            Case BuyOrderCompareTypes.eGreaterThan
                Return ">"
            Case BuyOrderCompareTypes.eGreaterThanEqualTo
                Return ">="
            Case BuyOrderCompareTypes.eLessThan
                Return "<"
            Case BuyOrderCompareTypes.eLessThanEqualTo
                Return "<="
            Case BuyOrderCompareTypes.eNotEqualTo
                Return "Not"
        End Select
        Return "="
    End Function
End Class


'Public Class BuyOrder
'    Public lTradePostID As Int32
'    Public lBuyOrderID As Int32
'    Public yRouteType As Byte

'    Public yBuyOrderState As Byte
'    Public lSpecificID As Int32 = -1
'    Public iTradeTypeID As Int16
'    Public lAcceptedByID As Int32 = -1
'    Public lAcceptedOn As Int32
'    Public lDeadline As Int32       'seconds remaining
'    Public blEscrow As Int64
'    Public blQuantity As Int64
'    Public blPaymentAmt As Int64

'    Private Structure BuyOrderProperty
'        Public lPropertyID As Int32
'        Public lPropertyValue As Int32
'        Public yCompareType As Byte             '0 = EQUALS, 1 = LESS THAN, 2 = GREATER THAN
'    End Structure
'    Private muProperties() As BuyOrderProperty
'    Private mlPropertyUB As Int32 = -1

'    Private mbDetailsRequested As Boolean = False
'    Public Sub RequestDetails(ByVal lFromTPID As Int32)
'        If mbDetailsRequested = True Then Return
'        mbDetailsRequested = True

'        Dim yMsg(16) As Byte
'        System.BitConverter.GetBytes(EpicaMessageCode.eGetOrderSpecifics).CopyTo(yMsg, 0)
'        System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, 2)
'        yMsg(6) = 1         'buy order
'        System.BitConverter.GetBytes(lBuyOrderID).CopyTo(yMsg, 7)
'        System.BitConverter.GetBytes(-1S).CopyTo(yMsg, 11)
'        System.BitConverter.GetBytes(lFromTPID).CopyTo(yMsg, 13)
'        goUILib.SendMsgToPrimary(yMsg)
'    End Sub

'    Public Sub SetFromMsg(ByRef yData() As Byte, ByVal lPos As Int32)
'        yRouteType = yData(lPos) : lPos += 1
'        lTradePostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'        lBuyOrderID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'        yBuyOrderState = yData(lPos) : lPos += 1
'        iTradeTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
'        lDeadline = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

'        blEscrow = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
'        blQuantity = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
'        blPaymentAmt = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
'    End Sub

'    Public Sub SetFromDetailMsg(ByRef yData() As Byte, ByVal lPos As Int32)
'        '
'    End Sub

'End Class
