Option Strict On

'Manages a Buy Order - someone wishes to buy something
Public Class BuyOrder

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

    'PK Begin
    Public BuyOrderID As Int32
    'END OF PK
    Public TradePostID As Int32         'where the buy order originates

    Public lSpecificID As Int32 = -1    'if a specific item of iTradeTypeID is requested, this indicates it
    Private miTradeTypeID As Int16 = -1 'ObjectTypeID being bought
    Public lAcceptedByID As Int32 = -1  'player id of the player who accepted this buy order
    Public lAcceptedOn As Int32 = -1    'when the buy order was accepted (date)
    Public yBuyOrderState As Byte       '0=new and unaccepted, 1=trade accepted
    Public lDeadline As Int32 = -1      'when the buy order is to be delivered by (date)
    Public blEscrow As Int64            'amount of credits pulled from Acceptor's treasury when accepted, if the trade is successful, this is put back to the acceptor, otherwise, it put in the buy order originator
    Public blQuantity As Int64          'quantity of the item required
    Public blPaymentAmt As Int64        'amount paid to seller, this amount is taken from buyer when buy order is made
    Public yBuyOrderType As Byte        'type of buy order
    Public lAcceptorTradepostID As Int32 'ID of the tradepost of the acceptor that accepted this trade

    Private Structure BuyOrderProperty
        Public lPropertyID As Int32
        Public lPropertyValue As Int32
        Public yCompareType As Byte             '0 = EQUALS, 1 = LESS THAN, 2 = GREATER THAN
    End Structure
    Private muProperties() As BuyOrderProperty
    Private mlPropertyUB As Int32 = -1

    Public Shared Function GetTradeTypeValue(ByVal yType As Byte) As Int16
        Select Case CType(yType, GalacticTradeSystem.MarketListType)
            Case GalacticTradeSystem.MarketListType.eArmorComponent
                Return ObjectType.eArmorTech
            Case GalacticTradeSystem.MarketListType.eBeamPulseComponent, GalacticTradeSystem.MarketListType.eBeamSolidComponent, GalacticTradeSystem.MarketListType.eMissileComponent, GalacticTradeSystem.MarketListType.eProjectileComponent, GalacticTradeSystem.MarketListType.eBombComponent, GalacticTradeSystem.MarketListType.eMineWeaponComponent
                Return ObjectType.eWeaponTech
            Case GalacticTradeSystem.MarketListType.eEngineComponent
                Return ObjectType.eEngineTech
            Case GalacticTradeSystem.MarketListType.eMinerals, GalacticTradeSystem.MarketListType.eAlloys
                Return ObjectType.eMineralCache
            Case GalacticTradeSystem.MarketListType.eRadarComponent
                Return ObjectType.eRadarTech
            Case GalacticTradeSystem.MarketListType.eShieldComponent
				Return ObjectType.eShieldTech
			Case Else
				Return -1
		End Select
    End Function

    Public Property iTradeTypeID() As Int16
        Get
            If miTradeTypeID = -1 Then
                miTradeTypeID = GetTradeTypeValue(yBuyOrderType)
            End If
            Return miTradeTypeID
        End Get
        Set(ByVal value As Int16)
            miTradeTypeID = value
        End Set
    End Property

    Public Const ml_OBJECT_MSG_LEN As Int32 = 43
    Public Function GetObjAsString() As Byte()
        If TradePost Is Nothing OrElse TradePost.Owner Is Nothing Then Return Nothing

        Dim yMsg(ml_OBJECT_MSG_LEN - 1) As Byte
        Dim lPos As Int32 = 0

        'PK
        System.BitConverter.GetBytes(TradePostID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(BuyOrderID).CopyTo(yMsg, lPos) : lPos += 4
        'end pk

        yMsg(lPos) = yBuyOrderState : lPos += 1
        System.BitConverter.GetBytes(iTradeTypeID).CopyTo(yMsg, lPos) : lPos += 2

        If lDeadline <> -1 Then
            System.BitConverter.GetBytes(CInt(Math.Abs(Now.Subtract(GetDateFromNumber(lDeadline)).TotalSeconds))).CopyTo(yMsg, lPos) : lPos += 4
        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        End If
        System.BitConverter.GetBytes(blEscrow).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blQuantity).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blPaymentAmt).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(lAcceptedByID).CopyTo(yMsg, lPos) : lPos += 4

        Return yMsg
    End Function

    Public Function GetObjDetailMsg(ByVal lTP_ID As Int32, ByVal lTime As Int32) As Byte()
        Dim yMsg(35 + ((mlPropertyUB + 1) * 9)) As Byte
        Dim lPos As Int32

		System.BitConverter.GetBytes(GlobalMessageCode.eGetOrderSpecifics).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lTP_ID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = 1 : lPos += 1              'buy order
        System.BitConverter.GetBytes(BuyOrderID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lTime).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(TradePost.Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

        System.BitConverter.GetBytes(lSpecificID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(iTradeTypeID).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lAcceptedByID).CopyTo(yMsg, lPos) : lPos += 4
        If lAcceptedOn <> -1 Then
            System.BitConverter.GetBytes(CInt(Now.Subtract(GetDateFromNumber(lAcceptedOn)).TotalSeconds)).CopyTo(yMsg, lPos) : lPos += 4
        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        End If

        yMsg(lPos) = CByte(mlPropertyUB + 1) : lPos += 1
        For X As Int32 = 0 To mlPropertyUB
            With muProperties(X)
                System.BitConverter.GetBytes(.lPropertyID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lPropertyValue).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = .yCompareType : lPos += 1
            End With
        Next X

        Return yMsg
    End Function

    Private moTradePost As Facility = Nothing
    Public ReadOnly Property TradePost() As Facility
        Get
            If moTradePost Is Nothing Then moTradePost = GetEpicaFacility(TradePostID)
            Return moTradePost
        End Get
    End Property

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            If BuyOrderID = -1 Then
                sSQL = "INSERT INTO tblBuyOrder (TradePostID, SpecificID, TradeTypeID, AcceptedByID, AcceptedOn, BuyOrderState, Deadline, " & _
                  "Escrow, Quantity, PaymentAmt, BuyOrderType) VALUES (" & Me.TradePostID & ", " & lSpecificID & ", " & iTradeTypeID & ", " & _
                  lAcceptedByID & ", " & lAcceptedOn & ", " & yBuyOrderState & ", " & lDeadline & ", " & blEscrow & ", " & _
                  blQuantity & ", " & blPaymentAmt & ", " & yBuyOrderType & ")"
            Else
                sSQL = "UPDATE tblBuyOrder SET TradePostID = " & Me.TradePostID & ", SpecificID = " & lSpecificID & ", TradeTypeID = " & iTradeTypeID & _
                  ", AcceptedByID = " & lAcceptedByID & ", AcceptedOn = " & lAcceptedOn & ", BuyOrderState = " & yBuyOrderState & _
                  ", Deadline = " & lDeadline & ", Escrow = " & blEscrow & ", Quantity = " & blQuantity & ", PaymentAmt = " & _
                  blPaymentAmt & ", BuyOrderType = " & yBuyOrderType & " WHERE BuyOrderID = " & BuyOrderID
            End If

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If BuyOrderID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(BuyOrderID) FROM tblBuyOrder WHERE TradePostID = " & Me.TradePostID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    BuyOrderID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            oComm = Nothing

            'Save the buy order properties
            sSQL = "DELETE FROM tblBuyOrderProperty WHERE BuyOrderID = " & BuyOrderID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm = Nothing

            For X As Int32 = 0 To mlPropertyUB
                With muProperties(X)
                    sSQL = "INSERT INTO tblBuyOrderProperty (BuyOrderID, PropertyID, PropertyValue, CompareType) VALUES (" & _
                      BuyOrderID & ", " & .lPropertyID & ", " & .lPropertyValue & ", " & .yCompareType & ")"
                End With
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving Buy Order Property!")
                End If
                oComm = Nothing
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object BuyOrder. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub AddBOProp(ByVal lPropertyID As Int32, ByVal lValue As Int32, ByVal yCompare As Byte)
        mlPropertyUB += 1
        ReDim Preserve muProperties(mlPropertyUB)
        With muProperties(mlPropertyUB)
            .lPropertyID = lPropertyID
            .lPropertyValue = lValue
            .yCompareType = yCompare
        End With
    End Sub

    ''' <summary>
    ''' Thread safely checks and alters the Acceptor from the param passed returning true if the acceptor is accepted
    ''' </summary>
    ''' <param name="lAcceptorID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AcceptMe(ByVal lAcceptorID As Int32) As Boolean
        Dim bResult As Boolean = False
        Dim lAcceptedDate As Int32 = GetDateAsNumber(Now)

        SyncLock Me
            If Me.lAcceptedByID = -1 Then
                lAcceptedByID = lAcceptorID
                lAcceptedOn = lAcceptedDate
                bResult = True
            End If
        End SyncLock
        Return bResult
    End Function

    Public Function ValidateData() As Boolean
        If mlPropertyUB = -1 Then Return False

        Dim lMin As Int32 = -1
        Dim lMax As Int32 = Int32.MaxValue

        Select Case CType(yBuyOrderType, GalacticTradeSystem.MarketListType)
            Case GalacticTradeSystem.MarketListType.eAlloys, GalacticTradeSystem.MarketListType.eMinerals
                lMin = 1 : lMax = Int32.MaxValue
            Case GalacticTradeSystem.MarketListType.eArmorComponent
                lMin = elBuyOrderPropID.eArmorIntegrity : lMax = elBuyOrderPropID.eArmorBeamResist
            Case GalacticTradeSystem.MarketListType.eBeamPulseComponent
                lMin = elBuyOrderPropID.ePulsePower : lMax = elBuyOrderPropID.ePulseAccuracy
            Case GalacticTradeSystem.MarketListType.eBeamSolidComponent
                lMin = elBuyOrderPropID.eSolidPower : lMax = elBuyOrderPropID.eSolidAccuracy
            Case GalacticTradeSystem.MarketListType.eEngineComponent
                lMin = elBuyOrderPropID.eEngineHull : lMax = elBuyOrderPropID.eEngineThrust
            Case GalacticTradeSystem.MarketListType.eMissileComponent
                lMin = elBuyOrderPropID.eMissilePower : lMax = elBuyOrderPropID.eMissileROF
            Case GalacticTradeSystem.MarketListType.eProjectileComponent
                lMin = elBuyOrderPropID.eProjectilePower : lMax = elBuyOrderPropID.eProjectileROF
            Case GalacticTradeSystem.MarketListType.eRadarComponent
                lMin = elBuyOrderPropID.eRadarHull : lMax = elBuyOrderPropID.eRadarAccuracy
            Case GalacticTradeSystem.MarketListType.eShieldComponent
                lMin = elBuyOrderPropID.eShieldHull : lMax = elBuyOrderPropID.eShieldMaxHP
            Case GalacticTradeSystem.MarketListType.eBuyOrderSpecificMineral
                If muProperties(0).lPropertyValue > 0 Then Return True Else Return False
        End Select

        For X As Int32 = 0 To mlPropertyUB
            If muProperties(X).lPropertyID < lMin OrElse muProperties(X).lPropertyID > lMax Then Return False
        Next X

        Return True
    End Function

    Public Function ValidateCache(ByRef oObject As Epica_GUID) As Boolean
        Try
            If oObject.ObjTypeID = ObjectType.eMineralCache Then
                Dim oCache As MineralCache = CType(oObject, MineralCache)

                If oCache.oMineral Is Nothing Then Return False

                With oCache.oMineral

                    If Me.yBuyOrderType = GalacticTradeSystem.MarketListType.eBuyOrderSpecificMineral Then
                        If mlPropertyUB >= 0 Then
                            If .ObjectID <> muProperties(0).lPropertyValue Then Return False
                        Else : Return False
                        End If
                    Else
                        For X As Int32 = 0 To mlPropertyUB
                            Dim lCacheVal As Int32 = PlayerMineral.GetClientDisplayedPropertyValue(.GetPropertyValue(muProperties(X).lPropertyID))
                            Dim lPropVal As Int32 = muProperties(X).lPropertyValue
                            If CheckCriteria(lCacheVal, lPropVal, muProperties(X).yCompareType) = False Then Return False
                        Next X
                    End If
                End With

                'Ok, check the quantity
                If oCache.Quantity < Me.blQuantity Then Return False
            ElseIf oObject.ObjTypeID = ObjectType.eComponentCache Then
                Dim oCache As ComponentCache = CType(oObject, ComponentCache)
                If oCache.GetComponent Is Nothing Then Return False
                Dim oComponent As Epica_Tech = oCache.GetComponent()

                For X As Int32 = 0 To mlPropertyUB
                    Dim lCacheVal As Int32 = GetComponentPropValue(oComponent, muProperties(X).lPropertyID)
                    If lCacheVal = Int32.MinValue Then Return False
                    Dim lPropVal As Int32 = muProperties(X).lPropertyValue
                    If CheckCriteria(lCacheVal, lPropVal, muProperties(X).yCompareType) = False Then Return False
                Next X

                'finally, check the quantity
                If oCache.Quantity < Me.blQuantity Then Return False
            Else : Return False
            End If

            'If we are here, everything is good
            Return True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "BuyOrder.ValidateCache: " & ex.Message)
            Return False
        End Try
    End Function
    Private Shared Function CheckCriteria(ByVal lCacheVal As Int32, ByVal lPropVal As Int32, ByVal yCompareType As Byte) As Boolean
        Select Case CType(yCompareType, BuyOrderCompareTypes)
            Case BuyOrderCompareTypes.eEqualTo
                Return lCacheVal = lPropVal
            Case BuyOrderCompareTypes.eGreaterThan
                Return lCacheVal > lPropVal
            Case BuyOrderCompareTypes.eGreaterThanEqualTo
                Return lCacheVal >= lPropVal
            Case BuyOrderCompareTypes.eLessThan
                Return lCacheVal < lPropVal
            Case BuyOrderCompareTypes.eLessThanEqualTo
                Return lCacheVal <= lPropVal
            Case BuyOrderCompareTypes.eNotEqualTo
                Return lCacheVal <> lPropVal
        End Select
        Return False
    End Function
    Private Shared Function GetComponentPropValue(ByRef oComponent As Epica_Tech, ByVal lPropID As Int32) As Int32
        Select Case CType(lPropID, elBuyOrderPropID)
            Case elBuyOrderPropID.eArmorBeamResist
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).yBeamResist
            Case elBuyOrderPropID.eArmorImpactResist
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).yImpactResist
            Case elBuyOrderPropID.eArmorPierceResist
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).yPiercingResist
            Case elBuyOrderPropID.eArmorMagneticResist
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).yMagneticResist
            Case elBuyOrderPropID.eArmorChemicalResist
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).yChemicalResist
            Case elBuyOrderPropID.eArmorBurnResist
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).yBurnResist
            Case elBuyOrderPropID.eArmorRadarResist
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).yRadarResist
            Case elBuyOrderPropID.eArmorHullUsagePerPlate
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).lHullUsagePerPlate
            Case elBuyOrderPropID.eArmorHPPerPlate
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).lHPPerPlate
            Case elBuyOrderPropID.eArmorIntegrity
                If oComponent.ObjTypeID <> ObjectType.eArmorTech Then Return Int32.MinValue
                Return CType(oComponent, ArmorTech).lDisplayedIntegrity
            Case elBuyOrderPropID.ePulseAccuracy
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).GetAccuracy()
            Case elBuyOrderPropID.ePulseAOE
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).ScatterRadius
            Case elBuyOrderPropID.ePulseMaxDmg
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).GetMaxDamage()
            Case elBuyOrderPropID.ePulseMinDmg
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).GetMinDamage
            Case elBuyOrderPropID.ePulseRange
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).MaxRange
            Case elBuyOrderPropID.ePulseROF
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).ROF
            Case elBuyOrderPropID.ePulseHull
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).HullRequired
            Case elBuyOrderPropID.ePulsePower
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyPulse Then Return Int32.MinValue
                Return CType(oComponent, PulseWeaponTech).PowerRequired
            Case elBuyOrderPropID.eSolidAccuracy
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Return CType(oComponent, BeamWeaponTech).Accuracy
            Case elBuyOrderPropID.eSolidRange
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Return CType(oComponent, BeamWeaponTech).MaxRange
            Case elBuyOrderPropID.eSolidROF
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Return CType(oComponent, BeamWeaponTech).ROF
            Case elBuyOrderPropID.eSolidMaxDmg
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Dim oWpnDef As WeaponDef = CType(oComponent, BeamWeaponTech).GetWeaponDefResult()
                With oWpnDef
                    Return .BeamMaxDmg + .ChemicalMaxDmg + .ECMMaxDmg + .FlameMaxDmg + .ImpactMaxDmg + .PiercingMaxDmg
                End With
            Case elBuyOrderPropID.eSolidMinDmg
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Dim oWpnDef As WeaponDef = CType(oComponent, BeamWeaponTech).GetWeaponDefResult()
                With oWpnDef
                    Return .BeamMinDmg + .ChemicalMinDmg + .ECMMinDmg + .FlameMinDmg + .ImpactMinDmg + .PiercingMinDmg
                End With
            Case elBuyOrderPropID.eSolidDmgType
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Return CType(oComponent, BeamWeaponTech).yDmgType
            Case elBuyOrderPropID.eSolidHull
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Return CType(oComponent, BeamWeaponTech).HullRequired
            Case elBuyOrderPropID.eSolidPower
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then Return Int32.MinValue
                Return CType(oComponent, BeamWeaponTech).PowerRequired
            Case elBuyOrderPropID.eEngineThrust
                If oComponent.ObjTypeID <> ObjectType.eEngineTech Then Return Int32.MinValue
                Return CType(oComponent, EngineTech).Thrust
            Case elBuyOrderPropID.eEngineManeuver
                If oComponent.ObjTypeID <> ObjectType.eEngineTech Then Return Int32.MinValue
                Return CType(oComponent, EngineTech).Maneuver
            Case elBuyOrderPropID.eEngineMaxSpeed
                If oComponent.ObjTypeID <> ObjectType.eEngineTech Then Return Int32.MinValue
                Return CType(oComponent, EngineTech).Speed
            Case elBuyOrderPropID.eEnginePowerProd
                If oComponent.ObjTypeID <> ObjectType.eEngineTech Then Return Int32.MinValue
                Return CType(oComponent, EngineTech).PowerProd
            Case elBuyOrderPropID.eEngineHull
                If oComponent.ObjTypeID <> ObjectType.eEngineTech Then Return Int32.MinValue
                Return CType(oComponent, EngineTech).HullRequired
            Case elBuyOrderPropID.eMissileROF
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).ROF
            Case elBuyOrderPropID.eMissileMaxDmg
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).MaxDmg
            Case elBuyOrderPropID.eMissileMissileHullSize
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).MissileHullSize
            Case elBuyOrderPropID.eMissileMaxSpeed
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).MaxSpeed
            Case elBuyOrderPropID.eMissileManeuver
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).Maneuver
            Case elBuyOrderPropID.eMissileFlightTime
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).FlightTime
            Case elBuyOrderPropID.eMissileAccuracy
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).HomingAccuracy
            Case elBuyOrderPropID.eMissilePayload
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).PayloadType
            Case elBuyOrderPropID.eMissileAOE
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).ExplosionRadius
            Case elBuyOrderPropID.eMissileStructHP
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).StructureHP
            Case elBuyOrderPropID.eMissileHull
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).HullRequired
            Case elBuyOrderPropID.eMissilePower
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eMissile Then Return Int32.MinValue
                Return CType(oComponent, MissileWeaponTech).PowerRequired
            Case elBuyOrderPropID.eProjectileROF
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return CType(oComponent, ProjectileWeaponTech).ROF
            Case elBuyOrderPropID.eProjectileMaxRng
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return CType(oComponent, ProjectileWeaponTech).MaxRange
            Case elBuyOrderPropID.eProjectilePayload
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return CType(oComponent, ProjectileWeaponTech).PayloadType
            Case elBuyOrderPropID.eProjectileAOE
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return CType(oComponent, ProjectileWeaponTech).ExplosionRadius
            Case elBuyOrderPropID.eProjectileAmmoSize
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return 0 'CInt(CType(oComponent, ProjectileWeaponTech).GetAmmoSize)
            Case elBuyOrderPropID.eProjectileMaxDmg
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Dim oWpnDef As WeaponDef = CType(oComponent, ProjectileWeaponTech).GetWeaponDefResult()
                With oWpnDef
                    Return .BeamMaxDmg + .ChemicalMaxDmg + .ECMMaxDmg + .FlameMaxDmg + .ImpactMaxDmg + .PiercingMaxDmg
                End With
            Case elBuyOrderPropID.eProjectileMinDmg
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Dim oWpnDef As WeaponDef = CType(oComponent, ProjectileWeaponTech).GetWeaponDefResult()
                With oWpnDef
                    Return .BeamMinDmg + .ChemicalMinDmg + .ECMMinDmg + .FlameMinDmg + .ImpactMinDmg + .PiercingMinDmg
                End With
            Case elBuyOrderPropID.eProjectilePierceRatio
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return CType(oComponent, ProjectileWeaponTech).PierceRatio
            Case elBuyOrderPropID.eProjectileHull
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return CType(oComponent, ProjectileWeaponTech).HullRequired
            Case elBuyOrderPropID.eProjectilePower
                If oComponent.ObjTypeID <> ObjectType.eWeaponTech Then Return Int32.MinValue
                If CType(oComponent, BaseWeaponTech).WeaponClassTypeID <> WeaponClassType.eProjectile Then Return Int32.MinValue
                Return CType(oComponent, ProjectileWeaponTech).PowerRequired
            Case elBuyOrderPropID.eRadarAccuracy
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).WeaponAcc
            Case elBuyOrderPropID.eRadarScanRes
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).ScanResolution
            Case elBuyOrderPropID.eRadarOptRng
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).OptimumRange
            Case elBuyOrderPropID.eRadarMaxRng
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).MaximumRange
            Case elBuyOrderPropID.eRadarDisRes
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).DisruptionResistance
            Case elBuyOrderPropID.eRadarJamImm
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).JamImmunity
            Case elBuyOrderPropID.eRadarJamStr
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).JamStrength
            Case elBuyOrderPropID.eRadarJamTargets
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).JamTargets
            Case elBuyOrderPropID.eRadarJamEffect
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).JamEffect
            Case elBuyOrderPropID.eRadarPower
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).PowerRequired
            Case elBuyOrderPropID.eRadarHull
                If oComponent.ObjTypeID <> ObjectType.eRadarTech Then Return Int32.MinValue
                Return CType(oComponent, RadarTech).HullRequired
            Case elBuyOrderPropID.eShieldMaxHP
                If oComponent.ObjTypeID <> ObjectType.eShieldTech Then Return Int32.MinValue
                Return CType(oComponent, ShieldTech).MaxHitPoints
            Case elBuyOrderPropID.eShieldRechargeRate
                If oComponent.ObjTypeID <> ObjectType.eShieldTech Then Return Int32.MinValue
                Return CType(oComponent, ShieldTech).RechargeRate
            Case elBuyOrderPropID.eShieldRechargeFreq
                If oComponent.ObjTypeID <> ObjectType.eShieldTech Then Return Int32.MinValue
                Return CType(oComponent, ShieldTech).RechargeFreq
            Case elBuyOrderPropID.eShieldProjectionHullSize
                If oComponent.ObjTypeID <> ObjectType.eShieldTech Then Return Int32.MinValue
                Return CType(oComponent, ShieldTech).lProjectionHullSize
            Case elBuyOrderPropID.eShieldPower
                If oComponent.ObjTypeID <> ObjectType.eShieldTech Then Return Int32.MinValue
                Return CType(oComponent, ShieldTech).PowerRequired
            Case elBuyOrderPropID.eShieldHull
                If oComponent.ObjTypeID <> ObjectType.eShieldTech Then Return Int32.MinValue
                Return CType(oComponent, ShieldTech).HullRequired
        End Select

        Return Int32.MinValue
    End Function

    Public Sub DeleteMe()
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            sSQL = "DELETE FROM tblbuyorder WHERE buyorderid = " & Me.BuyOrderID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to DELETE object BUyOrder. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
    End Sub
End Class
