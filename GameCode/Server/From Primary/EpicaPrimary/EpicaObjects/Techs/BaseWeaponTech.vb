Public MustInherit Class BaseWeaponTech
    Inherits Epica_Tech

    Public Const BASE_WEAPON_MSG_HEADER_LENGTH As Int32 = BASE_OBJ_STRING_SIZE + 31

    Public MustOverride Function GetWeaponDefResult() As WeaponDef

	Protected MustOverride Function GetSaveWeaponText() As String

    Public WeaponName(19) As Byte
    Public WeaponClassTypeID As WeaponClassType
    Public WeaponTypeID As WeaponType

    Public PowerRequired As Int32
    Public HullRequired As Int32

    Public yHullTypeID As Byte

    Protected mySendString() As Byte

    ''' <summary>
    ''' Creates one of the children classes (beam, etc...) from a WeaponClassType
    ''' </summary>
    ''' <param name="yWeaponClassType"> the class type to make </param>
    ''' <returns> A BaseWeaponTech class instantiated  </returns>
    ''' <remarks></remarks>
    Public Shared Function CreateWeaponClass(ByVal yWeaponClassType As Byte) As BaseWeaponTech
        Select Case CType(yWeaponClassType, WeaponClassType)
			Case EpicaPrimary.WeaponClassType.eBomb
				Return New BombWeaponTech
			Case EpicaPrimary.WeaponClassType.eEnergyBeam
				Return New BeamWeaponTech
            Case EpicaPrimary.WeaponClassType.eMine
            Case EpicaPrimary.WeaponClassType.eMissile
                Return New MissileWeaponTech
            Case EpicaPrimary.WeaponClassType.eProjectile
                Return New ProjectileWeaponTech
            Case WeaponClassType.eEnergyPulse
                Return New PulseWeaponTech
        End Select

        Return Nothing
    End Function

    ''' <summary>
    ''' Creates one of the children classes (beam, etc...) from a DataReader object containing the recordset
    ''' </summary>
    ''' <param name="oData"></param>
    ''' <returns> the method filled in with the data from the datareader </returns>
    ''' <remarks></remarks>
    Public Shared Function CreateFromDataReader(ByRef oData As OleDb.OleDbDataReader) As BaseWeaponTech
        'First, get the WeaponClassType
        Dim yWCT As Byte = CByte(oData("WeaponClassType"))
        Dim oReturn As BaseWeaponTech = CreateWeaponClass(yWCT)

        oReturn.ErrorReasonCode = CByte(oData("ErrorReasonCode"))
        oReturn.MajorDesignFlaw = CByte(oData("MajorDesignFlaw"))
		oReturn.PopIntel = CInt(oData("PopIntel"))
        oReturn.yArchived = CByte(oData("bArchived"))
        With oReturn
            .lSpecifiedColonists = CInt(oData("SpecifiedColonist"))
            .lSpecifiedEnlisted = CInt(oData("SpecifiedEnlisted"))
            .lSpecifiedHull = CInt(oData("SpecifiedHull"))
            .lSpecifiedMin1 = CInt(oData("SpecifiedMin1"))
            .lSpecifiedMin2 = CInt(oData("SpecifiedMin2"))
            .lSpecifiedMin3 = CInt(oData("SpecifiedMin3"))
            .lSpecifiedMin4 = CInt(oData("SpecifiedMin4"))
            .lSpecifiedMin5 = CInt(oData("SpecifiedMin5"))
            .lSpecifiedMin6 = CInt(oData("SpecifiedMin6"))
            .lSpecifiedOfficers = CInt(oData("SpecifiedOfficer"))
            .lSpecifiedPower = CInt(oData("SpecifiedPower"))
            .blSpecifiedProdCost = CLng(oData("SpecifiedProdCost"))
            .blSpecifiedProdTime = CLng(oData("SpecifiedProdTime"))
            .blSpecifiedResCost = CLng(oData("SpecifiedResCost"))
            .blSpecifiedResTime = CLng(oData("SpecifiedResTime"))
            .yHullTypeID = CByte(oData("HullTypeID"))
            '.lVersionNum = CInt(oData("VersionNumber"))
            If oData("VersionNumber") Is DBNull.Value Then
                .lVersionNum = Epica_Tech.TechVersionNum
            Else
                .lVersionNum = CInt(oData("VersionNumber"))
            End If
        End With

        Dim lTempID As Int32 = CInt(oData("OwnerID"))
        If lTempID = 0 Then oReturn.Owner = goInitialPlayer Else oReturn.Owner = GetEpicaPlayer(lTempID)

        Select Case CType(yWCT, WeaponClassType)
            Case EpicaPrimary.WeaponClassType.eEnergyBeam
                With CType(oReturn, BeamWeaponTech)
                    .ComponentDevelopmentPhase = CType(CInt(oData("ResPhase")), Epica_Tech.eComponentDevelopmentPhase)
                    .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                    .HullRequired = CInt(oData("HullRequired"))
                    .ObjectID = CInt(oData("WeaponID"))
                    .ObjTypeID = ObjectType.eWeaponTech 
                    .PowerRequired = CInt(oData("PowerRequired"))
                    .RandomSeed = CSng(oData("RandomSeed"))
                    .ResearchAttempts = CInt(oData("ResearchAttempts"))
                    .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                    .WeaponName = StringToBytes(CStr(oData("WeaponName")))
                    .WeaponClassTypeID = CType(CByte(oData("WeaponClassType")), WeaponClassType)
                    .WeaponTypeID = CType(CByte(oData("WeaponTypeID")), WeaponType)

                    .MaxDamage = CInt(oData("MaxDmg"))
                    .MaxRange = CShort(oData("OptimumRange"))
                    .ROF = CShort(oData("ROF"))
                    .Accuracy = CByte(oData("Accuracy"))
                    .yDmgType = CByte(oData("PayloadType"))

                    .CoilID = CInt(oData("Mineral1ID"))
                    .CouplerID = CInt(oData("Mineral2ID"))
                    .CasingID = CInt(oData("Mineral3ID"))
                    .FocuserID = CInt(oData("Mineral4ID"))
                    .MediumID = CInt(oData("Mineral5ID"))
                End With
            Case WeaponClassType.eMissile
                With CType(oReturn, MissileWeaponTech)
                    .ComponentDevelopmentPhase = CType(CInt(oData("ResPhase")), Epica_Tech.eComponentDevelopmentPhase)
                    .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                    .HullRequired = CInt(oData("HullRequired"))
                    .ObjectID = CInt(oData("WeaponID"))
                    .ObjTypeID = ObjectType.eWeaponTech
                    .PowerRequired = CInt(oData("PowerRequired"))
                    .RandomSeed = CSng(oData("RandomSeed"))
                    .ResearchAttempts = CInt(oData("ResearchAttempts"))
                    .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                    .WeaponName = StringToBytes(CStr(oData("WeaponName")))
                    .WeaponClassTypeID = CType(CByte(oData("WeaponClassType")), WeaponClassType)
                    .WeaponTypeID = CType(CByte(oData("WeaponTypeID")), WeaponType)

                    .ExplosionRadius = CByte(oData("ExplosionRadius"))
                    .FlightTime = CShort(oData("FlightTime"))
                    .HomingAccuracy = CByte(oData("Accuracy"))
                    .Maneuver = CByte(oData("Maneuver"))
                    .MaxDmg = CInt(oData("MaxDmg"))
                    .MaxSpeed = CByte(oData("MaxSpeed"))
                    .MissileHullSize = CInt(oData("ShotHullSize"))
                    .PayloadType = CByte(oData("PayloadType"))
                    .ROF = CShort(oData("ROF"))
                    .StructureHP = CInt(oData("StructureHP"))

                    '.blAmmoCostCredits = CLng(oData("AmmoCostCredits"))
                    '.blAmmoCostPoints = CLng(oData("AmmoCostPoints"))
                    '.fAmmoMin1Cost = CSng(oData("AmmoMin1Cost"))
                    '.fAmmoMin2Cost = CSng(oData("AmmoMin2Cost"))
                    '.fAmmoMin3Cost = CSng(oData("AmmoMin3Cost"))
                    '.fAmmoMin4Cost = CSng(oData("AmmoMin4Cost"))
                    '.fAmmoMin5Cost = CSng(oData("AmmoMin5Cost"))

                    .lBodyMineralID = CInt(oData("Mineral1ID"))
                    .lNoseMineralID = CInt(oData("Mineral2ID"))
                    .lFlapsMineralID = CInt(oData("Mineral3ID"))
                    .lFuelMineralID = CInt(oData("Mineral4ID"))
                    .lPayloadMineralID = CInt(oData("Mineral5ID"))
                End With
            Case WeaponClassType.eProjectile
                With CType(oReturn, ProjectileWeaponTech)
                    .ComponentDevelopmentPhase = CType(CInt(oData("ResPhase")), Epica_Tech.eComponentDevelopmentPhase)
                    .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                    .HullRequired = CInt(oData("HullRequired"))
                    .ObjectID = CInt(oData("WeaponID"))
                    .ObjTypeID = ObjectType.eWeaponTech
                    .PowerRequired = CInt(oData("PowerRequired"))
                    .RandomSeed = CSng(oData("RandomSeed"))
                    .ResearchAttempts = CInt(oData("ResearchAttempts"))
                    .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                    .WeaponName = StringToBytes(CStr(oData("WeaponName")))
                    .WeaponClassTypeID = CType(CByte(oData("WeaponClassType")), WeaponClassType)
                    .WeaponTypeID = CType(CByte(oData("WeaponTypeID")), WeaponType)

                    .ProjectionType = CByte(oData("MaxSpeed"))
                    .CartridgeSize = CInt(oData("ShotHullSize"))
                    .PierceRatio = CByte(oData("PierceRating"))
                    .ROF = CShort(oData("ROF"))
                    .MaxRange = CShort(oData("OptimumRange"))
                    .PayloadType = CByte(oData("PayloadType"))
                    .ExplosionRadius = CByte(oData("ExplosionRadius"))
                    '.mfAmmoSize = CSng(oData("ProjectileHullSize"))
                    .mfPayload1PotentialImpact = CSng(oData("Payload1PotentialImpact"))
                    .mfPayload1PotentialPierce = CSng(oData("Payload1PotentialPierce"))
                    .mfPayload2Potential = CSng(oData("Payload2Potential"))

                    '.blAmmoCostCredits = CLng(oData("AmmoCostCredits"))
                    '.blAmmoCostPoints = CLng(oData("AmmoCostPoints"))
                    '.fAmmoMin1Cost = CSng(oData("AmmoMin1Cost"))
                    '.fAmmoMin2Cost = CSng(oData("AmmoMin2Cost"))
                    '.fAmmoMin3Cost = CSng(oData("AmmoMin3Cost"))
                    '.fAmmoMin4Cost = CSng(oData("AmmoMin4Cost"))
                    '.fAmmoMin5Cost = CSng(oData("AmmoMin5Cost"))

                    .BarrelMineralID = CInt(oData("Mineral1ID"))
                    .CasingMineralID = CInt(oData("Mineral2ID"))
                    .Payload1MineralID = CInt(oData("Mineral3ID"))
                    .Payload2MineralID = CInt(oData("Mineral4ID"))
                    .ProjectionMineralID = CInt(oData("Mineral5ID"))
                End With
            Case WeaponClassType.eEnergyPulse
                With CType(oReturn, PulseWeaponTech)
                    .ComponentDevelopmentPhase = CType(CInt(oData("ResPhase")), Epica_Tech.eComponentDevelopmentPhase)
                    .CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
                    .HullRequired = CInt(oData("HullRequired"))
                    .ObjectID = CInt(oData("WeaponID"))
                    .ObjTypeID = ObjectType.eWeaponTech
                    .PowerRequired = CInt(oData("PowerRequired"))
                    .RandomSeed = CSng(oData("RandomSeed"))
                    .ResearchAttempts = CInt(oData("ResearchAttempts"))
                    .SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
                    .WeaponName = StringToBytes(CStr(oData("WeaponName")))
                    .WeaponClassTypeID = CType(CByte(oData("WeaponClassType")), WeaponClassType)
                    .WeaponTypeID = CType(CByte(oData("WeaponTypeID")), WeaponType)

                    .mfPulseDegradation = CSng(oData("ProjectileHullSize"))

                    .InputEnergy = CInt(oData("ShotHullSize"))
                    .CompressFactor = CSng(oData("Payload2Potential"))
                    .MaxRange = CByte(oData("OptimumRange"))
                    .ROF = CShort(oData("ROF"))
                    .ScatterRadius = CByte(oData("ExplosionRadius"))

                    .CoilMatID = CInt(oData("Mineral1ID"))
                    .AcceleratorMatID = CInt(oData("Mineral2ID"))
                    .CasingMatID = CInt(oData("Mineral3ID"))
                    .FocuserMatID = CInt(oData("Mineral4ID"))
                    .CompressChmbrMatID = CInt(oData("Mineral5ID"))
				End With
			Case WeaponClassType.eBomb
				With CType(oReturn, BombWeaponTech)
					.ComponentDevelopmentPhase = CType(CInt(oData("ResPhase")), Epica_Tech.eComponentDevelopmentPhase)
					.CurrentSuccessChance = CInt(oData("CurrentSuccessChance"))
					.HullRequired = CInt(oData("HullRequired"))
					.ObjectID = CInt(oData("WeaponID"))
					.ObjTypeID = ObjectType.eWeaponTech
					.PowerRequired = CInt(oData("PowerRequired"))
					.RandomSeed = CSng(oData("RandomSeed"))
					.ResearchAttempts = CInt(oData("ResearchAttempts"))
					.SuccessChanceIncrement = CInt(oData("SuccessChanceIncrement"))
					.WeaponName = StringToBytes(CStr(oData("WeaponName")))
					.WeaponClassTypeID = CType(CByte(oData("WeaponClassType")), WeaponClassType)
					.WeaponTypeID = CType(CByte(oData("WeaponTypeID")), WeaponType)

					.lPayloadSize = CInt(oData("ShotHullSize"))
					.yRange = CByte(oData("OptimumRange"))
					.yPayloadType = CByte(oData("PayloadType"))
					.yAOE = CByte(oData("ExplosionRadius"))
					.yGuidance = CByte(oData("Accuracy"))
					.iROF = CShort(oData("ROF"))

					.lPayloadMatID = CInt(oData("Mineral1ID"))
					.lGuidanceMatID = CInt(oData("Mineral2ID"))
					.lCasingID = CInt(oData("Mineral3ID"))
				End With
		End Select

        Return oReturn
    End Function

    ''' <summary>
    ''' Returns the position to append any extra data once this is done
    ''' DOES NOT REDIM THE ARRAY!!!
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function FillBaseWeaponMsgHdr() As Int32
        Dim lPos As Int32 = 0
        MyBase.GetBaseObjAsString(False).CopyTo(mySendString, lPos) : lPos += BASE_OBJ_STRING_SIZE
        WeaponName.CopyTo(mySendString, lPos) : lPos += 20
        mySendString(lPos) = CByte(WeaponClassTypeID) : lPos += 1
        mySendString(lPos) = CByte(WeaponTypeID) : lPos += 1
        System.BitConverter.GetBytes(PowerRequired).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(HullRequired).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = yHullTypeID : lPos += 1
 
        Return lPos
    End Function

    '   Public Function GetAmmoSize() As Single
    '       Select Case WeaponClassTypeID
    '           Case WeaponClassType.eMissile
    '               Return CType(Me, MissileWeaponTech).MissileHullSize
    '           Case WeaponClassType.eProjectile
    '               Return CType(Me, ProjectileWeaponTech).mfAmmoSize
    '           Case Else
    '               Return 0.0F
    '       End Select
    'End Function

	Public Function GetSaveObjectText() As String
		Dim oSB As New System.Text.StringBuilder()
		Dim sTemp As String = GetSaveWeaponText()
		If sTemp = "" Then Return ""
		oSB.AppendLine(sTemp)
		oSB.AppendLine(MyBase.GetFinalizeSaveText())
		Return oSB.ToString
	End Function

    Protected Function CheckDesignImpossibility() As Boolean
        Dim oTB As TechBuilderComputer

        Select Case Me.WeaponClassTypeID
            Case WeaponClassType.eBomb
                Dim oBomb As New BombTechComputer()
                With CType(Me, BombWeaponTech)
                    oBomb.lMineral1ID = .lCasingID
                    oBomb.lMineral2ID = .lGuidanceMatID
                    oBomb.lMineral3ID = .lPayloadMatID
                    oBomb.lMineral4ID = -1
                    oBomb.lMineral5ID = -1
                    oBomb.lMineral6ID = -1

                    oBomb.blAOE = .yAOE
                    oBomb.blGuidance = .yGuidance
                    oBomb.blPayloadSize = .lPayloadSize
                    oBomb.blRange = .yRange
                    oBomb.blROF = .iROF

                    Dim fROF As Single = Math.Max(1, .iROF) / 30.0F
                    Dim lImpactMin As Int32 = 0
                    Dim lImpactMax As Int32 = 0
                    Dim lFlameMin As Int32 = 0
                    Dim lFlameMax As Int32 = 0
                    If .yPayloadType = 0 Then
                        Dim lHalfPayload As Int32 = .lPayloadSize \ 2
                        Dim lQtrPayload As Int32 = lHalfPayload \ 2
                        lImpactMax = lHalfPayload : lImpactMin = lQtrPayload
                        lFlameMax = lHalfPayload : lFlameMin = lQtrPayload
                    End If
                    oBomb.decDPS = CDec((CDec(lImpactMax) + CDec(lFlameMax)) / fROF)
                    oBomb.decDPS = Math.Max(oBomb.decDPS, (oBomb.decDPS + CDec(lImpactMax) + CDec(lFlameMax)) / 2D)
                End With
                oTB = oBomb
            Case WeaponClassType.eEnergyBeam
                Dim oSolid As New SolidTechComputer()
                With CType(Me, BeamWeaponTech)
                    oSolid.lMineral1ID = .CoilID
                    oSolid.lMineral2ID = .CouplerID
                    oSolid.lMineral3ID = .CasingID
                    oSolid.lMineral4ID = .FocuserID
                    oSolid.lMineral5ID = .MediumID
                    oSolid.lMineral6ID = -1

                    oSolid.blAccuracy = .Accuracy
                    oSolid.blMaxDmg = .MaxDamage

                    'Bogus: Do not back out specials.  These bonus's are applied after the design is finalized.
                    'oSolid.blMaxRng = Math.Max(0, .MaxRange - CShort(Owner.oSpecials.ySolidBeamOptRngBonus))
                    'Dim iROF As Int16 = CShort(Math.Min(Int16.MaxValue, Math.Max(1, CInt(.ROF) + CInt(Owner.oSpecials.ySolidBeamROFBonus))))

                    oSolid.blMaxRng = Math.Max(0, .MaxRange)
                    Dim iROF As Int16 = CShort(Math.Min(Int16.MaxValue, Math.Max(1, CInt(.ROF))))
                    Dim fROF As Single = Math.Max(1, iROF) / 30.0F

                    Dim lMinPierce As Int32 = 0
                    Dim lMaxPierce As Int32 = 0
                    Dim lMinBurn As Int32 = 0
                    Dim lMaxBurn As Int32 = 0
                    Dim lMinBeam As Int32 = 0
                    Dim lMaxBeam As Int32 = 0

                    Dim lMaxDmg As Int32 = .MaxDamage
                    'Bogus: Do not back out specials.  These bonus's are applied after the design is finalized.
                    'If Owner.oSpecials.ySolidBeamMaxDmgBonus > 0 Then
                    'Dim fMult As Single = 1.0F + (Owner.oSpecials.ySolidBeamMaxDmgBonus / 100.0F)
                    'fMult = 1.0F / fMult
                    'lMaxDmg = CInt(Math.Ceiling(lMaxDmg * fMult))
                    'End If

                    If .yDmgType = 0 Then
                        'Piercing
                        Dim lVal As Int32 = CInt(lMaxDmg * 0.3F)
                        lMaxPierce = lVal
                        lVal = .MaxDamage - lVal
                        lMaxBeam = lVal
                    Else
                        'thermal
                        Dim lVal As Int32 = CInt(lMaxDmg * 0.3F)
                        lMaxBurn = lVal
                        lVal = .MaxDamage - lVal
                        lMaxBeam = lVal
                    End If
                    lMinBeam = lMaxBeam \ 10
                    oSolid.decDPS = lMaxDmg / CDec(fROF)
                    oSolid.decDPS = Math.Max(oSolid.decDPS, (oSolid.decDPS + lMaxDmg) / 2D)
                End With
                oTB = oSolid
            Case WeaponClassType.eEnergyPulse
                Dim oPulse As New PulseTechComputer()
                With CType(Me, PulseWeaponTech)
                    oPulse.lMineral1ID = .CoilMatID
                    oPulse.lMineral2ID = .AcceleratorMatID
                    oPulse.lMineral3ID = .CasingMatID
                    oPulse.lMineral4ID = .FocuserMatID
                    oPulse.lMineral5ID = .CompressChmbrMatID
                    oPulse.lMineral6ID = -1
                    'Bogus: Do not back out specials.  These bonus's are applied after the design is finalized.
                    'Dim iROF As Int16 = CShort(Math.Min(Int16.MaxValue, Math.Max(1, CInt(.ROF) + CInt(Owner.oSpecials.yPulseROFBonus))))
                    Dim iROF As Int16 = CShort(Math.Min(Int16.MaxValue, Math.Max(1, CInt(.ROF))))
                    Dim fROF As Single = Math.Max(1, iROF) / 30.0F
                    Dim lMaxBeam As Int32 = .GetMaxBeamDamage()
                    Dim lMinBeam As Int32 = .GetMinBeamDamage()
                    Dim lMaxImpact As Int32 = .GetMaxImpactDamage()
                    Dim lMinImpact As Int32 = .GetMinImpactDamage()
                    oPulse.decDPS = (CDec(lMaxBeam) + CDec(lMaxImpact)) / CDec(fROF)
                    oPulse.decDPS = Math.Max(oPulse.decDPS, (oPulse.decDPS + CDec(lMaxBeam) + CDec(lMaxImpact)) / 2D)
                    oPulse.blCompress = CLng(.CompressFactor * 10)
                    oPulse.blInputEnergy = .InputEnergy
                    'oPulse.blMaxRange = Math.Max(1, CInt(.MaxRange) - CInt(Owner.oSpecials.yPulseOptRngBonus))
                    oPulse.blMaxRange = Math.Max(1, CInt(.MaxRange))
                    oPulse.blROF = iROF
                    oPulse.blScatterRadius = .ScatterRadius
                End With
                oTB = oPulse
            Case WeaponClassType.eMissile
                Dim oMissile As New MissileTechComputer()
                With CType(Me, MissileWeaponTech)
                    oMissile.lMineral1ID = .lBodyMineralID
                    oMissile.lMineral2ID = .lNoseMineralID
                    oMissile.lMineral3ID = .lFlapsMineralID
                    oMissile.lMineral4ID = .lFuelMineralID
                    oMissile.lMineral5ID = .lPayloadMineralID
                    oMissile.lMineral6ID = -1

                    Dim fROF As Single = Math.Max(1, .ROF) / 30.0F
                    oMissile.decDPS = CDec(.MaxDmg) / CDec(fROF)
                    oMissile.decDPS = Math.Max(oMissile.decDPS, (oMissile.decDPS + CDec(.MaxDmg)) / 2D)

                    'Bogus: Do not back out specials.  These bonus's are applied after the design is finalized.
                    'oMissile.blExplosionRadius = Math.Max(0, CInt(.ExplosionRadius) - CInt(Owner.oSpecials.yMissileExpRadBonus))
                    'oMissile.blHomingAccuracy = .HomingAccuracy
                    'Dim fMult As Single = 1.0F - (Owner.oSpecials.yMissileHullSizeBonus / 100.0F)
                    'fMult = 1.0F / fMult
                    'oMissile.blHullSize = Math.Max(1, CInt(.MissileHullSize * fMult))
                    'oMissile.blManeuver = Math.Max(1, CInt(.Maneuver) - CInt(Owner.oSpecials.yMissileManeuverBonus))

                    'If Owner.oSpecials.yMissileMaxDmgBonus > 0 Then
                    'fMult = 1.0F + (Owner.oSpecials.yMissileMaxDmgBonus / 100.0F)
                    'fMult = 1.0F / fMult
                    'oMissile.blMaxDamage = CLng(Math.Ceiling(.MaxDmg * fMult))
                    'Else
                    'oMissile.blMaxDamage = .MaxDmg
                    'End If

                    'oMissile.blMaxSpeed = Math.Max(1, CInt(.MaxSpeed) - CInt(Owner.oSpecials.yMissileMaxSpeedBonus))
                    'oMissile.blRange = Math.Max(0, CInt(.FlightTime) - CInt(Owner.oSpecials.yMissileFlightTimeBonus))

                    oMissile.blExplosionRadius = Math.Max(0, CInt(.ExplosionRadius))
                    oMissile.blHomingAccuracy = .HomingAccuracy
                    oMissile.blHullSize = Math.Max(1, CInt(.MissileHullSize))
                    oMissile.blManeuver = Math.Max(1, CInt(.Maneuver))
                    oMissile.blMaxDamage = .MaxDmg
                    oMissile.blMaxSpeed = Math.Max(1, CInt(.MaxSpeed))
                    oMissile.blRange = Math.Max(0, CInt(.FlightTime))

                    oMissile.blROF = .ROF
                    oMissile.blStructHP = .StructureHP
                    oMissile.yPayloadType = .PayloadType
                End With
                oTB = oMissile
            Case WeaponClassType.eProjectile
                Dim oProj As New ProjectileTechComputer()
                With CType(Me, ProjectileWeaponTech)
                    oProj.lMineral1ID = .BarrelMineralID
                    oProj.lMineral2ID = .CasingMineralID
                    oProj.lMineral3ID = .Payload1MineralID
                    oProj.lMineral4ID = .Payload2MineralID
                    oProj.lMineral5ID = .ProjectionMineralID
                    oProj.lMineral6ID = -1

                    Dim fROF As Single = Math.Max(1, .ROF) / 30.0F
                    Dim lHalfCart As Int32 = .CartridgeSize \ 2

                    Dim lMinPierce As Int32 = 0
                    Dim lMinImpact As Int32 = 0
                    Dim lMaxPierce As Int32 = CInt(lHalfCart * (CSng(.PierceRatio) / 100.0F))
                    Dim lMaxImpact As Int32 = CInt(lHalfCart * ((100 - CSng(.PierceRatio)) / 100.0F))
                    Dim lMinPayload As Int32 = 0
                    Dim lMaxPayload As Int32 = lHalfCart

                    Dim lMaxFlame As Int32 = 0
                    Dim lMinFlame As Int32 = 0
                    Dim lMaxChemical As Int32 = 0
                    Dim lMinChemical As Int32 = 0
                    Dim lMaxECM As Int32 = 0
                    Dim lMinECM As Int32 = 0
                    Select Case .PayloadType
                        Case 1
                            lMaxFlame = lMaxPayload : lMinFlame = lMinPayload
                        Case 2
                            lMaxChemical = lMaxPayload : lMinChemical = lMinPayload
                        Case 3
                            lMaxECM = lMaxPayload : lMinECM = lMinPayload
                        Case Else
                            lMaxImpact += lMaxPayload
                            lMinImpact += lMinPayload
                    End Select
                    oProj.decDPS = (CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / CDec(fROF)
                    oProj.decDPS = Math.Max(oProj.decDPS, (oProj.decDPS + CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / 2D)

                    oProj.blCartridgeSize = .CartridgeSize
                    oProj.blExplosionRadius = .ExplosionRadius
                    oProj.blMaxRange = .MaxRange
                    oProj.blPierceRatio = .PierceRatio
                    oProj.blROF = .ROF
                    oProj.yPayloadType = .PayloadType
                    oProj.yProjectionType = .ProjectionType
                End With
                oTB = oProj
            Case Else
                Return False
        End Select

        With Me
            oTB.blLockedProdCost = .blSpecifiedProdCost
            oTB.blLockedProdTime = .blSpecifiedProdTime
            oTB.blLockedResCost = .blSpecifiedResCost
            oTB.blLockedResTime = .blSpecifiedResTime
            oTB.lHullTypeID = .yHullTypeID
            oTB.lLockedColonists = .lSpecifiedColonists
            oTB.lLockedEnlisted = .lSpecifiedEnlisted
            oTB.lLockedHull = .lSpecifiedHull
            oTB.lLockedMin1 = .lSpecifiedMin1
            oTB.lLockedMin2 = .lSpecifiedMin2
            oTB.lLockedMin3 = .lSpecifiedMin3
            oTB.lLockedMin4 = .lSpecifiedMin4
            oTB.lLockedMin5 = .lSpecifiedMin5
            oTB.lLockedMin6 = .lSpecifiedMin6
            oTB.lLockedOfficers = .lSpecifiedOfficers
            oTB.lLockedPower = .lSpecifiedPower
        End With
        Return oTB.IsDesignImpossible(Me.ObjTypeID, Owner)
    End Function
End Class
