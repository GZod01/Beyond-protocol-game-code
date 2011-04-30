Public Class Epica_Entity_Def
    Inherits Epica_GUID

    Public Enum eExtendedFlagValues As Int32
        'Bitwise
        ContainsMicroTech = 1
    End Enum

    Protected Const ml_WEAPON_DEF_LEN As Int32 = 94 '70           'the weapon def itself is 69, but the arcid is 1 byte
    Private Const ml_MSG_BASE_UB As Int32 = 125

    'a unitdef is not owned, its more of a way to help the engine do things faster
    Public oPrototype As Prototype
    Public Q1_MaxHP As Int32
    Public Q2_MaxHP As Int32
    Public Q3_MaxHP As Int32
    Public Q4_MaxHP As Int32
    Public Maneuver As Byte
    Public MaxSpeed As Byte
    Public FuelEfficiency As Byte
    Public Structure_MaxHP As Int32
    Public Hangar_Cap As Int32

    Protected moHangarDoors() As EntityDefHangarDoor
    Public lHangarDoorUB As Int32 = -1

    Public HullSize As Int32
    Public Cargo_Cap As Int32
    Public Fuel_Cap As Int32
    Public Weapon_Acc As Byte
    Public ScanResolution As Byte
    Public OptRadarRange As Byte
    Public MaxRadarRange As Byte
    Public DisruptionResistance As Byte
    Public JamImmunity As Byte
    Public JamStrength As Byte
    Public JamTargets As Byte
    Public JamEffect As Byte
    Public PiercingResist As Byte
    Public ImpactResist As Byte
    Public BeamResist As Byte
    Public ECMResist As Byte
    Public FlameResist As Byte
    Public ChemicalResist As Byte
    Public DetectionResist As Byte
    Public ArmorIntegrity As Byte
    Public Shield_MaxHP As Int32
    Public ShieldRecharge As Int32
    Public ShieldRechargeFreq As Int32

    Public lSideCrits(3) As Int32

    'Public ModelID As Int16
    Private miModelID As Int16
    Public Property ModelID() As Int16
        Get
            Return miModelID
        End Get
        Set(ByVal value As Int16)
            miModelID = value
            Dim oModelDef As ModelDef = goModelDefs.GetModelDef(miModelID)
            If oModelDef Is Nothing = False Then
                lHullSpecialTraitID = CType(oModelDef.lSpecialTraitID, elModelSpecialTrait)
            End If
        End Set
    End Property
    Public lHullSpecialTraitID As elModelSpecialTrait = elModelSpecialTrait.NoSpecialTrait  'NOT SAVED!

    Public DefName(19) As Byte

    Public yChassisType As Byte

    Public OwnerID As Int32

    Public WeaponDefs() As UnitWeaponDef
    Public WeaponDefUB As Int32 = -1

    Public ProductionRequirements As ProductionCost
    Public ProductionTypeID As Byte
    Public RequiredProductionTypeID As Byte

    Public yFXColors As Byte

    Public lExtendedFlags As Int32 = 0

    Public EntityDefMinerals() As EntityDefMineral
    Public lEntityDefMineralUB As Int32 = -1
    Public Sub AddEntityDefMineral(ByVal lMineralID As Int32, ByVal lQty As Int32)
        For X As Int32 = 0 To lEntityDefMineralUB
            If EntityDefMinerals(X).oMineral.ObjectID = lMineralID Then
                EntityDefMinerals(X).lQuantity += lQty
                Return
            End If
        Next X

        lEntityDefMineralUB += 1
        ReDim Preserve EntityDefMinerals(lEntityDefMineralUB)
        EntityDefMinerals(lEntityDefMineralUB) = New EntityDefMineral
        With EntityDefMinerals(lEntityDefMineralUB)
            .oEntityDef = Me
            .oMineral = GetEpicaMineral(lMineralID)
            .iEntityDefTypeID = Me.ObjTypeID
            .lQuantity = lQty
        End With
    End Sub

    Protected mySendString() As Byte

    'Private mlCombatRating As Int32 = Int32.MinValue
    'Public ReadOnly Property CombatRating() As Int32
    '    Get
    '        Try
    '            If mlCombatRating = Int32.MinValue Then
    '                Dim lTotalArmor As Int32 = Q1_MaxHP + Q2_MaxHP + Q3_MaxHP + Q4_MaxHP
    '                Dim lResists As Int32 = CInt(PiercingResist) + CInt(ImpactResist) + CInt(BeamResist) + CInt(ECMResist) + CInt(FlameResist) + CInt(ChemicalResist)
    '                lResists \= 6
    '                Dim fShieldRegenPerSec As Single = 0
    '                If Me.ShieldRechargeFreq <> 0 Then fShieldRegenPerSec = Me.ShieldRecharge * (30.0F / Me.ShieldRechargeFreq)
    '                Dim lWpnStrength As Int32 = 0
    '                For X As Int32 = 0 To WeaponDefUB
    '                    lWpnStrength += WeaponDefs(X).oWeaponDef.lFirePowerRating
    '                Next X
    '                mlCombatRating = (lTotalArmor \ ((lResists \ 100) + 1))
    '                mlCombatRating += CInt(Shield_MaxHP * fShieldRegenPerSec)
    '                mlCombatRating += CInt(MaxSpeed * (2I * Maneuver))
    '                mlCombatRating += (lWpnStrength * 5)
    '            End If
    '        Catch
    '            mlCombatRating = 1
    '        End Try
    '        Return mlCombatRating
    '    End Get
    'End Property

    Private mlNewCombatRating As Int32 = Int32.MinValue
    Public ReadOnly Property CombatRating() As Int32
        Get
            Try
                '> ((TotalArmorHP * Integrity_As_Percentage) / ((AverageResist + 1) / 100) + Shield_MaxHP) ^ (1 +
                '> (ShieldRegenPerSec / Shield_MaxHP)) + Speed * 2 * Maneuver + WeaponStrength

                If mlNewCombatRating = Int32.MinValue Then

                    If WeaponDefUB = -1 Then
                        mlNewCombatRating = 0
                    Else
                        Dim lTotalArmor As Int32 = Q1_MaxHP + Q2_MaxHP + Q3_MaxHP + Q4_MaxHP
                        'lTotalArmor *= Me.ArmorIntegrity
                        Dim lResists As Int32 = CInt(PiercingResist) + CInt(ImpactResist) + CInt(BeamResist) + CInt(ECMResist) + CInt(FlameResist) + CInt(ChemicalResist)
                        lResists \= 6
                        Dim fShieldRegenPerSec As Single = 0
                        If Me.ShieldRechargeFreq <> 0 Then fShieldRegenPerSec = Me.ShieldRecharge * (30.0F / Me.ShieldRechargeFreq)
                        Dim lWpnStrength As Int32 = 0
                        For X As Int32 = 0 To WeaponDefUB
                            If WeaponDefs(X).yArcID = UnitArcs.eAllArcs Then
                                lWpnStrength += WeaponDefs(X).oWeaponDef.lFirePowerRating * 2
                            Else
                                lWpnStrength += WeaponDefs(X).oWeaponDef.lFirePowerRating
                            End If
                        Next X

                        Dim fTemp As Single = ((lTotalArmor * (1.0F - (Me.ArmorIntegrity / 100.0F))) / ((lResists + 1) / 100.0F))
                        If Shield_MaxHP > 0 Then
                            fTemp += Shield_MaxHP
                            fTemp = CSng(Math.Pow(fTemp, 1 + (fShieldRegenPerSec / Shield_MaxHP)))
                        End If
                        fTemp += (MaxSpeed * 2I * Maneuver)
                        fTemp += lWpnStrength

                        mlNewCombatRating = CInt(fTemp)

                    End If
                End If
            Catch
                mlNewCombatRating = 1
            End Try
            Return mlNewCombatRating
        End Get
    End Property
    'Private mlWarpointUpkeepCost As Int32 = Int32.MinValue
    'Public ReadOnly Property WarpointUpkeep() As Int32
    '    Get
    '        If mlWarpointUpkeepCost = Int32.MinValue Then
    '            mlWarpointUpkeepCost = Math.Max(1, CombatRating \ 10000)
    '        End If
    '        Return mlWarpointUpkeepCost
    '    End Get
    'End Property

#Region "  Critical Hit Chance Management  "
    'per side, per critical hit type, a chance
    Private Structure CriticalHitChance
        Public ySide As UnitArcs
        Public lCritical As Int32
        Public yChance As Byte
    End Structure
    Private muCriticalHitChances() As CriticalHitChance
    Private mlCriticalHitChanceUB As Int32 = -1
    Private myCriticalHitChanceMsg() As Byte
    Public Sub AddCriticalHitChance(ByVal ySide As UnitArcs, ByVal lCritical As Int32, ByVal yChance As Byte)
        For X As Int32 = 0 To mlCriticalHitChanceUB
            With muCriticalHitChances(X)
                If .lCritical = lCritical AndAlso .ySide = ySide Then
                    .yChance = Math.Max(.yChance, yChance)
                    Return
                End If
            End With
        Next X
        mlCriticalHitChanceUB += 1
        ReDim Preserve muCriticalHitChances(mlCriticalHitChanceUB)
        With muCriticalHitChances(mlCriticalHitChanceUB)
            .lCritical = lCritical
            .yChance = yChance
            .ySide = ySide
        End With
    End Sub
    Public Sub ResortCriticalHitChances()
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1

        If mlCriticalHitChanceUB = -1 Then Return

        For X As Int32 = 0 To mlCriticalHitChanceUB
            Dim lIdx As Int32 = -1
            For Y As Int32 = 0 To lSortedUB
                If muCriticalHitChances(Y).yChance < muCriticalHitChances(X).yChance Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X

        If lSortedUB <> mlCriticalHitChanceUB Then
            LogEvent(LogEventType.CriticalError, "Unable to sort Critical Hit Chances! UB's did not match: " & lSortedUB & ", " & mlCriticalHitChanceUB)
            Return
        End If
        If lSorted Is Nothing Then
            LogEvent(LogEventType.CriticalError, "Unable to sort Critical Hit Chances! Sorted array is nothing!")
            Return
        End If

        Dim uSorted(mlCriticalHitChanceUB) As CriticalHitChance
        For X As Int32 = 0 To mlCriticalHitChanceUB
            uSorted(X) = muCriticalHitChances(lSorted(X))
        Next X
        muCriticalHitChances = uSorted
    End Sub
    Public Function GetCriticalHitChanceMsg() As Byte()
        If mlCriticalHitChanceUB = -1 Then Return Nothing

        If myCriticalHitChanceMsg Is Nothing Then
            ReDim myCriticalHitChanceMsg(11 + ((mlCriticalHitChanceUB + 1) * 6))
            Dim lPos As Int32 = 0

			System.BitConverter.GetBytes(GlobalMessageCode.eEntityDefCriticalHitChances).CopyTo(myCriticalHitChanceMsg, lPos) : lPos += 2
            Me.GetGUIDAsString.CopyTo(myCriticalHitChanceMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(mlCriticalHitChanceUB + 1).CopyTo(myCriticalHitChanceMsg, lPos) : lPos += 4
            For X As Int32 = 0 To mlCriticalHitChanceUB
                With muCriticalHitChances(X)
                    myCriticalHitChanceMsg(lPos) = .ySide : lPos += 1
                    System.BitConverter.GetBytes(.lCritical).CopyTo(myCriticalHitChanceMsg, lPos) : lPos += 4
                    myCriticalHitChanceMsg(lPos) = .yChance : lPos += 1
                End With
            Next X
        End If

        Return myCriticalHitChanceMsg
    End Function
#End Region

    Public Sub AddWeaponDef(ByVal lID As Int32, ByRef oWeaponDef As WeaponDef, ByVal yArcID As UnitArcs, ByVal lEntityStatusGroup As Int32, ByVal lAmmoCap As Int32)
        WeaponDefUB += 1
        ReDim Preserve WeaponDefs(WeaponDefUB)
        WeaponDefs(WeaponDefUB) = New UnitWeaponDef()

        With WeaponDefs(WeaponDefUB)
            .ObjectID = lID
            .ObjTypeID = ObjectType.eUnitWeaponDef
            .oUnitDef = Me
            .oWeaponDef = oWeaponDef
            .yArcID = yArcID
            .mlAmmoCap = lAmmoCap
            .lEntityStatusGroup = lEntityStatusGroup
        End With

    End Sub

    Public Overridable Function GetObjAsString() As Byte()
        Dim X As Int32
        Dim lPos As Int32
        'here we will return the entire object as a string
        If mbStringReady = False Then
            'this equation may be slightly off...
            ReDim mySendString(ml_MSG_BASE_UB + ((ml_WEAPON_DEF_LEN + 12) * (WeaponDefUB + 1)))

            lPos = 0

            GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            System.BitConverter.GetBytes(OwnerID).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q1_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q2_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q3_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Q4_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            mySendString(lPos) = Maneuver : lPos += 1
            mySendString(lPos) = MaxSpeed : lPos += 1
            mySendString(lPos) = FuelEfficiency : lPos += 1
            System.BitConverter.GetBytes(Structure_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Hangar_Cap).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(HullSize).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Cargo_Cap).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Shield_MaxHP).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ShieldRecharge).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ShieldRechargeFreq).CopyTo(mySendString, lPos) : lPos += 4

            'System.BitConverter.GetBytes(Max_Door_Size).CopyTo(mySendString, lPos) : lPos += 4
            'mySendString(lPos) = NumberOfDoors : lPos += 1

			Dim lCrew As Int32 = 1
			If Me.ProductionRequirements Is Nothing = False Then
				lCrew += Me.ProductionRequirements.EnlistedCost + Me.ProductionRequirements.OfficerCost + Me.ProductionRequirements.ColonistCost
			End If
			System.BitConverter.GetBytes(lCrew).CopyTo(mySendString, lPos) : lPos += 4		  'was Fuel_Cap
            mySendString(lPos) = Weapon_Acc : lPos += 1
            mySendString(lPos) = ScanResolution : lPos += 1
            mySendString(lPos) = OptRadarRange : lPos += 1
            mySendString(lPos) = MaxRadarRange : lPos += 1
            mySendString(lPos) = DisruptionResistance : lPos += 1
            mySendString(lPos) = PiercingResist : lPos += 1
            mySendString(lPos) = ImpactResist : lPos += 1
            mySendString(lPos) = BeamResist : lPos += 1
            mySendString(lPos) = ECMResist : lPos += 1
            mySendString(lPos) = FlameResist : lPos += 1
            mySendString(lPos) = ChemicalResist : lPos += 1
            mySendString(lPos) = DetectionResist : lPos += 1
            System.BitConverter.GetBytes(ModelID).CopyTo(mySendString, lPos) : lPos += 2

            DefName.CopyTo(mySendString, lPos) : lPos += 20

            mySendString(lPos) = ProductionTypeID : lPos += 1
            mySendString(lPos) = RequiredProductionTypeID : lPos += 1
            mySendString(lPos) = yChassisType : lPos += 1
            mySendString(lPos) = yFXColors : lPos += 1
            mySendString(lPos) = ArmorIntegrity : lPos += 1
            mySendString(lPos) = JamImmunity : lPos += 1
            mySendString(lPos) = JamStrength : lPos += 1
            mySendString(lPos) = JamTargets : lPos += 1
            mySendString(lPos) = JamEffect : lPos += 1

            'side crits
            For X = 0 To 3
                System.BitConverter.GetBytes(lSideCrits(X)).CopyTo(mySendString, lPos)
                lPos += 4
			Next X

			''================== PUT IN FOR PRE-BALANCE FIXES
			'Dim blTotalDPS As Int64 = 0
			'For X = 0 To WeaponDefUB
			'	Dim blDPSTemp As Int64 = 0
			'	With WeaponDefs(X).oWeaponDef
			'		blDPSTemp += .BeamMaxDmg
			'		blDPSTemp += .ChemicalMaxDmg
			'		blDPSTemp += .ECMMaxDmg
			'		blDPSTemp += .FlameMaxDmg
			'		blDPSTemp += .ImpactMaxDmg
			'		blDPSTemp += .PiercingMaxDmg

			'		Dim fShotsPerSec As Single = 1.0F
			'		If .ROF <> 0 Then fShotsPerSec = 30.0F / .ROF
			'		blDPSTemp = CLng(fShotsPerSec * blDPSTemp)
			'		blTotalDPS += blDPSTemp
			'	End With
			'Next X
			'Dim fDPSExcess As Single = CSng(blTotalDPS / GetMaxDPS(Me.ModelID))
			'If fDPSExcess < 1 Then fDPSExcess = 1.0F 
			''================= END of Balancing Changes =====================


            'Now, we need to include the Unit Def's weapons...
            'next two bytes is the weapondef count
            System.BitConverter.GetBytes(CShort(WeaponDefUB + 1)).CopyTo(mySendString, lPos)
            lPos += 2

            For X = 0 To WeaponDefUB
                'now, we need to append the weapon def's message here...
                'first, append the Arc
                If WeaponDefs(X) Is Nothing OrElse WeaponDefs(X).oWeaponDef Is Nothing Then
                    lPos += 1 + ml_WEAPON_DEF_LEN - 2 + 4 + 4 + 4
                    Continue For
                End If

                mySendString(lPos) = WeaponDefs(X).yArcID
                lPos += 1
                'then, append the rest
				'WeaponDefs(X).oWeaponDef.GetObjAsStringEx(fDPSExcess).CopyTo(mySendString, lPos)
				WeaponDefs(X).oWeaponDef.GetObjAsString().CopyTo(mySendString, lPos)
                lPos += ml_WEAPON_DEF_LEN - 2

                System.BitConverter.GetBytes(WeaponDefs(X).lEntityStatusGroup).CopyTo(mySendString, lPos)
                lPos += 4

                System.BitConverter.GetBytes(WeaponDefs(X).mlAmmoCap).CopyTo(mySendString, lPos)
                lPos += 4

                System.BitConverter.GetBytes(WeaponDefs(X).ObjectID).CopyTo(mySendString, lPos)
                lPos += 4
            Next X

            System.BitConverter.GetBytes(lExtendedFlags).CopyTo(mySendString, lPos) : lPos += 4

            mbStringReady = True
        End If
        Return mySendString
	End Function

    Public Overridable Sub FillFromForwardedAddMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 0

        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q1_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q2_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q3_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q4_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Maneuver = yData(lPos) : lPos += 1
            .MaxSpeed = yData(lPos) : lPos += 1
            .FuelEfficiency = yData(lPos) : lPos += 1
            .Structure_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Hangar_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Cargo_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Shield_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ShieldRecharge = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ShieldRechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lPos += 4   'for crew
            Weapon_Acc = yData(lPos) : lPos += 1
            ScanResolution = yData(lPos) : lPos += 1
            OptRadarRange = yData(lPos) : lPos += 1
            MaxRadarRange = yData(lPos) : lPos += 1
            DisruptionResistance = yData(lPos) : lPos += 1
            PiercingResist = yData(lPos) : lPos += 1
            ImpactResist = yData(lPos) : lPos += 1
            BeamResist = yData(lPos) : lPos += 1
            ECMResist = yData(lPos) : lPos += 1
            FlameResist = yData(lPos) : lPos += 1
            ChemicalResist = yData(lPos) : lPos += 1
            DetectionResist = yData(lPos) : lPos += 1
            ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            ReDim .DefName(19)
            Array.Copy(yData, lPos, .DefName, 0, 20) : lPos += 20
 
            .ProductionTypeID = yData(lPos) : lPos += 1
            .RequiredProductionTypeID = yData(lPos) : lPos += 1
            .yChassisType = yData(lPos) : lPos += 1
            .yFXColors = yData(lPos) : lPos += 1
            .ArmorIntegrity = yData(lPos) : lPos += 1
            .JamImmunity = yData(lPos) : lPos += 1
            .JamStrength = yData(lPos) : lPos += 1
            .JamTargets = yData(lPos) : lPos += 1
            .JamEffect = yData(lPos) : lPos += 1

            For X As Int32 = 0 To 3
                lSideCrits(X) = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
            Next X

            WeaponDefUB = System.BitConverter.ToInt16(yData, lPos) - 1 : lPos += 2
            ReDim WeaponDefs(WeaponDefUB)

            For X As Int32 = 0 To WeaponDefUB
                WeaponDefs(X) = New UnitWeaponDef

                WeaponDefs(X).yArcID = CType(yData(lPos), UnitArcs) : lPos += 1

                Dim lWpnDefID As Int32 = System.BitConverter.ToInt32(yData, lPos)
                Dim oWpnDef As WeaponDef = GetEpicaWeaponDef(lWpnDefID)
                If oWpnDef Is Nothing = True Then
                    oWpnDef = New WeaponDef()
                    lPos = oWpnDef.FillFromPrimaryMsg(yData, lPos)
                    AddWeaponDefToGlobalArray(oWpnDef)
                    WeaponDefs(X).oWeaponDef = oWpnDef
                Else
                    'adjust lPos by size of the weapondef
                    'create a temporary wpn def to do our position calcualtion
                    Dim oTmp As WeaponDef = New WeaponDef
                    lPos = oTmp.FillFromPrimaryMsg(yData, lPos)
                    'DO NOT USE oTmp!!!

                    WeaponDefs(X).oWeaponDef = oWpnDef  'use oWpnDef as it is the global version that already exists
                End If

                WeaponDefs(X).lEntityStatusGroup = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                WeaponDefs(X).mlAmmoCap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                WeaponDefs(X).ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Next X

            .lExtendedFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        End With

        Me.ProductionRequirements = New ProductionCost
        With Me.ProductionRequirements
            lPos = .FillFromPrimaryAddMsg(yData, lPos)
        End With
    End Sub

	Protected Shared Function GetMaxDPS(ByVal iModelID As Int16) As Int32
		Dim oModelDef As ModelDef = goModelDefs.GetModelDef(iModelID)
		If oModelDef Is Nothing = False Then
			Select Case oModelDef.TypeID
				Case 0	'capital... 0 = battleship, 2 = battlecruiser
                    If oModelDef.SubTypeID = 0 Then Return 110 Else Return 90
				Case 1	'escort... 0 = corvette, 1 = cruiser, 2 = destroyer, 3 = frigate, 4 = escort
					Select Case oModelDef.SubTypeID
						Case 0
                            Return 20
						Case 1
                            Return 50
						Case 2
                            Return 70
						Case 3
                            Return 15
						Case Else
                            Return 8
					End Select
				Case 2	'facility
                    Return 1000
				Case 3	'fighter... 0 = light, 1 = medium, 2 = heavy
					Select Case oModelDef.SubTypeID
						Case 1
                            Return 6
						Case 2
                            Return 7
						Case Else
                            Return 5
					End Select
				Case 4	'small vehicle
                    Return 5
				Case 5	'tank
                    Return 10
				Case 6	'transport 
                    Return 8
				Case 7	'utility
                    Return 2
			End Select
		Else : Return 1
		End If
	End Function

    Public Overridable Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		If mbDeleted = True Then Return True

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblUnitDef (PrototypeID, Q1_MaxHP, Q2_MaxHP, Q3_MaxHP, Q4_MaxHP, Maneuver, " & _
                  "MaxSpeed, FuelEfficiency, Structure_MaxHP, Hangar_Cap, HullSize, Cargo_Cap, Fuel_Cap, Weapon_Acc, " & _
                  " ScanResolution, OptRadarRange, MaxRadarRange, DisruptionResistance, PiercingResist, ImpactResist, " & _
                  "BeamResist, ECMResist, FlameResist, ChemicalResist, DetectionResist, Shield_MaxHP, Shield_Recharge, " & _
                  "ShieldRechargeFreq, ModelID, DefName, Side1Crits, Side2Crits, Side3Crits, Side4Crits, ProductionTypeID, " & _
                  "RequiredProductionTypeID, ChassisType, OwnerID, FXColors, ArmorIntegrity, JamImmunity, JamStrength, " & _
                  "JamTargets, JamEffect, ExtendedFlags) VALUES ("
                If oPrototype Is Nothing = False Then
                    sSQL &= oPrototype.ObjectID
                Else : sSQL &= "0"
                End If
                sSQL &= ", " & Q1_MaxHP & ", " & Q2_MaxHP & ", " & Q3_MaxHP & ", " & Q4_MaxHP & _
                  ", " & Maneuver & ", " & MaxSpeed & ", " & FuelEfficiency & ", " & Structure_MaxHP & ", " & _
                  Hangar_Cap & ", " & HullSize & _
                  ", " & Cargo_Cap & ", " & Fuel_Cap & ", " & Weapon_Acc & ", " & ScanResolution & ", " & _
                  OptRadarRange & ", " & MaxRadarRange & ", " & DisruptionResistance & ", " & PiercingResist & _
                  ", " & ImpactResist & ", " & BeamResist & ", " & ECMResist & ", " & FlameResist & ", " & _
                  ChemicalResist & ", " & DetectionResist & ", " & Shield_MaxHP & ", " & ShieldRecharge & ", " & _
                  ShieldRechargeFreq & ", " & ModelID & ", '" & MakeDBStr(BytesToString(DefName)) & "', " & lSideCrits(0) & ", " & _
                  lSideCrits(1) & ", " & lSideCrits(2) & ", " & lSideCrits(3) & ", " & ProductionTypeID & ", " & RequiredProductionTypeID & _
                  ", " & yChassisType & ", " & OwnerID & ", " & yFXColors & ", " & ArmorIntegrity & ", " & JamImmunity & ", " & _
                  JamStrength & ", " & JamTargets & ", " & JamEffect & ", " & lExtendedFlags & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblUnitDef SET PrototypeID = "
                If oPrototype Is Nothing = False Then
                    sSQL &= oPrototype.ObjectID
                Else : sSQL &= "0"
                End If
                sSQL &= ", Q1_MaxHP = " & Q1_MaxHP & _
                  ", Q2_MaxHP = " & Q2_MaxHP & ", Q3_MaxHP = " & Q3_MaxHP & ", Q4_MaxHP = " & Q4_MaxHP & ", Maneuver=" & _
                  Maneuver & ", MaxSpeed = " & MaxSpeed & ", FuelEfficiency = " & FuelEfficiency & ", Structure_MaxHP=" & _
                  Structure_MaxHP & ", Hangar_Cap = " & Hangar_Cap & ", HullSize = " & _
                  HullSize & ", Cargo_Cap = " & Cargo_Cap & ", Fuel_cap = " & Fuel_Cap & ", Weapon_acc = " & _
                  Weapon_Acc & ", ScanResolution = " & ScanResolution & ", OptRadarRange = " & OptRadarRange & _
                  ", MaxRadarRange = " & MaxRadarRange & ", DisruptionResistance = " & DisruptionResistance & _
                  ", PiercingResist = " & PiercingResist & ", ImpactResist = " & ImpactResist & ", BeamResist = " & _
                  BeamResist & ", ECMResist = " & ECMResist & ", FlameResist = " & FlameResist & ", ChemicalResist=" & _
                  ChemicalResist & ", DetectionResist = " & DetectionResist & ", Shield_MaxHP = " & Shield_MaxHP & _
                  ", Shield_Recharge = " & ShieldRecharge & ", ShieldRechargeFreq = " & ShieldRechargeFreq & _
                  ", ModelID = " & ModelID & ", DefName='" & MakeDBStr(BytesToString(DefName)) & "', Side1Crits = " & lSideCrits(0) & _
                  ", Side2Crits = " & lSideCrits(1) & ", Side3Crits = " & lSideCrits(2) & ", Side4Crits = " & lSideCrits(3) & ", ProductionTypeID = " & _
                  ProductionTypeID & ", RequiredProductionTypeID = " & RequiredProductionTypeID & ", ChassisType = " & yChassisType & _
                  ", OwnerID = " & OwnerID & ", FXColors = " & yFXColors & ", ArmorIntegrity = " & ArmorIntegrity & _
                  ", JamImmunity = " & JamImmunity & ", JamStrength = " & JamStrength & ", JamTargets = " & JamTargets & _
                  ", JamEffect = " & JamEffect & ", ExtendedFlags = " & lExtendedFlags & _
                  " WHERE UnitDefID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                'TODO: If oPrototype is nothing, then this will fail
                sSQL = "SELECT MAX(UnitDefID) FROM tblUnitDef WHERE PrototypeID = " & oPrototype.ObjectID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            For X As Int32 = 0 To Me.WeaponDefUB
                Me.WeaponDefs(X).oUnitDef = Me
                If Me.WeaponDefs(X).oWeaponDef.SaveObject() = True Then
                    If Me.WeaponDefs(X).SaveObject() = False Then
                        LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save Entity Def's Weapon Def. Reason: " & Err.Description)
                    End If
                Else
                    LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save Entity Def's Weapon Def. Reason: " & Err.Description)
                End If
            Next X

            'Save our hangar doors
            For X As Int32 = 0 To Me.lHangarDoorUB
                If moHangarDoors(X) Is Nothing = False Then
                    moHangarDoors(X).lEntityDefID = Me.ObjectID
                    moHangarDoors(X).iEntityDefTypeID = Me.ObjTypeID
                    moHangarDoors(X).SaveObject()
                End If
            Next X

            sSQL = "DELETE FROM tblEntityDefMineral WHERE EntityDefID = " & Me.ObjectID & " and EntityDefTypeID = " & Me.ObjTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            For X As Int32 = 0 To Me.lEntityDefMineralUB
                If Me.EntityDefMinerals(X) Is Nothing = False Then
                    Me.EntityDefMinerals(X).SaveObject()
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Public Overridable Function GetSaveObjectText() As String
		Dim sSQL As String
		If mbDeleted = True Then Return ""
		Try
			If ObjectID = -1 Then
				SaveObject()
				Return ""
			End If

			Dim oSB As New System.Text.StringBuilder


			'UPDATE
			sSQL = "UPDATE tblUnitDef SET PrototypeID = "
			If oPrototype Is Nothing = False Then
				sSQL &= oPrototype.ObjectID
			Else : sSQL &= "0"
			End If
            sSQL &= ", Q1_MaxHP = " & Q1_MaxHP & _
              ", Q2_MaxHP = " & Q2_MaxHP & ", Q3_MaxHP = " & Q3_MaxHP & ", Q4_MaxHP = " & Q4_MaxHP & ", Maneuver=" & _
              Maneuver & ", MaxSpeed = " & MaxSpeed & ", FuelEfficiency = " & FuelEfficiency & ", Structure_MaxHP=" & _
              Structure_MaxHP & ", Hangar_Cap = " & Hangar_Cap & ", HullSize = " & _
              HullSize & ", Cargo_Cap = " & Cargo_Cap & ", Fuel_cap = " & Fuel_Cap & ", Weapon_acc = " & _
              Weapon_Acc & ", ScanResolution = " & ScanResolution & ", OptRadarRange = " & OptRadarRange & _
              ", MaxRadarRange = " & MaxRadarRange & ", DisruptionResistance = " & DisruptionResistance & _
              ", PiercingResist = " & PiercingResist & ", ImpactResist = " & ImpactResist & ", BeamResist = " & _
              BeamResist & ", ECMResist = " & ECMResist & ", FlameResist = " & FlameResist & ", ChemicalResist=" & _
              ChemicalResist & ", DetectionResist = " & DetectionResist & ", Shield_MaxHP = " & Shield_MaxHP & _
              ", Shield_Recharge = " & ShieldRecharge & ", ShieldRechargeFreq = " & ShieldRechargeFreq & _
              ", ModelID = " & ModelID & ", DefName='" & MakeDBStr(BytesToString(DefName)) & "', Side1Crits = " & lSideCrits(0) & _
              ", Side2Crits = " & lSideCrits(1) & ", Side3Crits = " & lSideCrits(2) & ", Side4Crits = " & lSideCrits(3) & ", ProductionTypeID = " & _
              ProductionTypeID & ", RequiredProductionTypeID = " & RequiredProductionTypeID & ", ChassisType = " & yChassisType & _
              ", OwnerID = " & OwnerID & ", FXColors = " & yFXColors & ", ArmorIntegrity = " & ArmorIntegrity & _
              ", JamImmunity = " & JamImmunity & ", JamStrength = " & JamStrength & ", JamTargets = " & JamTargets & _
              ", JamEffect = " & JamEffect & ", ExtendedFlags = " & lExtendedFlags & _
              " WHERE UnitDefID = " & ObjectID

			oSB.AppendLine(sSQL)

			For X As Int32 = 0 To Me.WeaponDefUB
				Me.WeaponDefs(X).oUnitDef = Me

				If Me.WeaponDefs(X).oWeaponDef.SaveObject() = True Then

					oSB.AppendLine(Me.WeaponDefs(X).GetSaveObjectText())

				Else
					LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save Entity Def's Weapon Def. Reason: " & Err.Description)
				End If
			Next X

			'Save our hangar doors
			For X As Int32 = 0 To Me.lHangarDoorUB
				If moHangarDoors(X) Is Nothing = False Then
					moHangarDoors(X).lEntityDefID = Me.ObjectID
					moHangarDoors(X).iEntityDefTypeID = Me.ObjTypeID
					oSB.AppendLine(moHangarDoors(X).GetSaveObjectText())
				End If
			Next X

			Return oSB.ToString
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return ""
    End Function

    Public Overridable Sub FillFromDataReader(ByVal oData As OleDb.OleDbDataReader, ByVal iObjTypeID As Int16)
        With Me
            .BeamResist = CByte(oData("BeamResist"))
            .Cargo_Cap = CInt(oData("Cargo_Cap"))
            .ChemicalResist = CByte(oData("ChemicalResist"))
            .DetectionResist = CByte(oData("DetectionResist"))
            .DisruptionResistance = CByte(oData("DisruptionResistance"))
            .ECMResist = CByte(oData("ECMResist"))
            .FlameResist = CByte(oData("FlameResist"))
            .Fuel_Cap = CInt(oData("Fuel_Cap"))
            .FuelEfficiency = CByte(oData("FuelEfficiency"))
            .Hangar_Cap = CInt(oData("Hangar_Cap"))
            .HullSize = CInt(oData("HullSize"))
            .ImpactResist = CByte(oData("ImpactResist"))
            .Maneuver = CByte(oData("Maneuver"))
            .MaxRadarRange = CByte(oData("MaxRadarRange"))
            .MaxSpeed = CByte(oData("MaxSpeed"))

            If iObjTypeID = ObjectType.eUnitDef Then
                .ObjectID = CInt(oData("UnitDefID"))
                .DefName = StringToBytes(CStr(oData("DefName")))
            Else
                .ObjectID = CInt(oData("FacilityDefID"))
                .DefName = StringToBytes(CStr(oData("FacilityDefName")))
            End If
            .ObjTypeID = iObjTypeID

            Dim oTmpPlayer As Player = Nothing
            Dim lPlayerID As Int32 = CInt(oData("OwnerID"))
            If lPlayerID = 0 Then
                oTmpPlayer = goInitialPlayer
            Else : oTmpPlayer = GetEpicaPlayer(lPlayerID)
            End If
            If oTmpPlayer Is Nothing = False Then
                .OwnerID = oTmpPlayer.ObjectID
                .oPrototype = CType(oTmpPlayer.GetTech(CInt(oData("PrototypeID")), ObjectType.ePrototype), Prototype)
            Else
                .oPrototype = Nothing
                .OwnerID = 0
            End If

            .OptRadarRange = CByte(oData("OptRadarRange"))
            .PiercingResist = CByte(oData("PiercingResist"))
            .Q1_MaxHP = CInt(oData("Q1_MaxHP"))
            .Q2_MaxHP = CInt(oData("Q2_MaxHP"))
            .Q3_MaxHP = CInt(oData("Q3_MaxHP"))
            .Q4_MaxHP = CInt(oData("Q4_MaxHP"))
            .ScanResolution = CByte(oData("ScanResolution"))
            .Shield_MaxHP = CInt(oData("Shield_MaxHP"))
            .ShieldRecharge = CInt(oData("Shield_Recharge"))
            .ShieldRechargeFreq = CInt(oData("ShieldRechargeFreq"))
            .Structure_MaxHP = CInt(oData("Structure_MaxHP"))
            .Weapon_Acc = CByte(oData("Weapon_Acc"))
            .ModelID = CShort(oData("ModelID"))

            .JamEffect = CByte(oData("JamEffect"))
            .JamImmunity = CByte(oData("JamImmunity"))
            .JamStrength = CByte(oData("JamStrength"))
            .JamTargets = CByte(oData("JamTargets"))

            .yChassisType = CByte(oData("ChassisType"))
            .yFXColors = CByte(oData("FXColors"))
            .ArmorIntegrity = CByte(oData("ArmorIntegrity"))

            .lSideCrits(0) = CInt(oData("Side1Crits"))
            .lSideCrits(1) = CInt(oData("Side2Crits"))
            .lSideCrits(2) = CInt(oData("Side3Crits"))
            .lSideCrits(3) = CInt(oData("Side4Crits"))

            .RequiredProductionTypeID = CByte(oData("RequiredProductionTypeID"))
            .ProductionTypeID = CByte(oData("ProductionTypeID"))

            .lExtendedFlags = CInt(oData("ExtendedFlags"))

            If .oPrototype Is Nothing = False Then
                .oPrototype.moExpectedDef = Me
                .oPrototype.ResultingDef = Me
            End If
        End With
    End Sub

    'should only be called when def is being created or loaded
    Public Sub AddHangarDoor(ByVal lID As Int32, ByVal lDoorSize As Int32, ByVal ySideArc As Byte)
        Me.lHangarDoorUB += 1
        ReDim Preserve Me.moHangarDoors(Me.lHangarDoorUB)
        moHangarDoors(Me.lHangarDoorUB) = New EntityDefHangarDoor()
        With moHangarDoors(Me.lHangarDoorUB)
            .ED_HD_ID = lID
            .DoorSize = lDoorSize
            .ySideArc = ySideArc
            .lEntityDefID = Me.ObjectID
            .iEntityDefTypeID = Me.ObjTypeID
        End With
    End Sub

    Public Function HasHangarDoorSize(ByVal lHullSize As Int32) As Boolean
        For X As Int32 = 0 To lHangarDoorUB
            If moHangarDoors(X).DoorSize >= lHullSize Then Return True
        Next X
        Return False
    End Function

    Public Function GetDoorSize(ByVal lIndex As Int32) As Int32
        'NOTE: Assumes index is within bounds
        Return moHangarDoors(lIndex).DoorSize
	End Function

	Private mbDeleted As Boolean = False
	Public Function DeleteMe() As Boolean
		Dim sSQL As String = ""
		Dim oComm As OleDb.OleDbCommand = Nothing
		Dim bResult As Boolean = False

		Try

			Dim sAdditionalSQL As String = ""
			If Me.ObjTypeID = ObjectType.eFacilityDef Then
				sAdditionalSQL = "DELETE FROM tblStructure WHERE FacilityDefID = " & Me.ObjectID
				sSQL = "DELETE FROM tblStructureDef WHERE FacilityDefID = " & Me.ObjectID
			Else
				sAdditionalSQL = "DELETE FROM tblUnit WHERE UnitDefID = " & Me.ObjectID
				sSQL = "DELETE FROM tblUnitDef WHERE UnitDefID = " & Me.ObjectID
			End If

			oComm = New OleDb.OleDbCommand(sAdditionalSQL, goCN)
			Dim lAddtlDels As Int32 = oComm.ExecuteNonQuery()
			oComm.Dispose()
			oComm = Nothing

			If lAddtlDels > 0 Then
				'shouldn't happen but...
				LogEvent(LogEventType.CriticalError, "Delete Epica_Entity_Def (" & Me.ObjectID & ", " & Me.ObjTypeID & ") had " & lAddtlDels & " related objects still using it.")
			End If

			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm.Dispose()
			oComm = Nothing

			bResult = True
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, ex.Message)
		Finally
			If oComm Is Nothing = False Then oComm.Dispose()
			oComm = Nothing
		End Try

		Return bResult
	End Function
End Class