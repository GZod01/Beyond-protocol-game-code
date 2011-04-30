Public Class PrototypeTech
	Inherits Base_Tech

	Public Structure PrototypeWeapon
		Public WeaponID As Int32
		Public SlotX As Byte
		Public SlotY As Byte
		Public WeaponGroupTypeID As WeaponGroupType
	End Structure

	Public PrototypeName As String
	Public EngineID As Int32
	Public ArmorID As Int32
	Public HangarID As Int32
	Public HullID As Int32
	Public RadarID As Int32
	Public ShieldID As Int32

    Public MaxCrew As Int32 = 0
    'Public MinCrew As Int16		'???

	Public ForeArmorUnits As Int32
	Public LeftArmorUnits As Int32
	Public RearArmorUnits As Int32
	Public RightArmorUnits As Int32
	Public ProductionHull As Int32

	Public oActualResults As EntityDef = Nothing
	Public lActualResultsWorkers As Int32 = 0

	Public oWeapons() As PrototypeWeapon
	Public lWeaponUB As Int32 = -1

	Public Overrides Sub FillFromMsg(ByRef yData() As Byte)
		With Me
			Dim lPos As Int32 = 2
			Dim lCnt As Int32

			.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If .OwnerID = 0 Then .OwnerID = glPlayerID
			.ComponentDevelopmentPhase = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ErrorCodeReason = yData(lPos) : lPos += 1
			.Researchers = yData(lPos) : lPos += 1
			.MajorDesignFlaw = yData(lPos) : lPos += 1

			'Now, the Prototype tech specifics...
			.PrototypeName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.EngineID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ArmorID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.HullID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.RadarID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ShieldID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.ForeArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.RearArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.LeftArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.RightArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ProductionHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .MaxCrew = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            '.MinCrew = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

			lCnt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			lWeaponUB = lCnt - 1
			ReDim oWeapons(lWeaponUB)
			For X As Int32 = 0 To lCnt - 1
				With oWeapons(X)
					.WeaponID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.SlotX = yData(lPos) : lPos += 1
					.SlotY = yData(lPos) : lPos += 1
					.WeaponGroupTypeID = CType(yData(lPos), WeaponGroupType) : lPos += 1
				End With
			Next X

			lActualResultsWorkers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			If .ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
				'ok, expect an entitydef at the end...
				If oActualResults Is Nothing Then oActualResults = New EntityDef()

				Dim lTemp As Int32

				With oActualResults
					.ProductionCost = New ProductionCost()
					.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    lPos += 4        'for ownerid
					.Armor_MaxHP(0) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.Armor_MaxHP(1) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.Armor_MaxHP(2) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.Armor_MaxHP(3) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.Maneuver = yData(lPos) : lPos += 1
					.MaxSpeed = yData(lPos) : lPos += 1

					'.FuelEfficiency = yData(lPos) : lPos += 1
					lPos += 1		'old fuelefficiency

					.Structure_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.Hangar_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.Cargo_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.Shield_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.ShieldRecharge = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.ShieldRechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					''Max Door Size (4)
					'lPos += 4
					''Number of Doors (1)
					'lPos += 1

					'.Fuel_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					lPos += 4		'old fuel_cap

					.Weapon_Acc = yData(lPos) : lPos += 1
					.ScanResolution = yData(lPos) : lPos += 1
					.OptRadarRange = yData(lPos) : lPos += 1
					.MaxRadarRange = yData(lPos) : lPos += 1
					.DisruptionResistance = yData(lPos) : lPos += 1
					.PiercingResist = yData(lPos) : lPos += 1
					.ImpactResist = yData(lPos) : lPos += 1
					.BeamResist = yData(lPos) : lPos += 1
					.ECMResist = yData(lPos) : lPos += 1
					.FlameResist = yData(lPos) : lPos += 1
					.ChemicalResist = yData(lPos) : lPos += 1
					'Detection Resist (1)
					lPos += 1
					.ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

					'DefName (20)
					.DefName = GetStringFromBytes(yData, lPos, 20)
					lPos += 20


					If .ObjTypeID = ObjectType.eFacilityDef Then
						lPos += 4 'WorkerFactor (4)
						lPos += 1 'MaxFacilitySize (1)
						.ProdFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.PowerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					End If
					.ProductionTypeID = yData(lPos) : lPos += 1
					.RequiredProductionTypeID = yData(lPos) : lPos += 1
					.yChassisType = yData(lPos) : lPos += 1
					.yFXColors = yData(lPos) : lPos += 1
					.yArmorIntegrityRoll = yData(lPos) : lPos += 1
					.JamImmunity = yData(lPos) : lPos += 1
					.JamStrength = yData(lPos) : lPos += 1
					.JamTargets = yData(lPos) : lPos += 1
					.JamEffect = yData(lPos) : lPos += 1

					'Side Critical Hit Locations...
					lPos += 16

					lTemp = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
					.WeaponDefUB = lTemp - 1
					ReDim .WeaponDefs(.WeaponDefUB)
				End With

				lTemp = 0

				For X As Int32 = 0 To oActualResults.WeaponDefUB
					oActualResults.WeaponDefs(lTemp) = New WeaponDef()
					With oActualResults.WeaponDefs(lTemp)
						.ArcID = yData(lPos) : lPos += 1
						.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

						'Weapon Name (20)
						.WeaponName = GetStringFromBytes(yData, lPos, 20)
						lPos += 20

                        .yWeaponType = yData(lPos) : lPos += 1
						.ROF_Delay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
						.Range = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
						.Accuracy = yData(lPos) : lPos += 1
						.PiercingMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.PiercingMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.ImpactMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.ImpactMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.BeamMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.BeamMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.ECMMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.ECMMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.FlameMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.FlameMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.ChemicalMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.ChemicalMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.WpnGroup = yData(lPos) : lPos += 1
						.lFirePowerRating = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						'wpn tech id
						lPos += 4
						.AOERange = yData(lPos) : lPos += 1
						.WeaponSpeed = yData(lPos) : lPos += 1
						.Maneuver = yData(lPos) : lPos += 1

						.lEntityStatusGroup = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.lAmmoCap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.WpnDefID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

						.MaxRangeAccuracy = .Range * (.Accuracy / 100.0F)
					End With
					lTemp += 1
				Next X

                If gl_CLIENT_VERSION > 308 Then
                    .oActualResults.lExtendedFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                End If
			End If
		End With
	End Sub

	Public Overrides Function GetDesignFlawText() As String
		Return ""
	End Function

	Public Overrides Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lPos As Integer, ByVal yTechKnow As Byte, ByVal sName As String)
		'TODO: And what do we do here?
	End Sub

	Public Overrides Function GetIntelReport(ByVal yLevel As PlayerTechKnowledge.KnowledgeType) As String
		Return "Not yet implemented, bug me"
	End Function
End Class
