Option Strict On

Public Enum PlayerSpecialAttributeID As Short
    eSpecialSpecialAttribute = 0            'hehehehe 
    eArmorCheaperHullToHP = 1
    eAutomaticMineralMovement = 2
    eBeamResistImprove = 3
    eBurnResistImprove = 4
    eBuyOrderSlots = 5
    eChemResistImprove = 6
    eCPLimit = 7
    eDiscoverTime = 8
    eEngineMaxSpeed = 9
    eEnginePowerBonus = 10
    eEnginePowerThrustLimit = 11
    eEnlistedsColonistsCost = 12
    eFactoryUpkeepReduction = 13
    eHomelessAllowedBeforePenalty = 14
    eImpactResistImprove = 15
    eJammingEffectAvailable = 16
    eJammingImmunityAvailable = 17
    eJammingStrength = 18
    eJammingTargets = 19
    eLinkChanceBonus = 20
    eMagResistImprove = 21
    eManeuverBonus = 22
    eManeuverLimit = 23
    eMaxBattlegroupUnits = 24
    eMaxBattlegroups = 25
    eMaxnumberofconcurrentLinks = 26
    eMaxSpeedBonus = 27
    eMineralConcentrationBonus = 28
    eMissileExplosionRadiusBonus = 29
    eMissileFlightTimeBonus = 30
    eMissileHullSizeImprove = 31
    eMissileManeuverBonus = 32
    eMissileMaxDmgBonus = 33
    eMissileMaxSpeedImprove = 34
    eNonAerialUpkeepReduction = 35
    eOfficersEnlistedCost = 36
    eOtherFacUpkeepReduction = 37
    ePierceResistImprove = 38
    ePopulationUpkeepReduction = 39
    eProjectileExplosionRadius = 40
    eProjectileMaxRng = 41
    ePulseBeamMaxDmgImprove = 42
    ePulseBeamMinDmgImprove = 43
    ePulseBeamOptRngImprove = 44
    ePulseBeamOptimumRange = 45
    ePulseBeamPowerReduced = 46
    ePulseBeamROFImprove = 47
    ePulseBeamMaxCompressFactor = 48
    ePulseBeamWpnROF = 49
    ePulseBeamInputEnergyMax = 50
    eRadarDisRes = 51
    eRadarMaxRange = 52
    eRadarMaxRangeBonus = 53
    eRadarOptRange = 54
    eRadarOptRangeBonus = 55
    eRadarResistImprove = 56
    eRadarScanRes = 57
    eRadarScanResBonus = 58
    eRadarWpnAcc = 59
    eRadarWpnAccBonus = 60
    eRepairSpeedBonus = 61
    eResearchFacAssistBonus = 62
    eResearchFacUpkeepReduction = 63
    eSellOrderSlots = 64
    eShieldMaxHP = 65
    eShieldMaxHPBonus = 66
    eShieldProjSize = 67
    eShieldProjSizeBonus = 68
    eShieldRechargeFreqDecrease = 69
    eShieldRechargeFreqLow = 70
    eShieldRechargeRate = 71
    eShieldRechargeRateBonus = 72
    eSolidBeamMaxDmgBonus = 73
    eSolidBeamMinDmgBonus = 74
    eSolidBeamOptRangeBonus = 75
    eSolidBeamOptimumRange = 76
    eSolidBeamPierceRatio = 77
    eSolidBeamPowerReduced = 78
    eSolidBeamROFDecrease = 79
    eSolidBeamWeaponMaxAccuracy = 80
    eSolidBeamWpnMaxDmg = 81
    eSolidBeamWpnROF = 82
    eSpaceportUpkeepReduction = 83
    eStudyMineralEffect = 84
    eTaxRateWithoutPenalty = 85
    eTradeBoardRange = 86
    eTradeCosts = 87
    eTradeGTCSpeed = 88
    eUnderwaterFacUpkeepReduct = 89
    eUnemploymentUpkeepReduct = 90
    eUnemploymentWithoutPenalty = 91
    eAerialUpkeepReduction = 92
    eAreaEffectJamming = 93
    eBuyAndSellOrderSlots = 94

    eNavalAbility = 95
    eCommandEspionageAbility = 96
    eConstructionEspionageAbility = 97
    ePayloadTypeAvailable = 98
    ePrecisionMissiles = 99
    eTransporters = 100
    eAlloyBuilderImprovements = 101
    eColonyMoraleBoost = 102
    eGrowthRateBonus = 103
    eBetterHullToResidence = 104
    eEnlistedTrainingFactor = 105
    eBaseEnlistedPerTrainer = 106
    eBaseOfficerPerTrainer = 107
    eOfficerTrainingFactor = 108
    eMiningWorkers = 109
    eResearchCrewQtrs = 110
    eResearchProdFactor = 111
    ePowerGenProdToWorker = 112
    eDiplomaticEspionage = 113
    eResearchEspionage = 114

    eSuperSpecials = 115

    ePersonnelCargoUsage = 116
    eControlledMorale = 117
    eControlledGrowth = 118

	eMaxAlloyImprovement = 119

	eBombPayloadType = 120
	eBombMaxRange = 121
	eBombMaxGuidance = 122
	eBombMaxAOE = 123
	eBombMinROF = 124

    eAddArmorComponent = 125
    eAddEngineComponent = 126
    eAddHullComponent = 127
    eAddRadarComponent = 128
    eAddShieldComponent = 129
    eAddWeaponComponent = 130

    eNoResistAt2X = 131

    eAttributeFinalValue
End Enum

Public Class SpecialTech
	Inherits Base_Tech

	Public sName As String
	Public sDesc As String
	Public sBenefits As String
	Public lProgramControl As Int32
	Public lNewValue As Int32
    Public yFlags As Byte

    Public ReadOnly Property bInTheTank() As Boolean
        Get
            Return (yFlags And 2) <> 0
        End Get
    End Property


	Public Overrides Sub FillFromMsg(ByRef yData() As Byte)
		With Me
			Dim lPos As Int32 = 2
			.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ComponentDevelopmentPhase = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ErrorCodeReason = yData(lPos) : lPos += 1		'not used
			.Researchers = yData(lPos) : lPos += 1
            .MajorDesignFlaw = yData(lPos) : lPos += 1

            .yFlags = yData(lPos) : lPos += 1

			.sName = GetStringFromBytes(yData, lPos, 50) : lPos += 50
			lProgramControl = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			lNewValue = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.sDesc = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp
			lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.sBenefits = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp
		End With
	End Sub

	Public Overrides Function GetDesignFlawText() As String
		Return ""
	End Function

	Public Overrides Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lPos As Integer, ByVal yTechKnow As Byte, ByVal psName As String)
		sName = psName
		sDesc = ""
		sBenefits = ""

		Try
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
				'settings 1 or better
				Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				sDesc = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
				'settings 2 or better
				Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				sBenefits = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
				'full knowledge
			End If
		Catch
			'do nothing?
		End Try

	End Sub

	Public Overrides Function GetIntelReport(ByVal yLevel As PlayerTechKnowledge.KnowledgeType) As String
		Dim oSB As New System.Text.StringBuilder
		If yLevel > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
			'ylevel = PlayerTechKnowledge.KnowledgeType.eSettingsLevel1			
			oSB.AppendLine("Description: " & sDesc & vbCrLf)
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'ylevel = PlayerTechKnowledge.KnowledgeType.eSettingsLevel2
			oSB.AppendLine("Benefits: " & sBenefits & vbCrLf)
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
			'ylevel = PlayerTechKnowledge.KnowledgeType.eFullKnowledge 
			 
		End If
		Return oSB.ToString
	End Function
End Class
