Public Enum eMissionPhase As Byte
	eInPlanning = 0
	ePreparationTime = 1
	eSettingTheStage = 2
	eFlippingTheSwitch = 3
	eReinfiltrationPhase = 4

	eWaitingToExecute = 10
	eCancelled = 11
	eMissionOverFailure = 12
	eMissionOverSuccess = 13

	eMissionPaused = 128			'logically OR'd with other values
End Enum
Public Enum eMissionResult As Int32
    eUndefined = 0
    eShowQueueAndContents = 1           'show prod/res/whatever else queue and contents
    eSlowProduction = 2                 'Slow production (research/mining/etc...)
    eSabotageProduction = 3             'sabotage production
    eCorruptProduction = 4              'reduces reliability
    eDestroyFacility = 5                'destroys target facility
    eClutterCargoBay = 6                'reduces cargo bay capacity
    eDoorJamHangar = 7                  'increases docking/undocking delay
    eStealCargo = 8                     'steal cargo
    eSpecialProjectList = 9             'displays the special project technology list
    eComponentDesignList = 10           'Gets list of designed components
    eComponentResearchList = 11         'gets list of researched components
    eGetMineralList = 12                'gets the mineral list
    eAcquireTechData = 13               'gets the data of a component/prototype/special tech
    eGetTradeTreaties = 14              'gets the current trade treaties
    eGetFacilityList = 15               'gets a list of the facilities for the provided production type
    eSiphonItem = 16                    'pulls credits/minerals/production/etc... based on infiltration
    eBadPublicity = 17                  'reduces trade revenue from allies
    eGetColonyBudget = 18               'returns budget for the colony
    eAssassinateGovernor = 19           'short-term morale and production penalty
    eDecreaseTaxMorale = 20             'colony morale decreases based on Taxes (tax morale modifier)
    eDecreaseUnemploymentMorale = 21    'colony morale decreases based on unemployment (unemployment modifier)
    eDecreaseHousingMorale = 22       'colony morale decreases based on war sentiment (sentiment modifier)
    eGetAlliesList = 23                 'gets the allies with the values
    eGetEnemyList = 24                  'gets the enemies with the values
    ePlantEvidence = 25                 'plants evidence
    eDiplomaticStrength = 26            'gets the diplomacy score
    eMilitaryStrength = 27              'gets the military score
    eTechnologyStrength = 28            'gets the tech score
    ePopulationScore = 29               'gets the population score
    eProductionScore = 30               'gets the production score
    eMilitaryCoup = 31                  'half units in an environment change alignment

    eFindMineral = 32                   'searches for a specific mineral within the infiltrated system
    eFindCommandCenters = 33            'searches for command centers within the infiltrated planet belonging to the target player
    eFindSpaceStations = 34             'searches for space stations within the infiltrated system belonging to the target player
    eFindProductionFac = 35             'searches for production facilities within the infiltrated planet belonging to the target player
    eFindResearchFac = 36               'searches for research facilities within the infiltrated planet belonging to the target player

    eAssassinateAgent = 37              'searches for an agent (specific or any) belonging to a target player within the infiltrated system and kills them
    eCaptureAgent = 38                  'searches for an agent (specific or any) belonging to a target player within the infiltrated system and captures them
    eSearchAndRescueAgent = 39          'searches for an agent within the infiltrated system and rescues them

    eReconPlanetMap = 40                'exposes an area of the fog of war for a period of time (temporary or permanent based on mission effect) of the infiltrated planet
    eGeologicalSurvey = 41              'Lists minerals for the infiltrated planet

    eRelayBattleReports = 42            'Provides detailed reports about combat
    eIFFSabotage = 43                   'enables friendly fire
	eDegradePay = 44					'degrade the pay for corporation
    eTutorialFindFactory = 45

    eAgencyAnalysis = 46                'get the strength of the target counter agent network
    eGetAgentList = 47
    eDropInvulField = 48
    eLocateWormhole = 49
    eAudit = 50
    eIncreasePowerNeeds = 51
    eGetAliasList = 52
    eImpedeCurrentDevelopment = 53
    eDestroyCurrentSpecialProject = 54
End Enum

Public Class Mission
    Inherits Base_GUID

    Public sMissionName As String
    Public sMissionDesc As String

    Public lInfiltrationType As eInfiltrationType = eInfiltrationType.eGeneralInfiltration

    Public ProgramControlID As Int16
    Public BaseEffect As Int16
    Public Modifier As Int16

    Public Goals() As Goal
    Public MethodIDs() As Int32     'stored as int32, but is transmitted as byte
    Public GoalUB As Int32 = -1

    Public Sub FillFromMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .sMissionName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            .Modifier = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .BaseEffect = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .ProgramControlID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .lInfiltrationType = CType(yData(lPos), eInfiltrationType) : lPos += 1
            .sMissionDesc = GetStringFromBytes(yData, lPos, 255) : lPos += 255

            GoalUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
            ReDim MethodIDs(GoalUB)
            ReDim Goals(GoalUB)
            For X As Int32 = 0 To GoalUB
                MethodIDs(X) = yData(lPos) : lPos += 1
                Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                For Y As Int32 = 0 To glGoalUB
                    If goGoals(Y).ObjectID = lID Then
                        Goals(X) = goGoals(Y)
                        Exit For
                    End If
                Next Y
            Next X
        End With
    End Sub

End Class
