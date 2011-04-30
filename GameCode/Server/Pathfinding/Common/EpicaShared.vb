Option Strict On

Imports System.Net
Imports System.Net.Sockets
Imports System.Text

#Region " Global Enumerations "

Public Enum ObjectType As Short
    eUniverse = 0       'the grandaddy of all objects, this object is intangible but I figured I'd include it here
    eGalaxy             'contained in Universe objects
    eSolarSystem        'contained in Galaxy objects
    ePlanet             'contained in Solar System objects
    eColony             'contained in either a solar system or planet
    'eLandUnit           'contained only on planets or in hangars
    'eSpaceUnit          'contained only in space
    'eAeroUnit           'contained in space or in air of planets
    eAgent              'an agent/spy
    eAlloyTech          'an alloy technology design specification (not the alloy itself)
    eArmorTech          'an armor technology design specification (not the armor itself)
    eCorporation        'a corporation entity, NPC representing the collective interests of a corporation
    eEngineTech         'an Engine technology design specification (not the engine itself)
    eFacility           'an instance of a structure
    eFacilityDef        'definition of a structure's basic properties (not of an actual structure)
    eGNS                'a GNS Body, for example, USA Today would be GNS 1, and New York Times would be GNS 2
    eGoal               'an agent or mission goal
    eHangarTech         'a Hangar technology design specificaton (not the hangar itself)
    eHullTech           'a Hull technology design specification (not the hull itself)
    eMineral            'a mineral definition (not representing a cache of minerals)
    eMineralCache       'a cache of minerals
    eMission            'an agent mission
    eNebula             'contained in galaxy objects
    eObjectOrder        'an order assigned to an object
    ePlayer             'a player object
    ePlayerComm         'a player communication (email)
    ePrototype          'a unit design (not the actual unit)
    eRadarTech          'a radar technology design specification (not the radar itself)
    eSenate             'a senate body
    eSenateLaw          'a senate law
    eShieldTech         'a shield technology design specification (not the shield itself)
    eSkill              'an skill for agents, goals, etc...
    eUnit               'an instance of a unit
    eUnitDef            'a more precise version of a unit's definition
    eUnitGroup          'a group of units (move together, etc...)
    eUnitWeaponDef      'a weapon def to unit def relationship
    eWeaponDef          'a temporary, smaller, more need-to-know information about the weapon
    eWeaponTech         'a weapon technology design specification (not the weapon itself)
    eWormhole           'contained in solar system objects and link two objects together

    eProductionCost     'production cost item...

    'Intangibles...
    ePower              'Power Points
    eCredits            'Galactic Bartering Units
    eStock              'stock certificates
    eMorale             'morale points
    eColonists          'colonists
    eFood               'Food
    eMiningPower        'Mining
    eEnlisted           'General military personnel (infantry and crew are synonomous?)
    eOfficers           'officers required for commanding ships and large armies
    eProductionPoints   'TODO: are we sure of this one???
    eAmmunition
    eMineralProperty
    eMineralTech        'represents mineral researching

    'Component Cache
    eComponentCache     'like a mineral cache but it references a technological component

    eSpecialTech
    eSpecialTechPreq

    eTrade              'a trade object

    eRepairItem         'intangible

    eBudget

    eGuild
    eVoteProposal
    ePlayerIntel

    eGalacticTrade

    ePlayerTechKnowledge
    eTradeTreaty        'intangible, money gained through alliance with another player
	ePlayerItemIntel	'intangible, knowledge of an item belonging to another player
End Enum

Public Enum eiTacticalAttrs As Short    '16 bit integer
    eFighterClass = 1               'hull < max fighter hull and flies
	eCargoShipClass = 2				'cargo space > hull / 3
	eCarrierClass = 4				'Hangar > hull / 3
	eTroopClass = 8					'(Heavy Escort) = 30k < HULL < min battlecruiser in Space, Infantry/Troopers/Mechanized Infantry/Walkers on ground
	eEscortClass = 16				'In Space: Hull > max fighter hull but <= 30k, on ground: Hull > max fighter hull but < min battlecruiser hull
    eCapitalClass = 32              'In Space: Hull >= min battlecruiser hull, on ground: Tanks
	eArmedUnit = 64					'Any unit with at least 1 weapon
	eUnarmedUnit = 128				'Any unit with no weapons
	eFacility = 256					'A building, whether it is in space or on the ground

	eUNUSED1 = 512

	eFighterTargetEngines = 1024	'fighter ability to target engine sub systems
	eFighterTargetHangar = 2048		'fighter ability to target hangar sub systems
	eFighterTargetRadar = 4096		'fighter ability to target radar sub systems
	eFighterTargetShields = 8192	'fighter ability to target shield sub systems
	eFighterTargetWeapons = 16384	'fighter ability to target weapon sub systems
End Enum

Public Enum eiBehaviorPatterns As Int32
    eEngagement_Hold_Fire = 1
    eEngagement_Stand_Ground = 2
    eEngagement_Pursue = 4
    eEngagement_Engage = 8
    eEngagement_Evade = 16
    eEngagement_Dock_With_Target = 32  'the unit's target will always be set to the target to dock with, the secondary target will be available though
    eTactics_Minimize_Damage_To_Self = 64  'Unit will always angle itself so its strongest armored side faces the enemy
    eHoldPrimaryWpnGrp = 128                  'if flagged, the unit will NOT use the primary weapon group
    eTactics_Maximize_Damage = 256            'the unit will always angle itself so its strongest weapons face the enemy
    eTactics_Normal = 512                   'A happy medium between Minimize_Damage_To_Self and Maximize_Damage
    eTactics_LaunchChildren = 1024          'the unit will launch any children units that it can
    eTactics_Maneuver = 2048            'the unit will "dance" around the target
    eDockDuringBattle = 4096
    eHoldSecondaryWpnGrp = 8192         'if flagged, the unit will NOT use the secondary weapon groups
    eHoldPointDefenseWpnGrp = 16384     'if flagged, the unit will NOT use teh point defense weapons
    eStopToEngage = 32768               'if flagged, the unit will stop player_initiated unit moves to engage targets.  Move will resume using TetherPoint
    eAssistOrbit = 65536                'if flagged, the unit will either go from Planet to Orbit or Orbit to Planet to defend against detected enemies.  Will return to prev location after battle.
End Enum

Public Enum elUnitStatus As Integer     '32 bit long
    eEngineOperational = 1
    eRadarOperational = 2
    eShieldOperational = 4          'even if the shield has no hitpoints, it is still operational
    eHangarOperational = 8
    eCargoBayOperational = 16
    eFuelBayOperational = 32        'first hit = inoperable, second hit = unit destroyed
    'Reduced out the weapon hit locs by half. we use to have 4 per side, now its 2 per side
    eForwardWeapon1 = 64
    eForwardWeapon2 = 128
    eLeftWeapon1 = 256
    eLeftWeapon2 = 512
    eAftWeapon1 = 1024
    eAftWeapon2 = 2048
    eRightWeapon1 = 4096
    eRightWeapon2 = 8192
    eUndefined14 = 16384            'was eAllArcWeapon1
    eUndefined15 = 32768            'was eAllArcWepaon2
    eUnitMoving = 65536                     'indicates if the unit is moving
    eMovedByPlayer = 131072                 'indicates that, if the unit is moving, the unit was moved by the player as opposed to the AI
    eTargetingByPlayer = 262144             'indicates that, if the unit is targeting, the target was assigned by the player as opposed to the AI
    eUnitEngaged = 524288                  'indicates that the unit has a primary target, aggressed
    eUnitCloaked = 1048576                  'Indicates the unit is cloaked
    eSide1HasTarget = 2097152               'indicates that the entity has a target in ArcID 0
    eSide2HasTarget = 4194304               'indicates that the entity has a target in ArcID 1
    eSide3HasTarget = 8388608               'indicates that the entity has a target in ArcID 2
    eSide4HasTarget = 16777216              'indicates that the entity has a target in ArcID 3
    eGuildAsset = 33554432

    eUndefined26 = 67108864
    eUndefined27 = 134217728
    eUndefined28 = 268435456

    eFacilityPowered = 536870912
    eMoveLock = 1073741824          'set by Primary server to Region Servers to indicate the entity is busy

    eUndefined32 = -2147483648      'DO NOT USE! OR and AND have issues especially when subtracting
End Enum

'Checking these values... imagine this as a score... higher scores = better relationship
Public Enum elRelTypes As Byte
    eBloodWar = 10                          'absolute war, no gray lines, no chance of mercy, kill on sight
    eWar = 40
    eNeutral = 60
    ePeace = 80
    eAlly = 100
    eBloodBrother = 150                     'absolute peace, units should assist each other, etc...
End Enum

Public Enum UnitArcs As Byte
    eForwardArc = 0
    '    eRightArc
    eLeftArc
    eBackArc
    '    eLeftArc
    eRightArc
    eAllArcs
End Enum

Public Enum PlanetType As Byte
    eAcidic = 0
	eBarren = 1
	eGeoPlastic = 2
	eDesert = 3
	eAdaptable = 4
	eTundra = 5
	eTerran = 6
	eWaterWorld = 7

    eSuperGiant = 254       'Jupiter
    eGasGiant = 255         'Saturn
End Enum

Public Enum WeaponType As Byte
    'Short Pulse - looks like a ball of charged energy, the color indicates the primary component of the color
    eShortGreenPulse = 0
    eShortRedPulse
    eShortTealPulse
    eShortYellowPulse
    eShortPurplePulse

    'Flicker Beam - a solid line to the target that flickers 3 times
    eFlickerGreenBeam
    eFlickerRedBeam
    eFlickerTealBeam
    eFlickerYellowBeam
    eFlickerPurpleBeam

    'Solid Beam - a solid line that acts like a cutting beam
    eSolidGreenBeam
    eSolidRedBeam
    eSolidTealBeam
    eSolidYellowBeam
    eSolidPurpleBeam

    'Capital Solid Beam - larger version of the Solid Beam
    eCapitalSolidGreenBeam
    eCapitalSolidRedBeam
    eCapitalSolidTealBeam
    eCapitalSolidYellowBeam
    eCapitalSolidPurpleBeam

    'Metallic Projectile - a bullet
    eMetallicProjectile_Gold
    eMetallicProjectile_Silver
    eMetallicProjectile_Lead
    eMetallicProjectile_Bronze
    eMetallicProjectile_Copper

    eMissile_Color_1
    eMissile_Color_2
    eMissile_Color_3
    eMissile_Color_4
    eMissile_Color_5
    eMissile_Color_6
    eMissile_Color_7
    eMissile_Color_8
    eMissile_Color_9

    'Bombs - rounded objects that fall to the surface
    eBombGreen_Small
    eBombGreen_Med
    eBombGreen_Large
    eBombGray_Small
    eBombGray_Med
    eBombGray_Large
    eBombRed_Small
    eBombRed_Med
    eBombRed_Large
    eBombTeal_Small
    eBombTeal_Med
    eBombTeal_Large

    'Land-Based Mines
    eLB_Mine

    'Space-based mines
    eSB_Mine

    eShieldHitBitMask = 128             'bit-wise
End Enum

Public Enum RemovalType As Byte
    eObjectDestroyed = 0
    eChangingEnvironments
    eDocking                    '???
    eEnteringWarp               '???
    eJumping                    'wormhole, jump gate, etc...
End Enum

Public Enum WeaponGroupType As Byte
    PrimaryWeaponGroup = 0
    SecondaryWeaponGroup
    PointDefenseWeaponGroup
    BombWeaponGroup
End Enum

Public Enum ProductionType As Byte
    eNoProduction = 0
    eProduction = 1
    ePersonnel = 2
    eFacility = 4
    eMiscellaneous = 8
    eMaterials = 16
    ePowerCenter = 32

    eType2Shift = 64
    eType3Shift = 128

    '================Helpers==============
    eLandProduction = eProduction
    eNavalProduction = eProduction Or eType2Shift
    eAerialProduction = eProduction Or eType3Shift
    eAllProduction = eLandProduction Or eNavalProduction Or eAerialProduction
    eColonists = ePersonnel
    eEnlisted = ePersonnel Or eType2Shift
    eOfficers = ePersonnel Or eType3Shift
    eEnlistedAndOfficers = eEnlisted Or eOfficers
    'eFacility
    eCommandCenterSpecial = eFacility Or eType2Shift
    eSpaceStationSpecial = eFacility Or eType3Shift
    eFood = eMiscellaneous
    eResearch = eMiscellaneous Or eType2Shift
    eMorale = eMiscellaneous Or eType3Shift
    eMining = eMaterials
    eRefining = eMaterials Or eType2Shift
    eCredits = eMaterials Or eType3Shift

    eWareHouse = ePowerCenter Or eType2Shift
    eTradePost = ePowerCenter Or eType3Shift
End Enum

'Public Enum ProductionType As Byte
'    'NOTE: THESE ARE NOT BIT-WISE!!!!
'    eNoProduction = 0
'    eLandUnitProductionPoints       'produces Land units
'    eNavalUnitProductionPoints      'produces Naval units (ocean-going)
'    eAerialUnitProductionPoints     'produces aerial/space units
'    eFacilityProductionPoints       'produces Facilities
'    ePowerCenter
'    eColonists
'    eFoodProduction
'    eMiningPoints       '???
'    eEnlisted
'    eOfficers
'    eMoralePoints
'    eResearchPoints
'    eCredits
'    eCommandCenterSpecial
'    eSpaceStationProductionPoints   'Produces Space Station Facilities
'    eRefiningPoints                 'refining ability

'    eSpaceStationSpecialAbility     'the space station special ability

'    'NOTE: we could eventually upgrade our production engine to use level-based... 2nd and 3rd... by using OR on these values:
'    e2ndLevel = 64
'    e3rdLevel = 128

'    'An example of this is:
'    'yLevel = yVal And (ProductionType.e2ndLevel + ProductionType.e3rdLevel)
'    'yType = yVal - yLevel
'End Enum

Public Enum MineralCacheType As Byte
	eMineable = 0		'subsurface, requires mining equipment to extract

	'ePickupable = 1		'on surface or in space, does not require mining equipment to extract
	'non-zero values indicate a pick-up capable cache... this value becomes a descriptor of that cache

	'These values indicate the type of debris
	eGround = 1			'ground based unit debris
	eFlying = 2			'aerial/space unit debris
	eNaval = 4			'naval unit debris

	'these values indicate the type of component...
	eArmorDebris = 8	'armor tech
	eEngineDebris = 16	'engine tech
	eRadarDebris = 32	'radar tech
	eShieldDebris = 64	'shield tech
	eWeaponDebris = 128	'weapon tech
End Enum

Public Enum ChangeEnvironmentType As Byte
    eNoChangeEnvironment = 0
    eSystemToSystem
    eSystemToPlanet
    ePlanetToSystem
    eDocking
    eUndocking
End Enum

Public Enum ChassisType As Byte     'BIT WISE ENUM
    'Examples, tank = 1... airplane = 2... capital ship = 4... space fighter = 6 (able to enter all environments)
    eGroundBased = 1        'cannot fly, cannot enter space
    eAtmospheric = 2        'Can fly in planet atmospheres
    eSpaceBased = 4         'Can fly in space...
    eNavalBased = 8         'ocean 
End Enum

Public Enum BombardType As Byte
    eNormal_BT = 0
    eHighYield_BT
    eTargeted_BT
    ePrecision_BT
End Enum

Public Enum WeaponClassType As Byte
    eUndefined = 0
    eEnergyBeam = 1
    eProjectile = 2
    eMissile = 3
    eBomb = 4
    eMine = 5
    eEnergyPulse = 6
End Enum

Public Enum DockRejectType As Byte
    eDockeeMoving = 0
    eDoorSizeExceeded = 1
    eInsufficientHangarCap = 2
    eHangarInoperable = 3
    eDockerCannotEnterEnvir = 4
End Enum

Public Enum PlayerAlertType As Byte
    eUnderAttack = 0
    eEngagedEnemy = 1
    eUnitLost = 2
    eFacilityLost = 3
    eColonyLost = 4
    eBuyOrderAccepted = 5
End Enum

Public Enum ChatMessageType As Byte
    eLocalMessage = 0
    eSysAdminMessage = 1
    eChannelMessage = 2
    eAllianceMessage = 3
    eSenateMessage = 4
    ePrivateMessage = 5
    eNotificationMessage = 6
    eAliasChatMessage = 7
End Enum

Public Enum NewsItemType As Byte
	eSpaceCombat = 0
	eSpaceCombatUpdate = 1
	eLowMorale = 2
	eTerrorism = 3
	eTitlePromotionLow = 4
	ePlanetFall = 5
	eMajorProdStart = 6
	ePlanetControlShifts = 7
	ePlanetWideLowResources = 8
	eLostColony = 9
	ePlayerDeath = 10
	ePirateElimination = 11
	eNewSpaceStation = 12
	eWarDeclaration = 13
	eCeaseHostilities = 14
	eMajorAlliances = 15
	eOrbitalBombardment = 16
	eSpaceCombatEnd = 17
	ePlanetCombat = 18
	ePlanetCombatUpdate = 19
	ePlanetCombatEnd = 20
	eOfflineWarDec = 21
	eTitlePromotionMed = 22
	eTitlePromotionHi = 23
End Enum
#End Region

Module EpicaShared

#Region " Environment Grid Constants Game-Wide "

    'Small Grid Array... two of them, one for system, one for planet
    Public Const gl_SMALL_PER_ROW As Int32 = 20

    'Public Const gl_SYSTEM_SMALL_PER_ROW As Int32 = 20
    'Public Const gl_SYSTEM_SMALL_GRID_UB As Int32 = gl_SYSTEM_SMALL_PER_ROW * gl_SYSTEM_SMALL_PER_ROW - 1
    'Public Const gl_PLANET_SMALL_PER_ROW As Int32 = 20
    'Public Const gl_PLANET_SMALL_GRID_UB As Int32 = gl_PLANET_SMALL_PER_ROW * gl_PLANET_SMALL_PER_ROW - 1

    'Both small grids use the same square size... this is so we can use the same final grid (space conservation)
    Public Const gl_SMALL_GRID_SQUARE_SIZE As Int32 = 400 '320

    'Final Grid Array (tiny)
    Public Const gl_FINAL_PER_ROW As Int32 = 16 '20 
    Public Const gl_HALF_FINAL_PER_ROW As Int32 = CInt(gl_FINAL_PER_ROW / 2)

    'Public Const gl_PLANET_FINAL_MAX As Int32 = (gl_PLANET_SMALL_PER_ROW * gl_FINAL_PER_ROW * 3)
    'Public Const gl_SYSTEM_FINAL_MAX As Int32 = (gl_SYSTEM_SMALL_PER_ROW * gl_FINAL_PER_ROW * 3)
    Public Const gl_FINAL_MAX As Int32 = (gl_SMALL_PER_ROW * gl_SMALL_PER_ROW * 3)
    Public Const gl_FINAL_GRID_SQUARE_SIZE As Int32 = CInt(gl_SMALL_GRID_SQUARE_SIZE / gl_FINAL_PER_ROW)

#End Region

    Public Const gl_EXP_POINT_BASE As Int32 = 3 '5     'NOTE: Change this in order to curb experience gains
    Public Const gl_MAX_EXP_GAIN As Int32 = 15 '50      'NOTE: This will be the maximum experience rewarded at once

    Public Const gl_TINY_PLANET_CELL_SPACING As Int32 = 200
    Public Const gl_SMALL_PLANET_CELL_SPACING As Int32 = 300
    Public Const gl_MEDIUM_PLANET_CELL_SPACING As Int32 = 400
    Public Const gl_LARGE_PLANET_CELL_SPACING As Int32 = 600
    Public Const gl_HUGE_PLANET_CELL_SPACING As Int32 = 800

    Public Const gl_MINIMUM_MINING_FAC_SNAP_TO_DISTANCE As Int32 = 1000

    Public Const gl_HARDCODE_PIRATE_PLAYER_ID As Int32 = 25

    Public gb_IN_OPEN_BETA As Boolean = True

    Public Const gl_CLIENT_VERSION As Int32 = 307

    Public Enum LoginResponseCodes As Integer
        eInvalidUserName = -1
        eInvalidPassword = -2
        eAccountSuspended = -3      'temporary suspension
        eAccountBanned = -4         'permanent suspension
        eLoginAttemptLockout = -5   'max number of login attempts exceeded
        eUnknownFailure = -6        '*shrug*
        eAccountInactive = -7
        eAccountSetup = -8
        eAccountInUse = -9
        ePlayerIsDying = -10
    End Enum

    Public Enum GlobalMessageCode As Short
        eRequestObject = 0
        eRequestObjectResponse

        eStopObjectRequest          'the client is requesting the object to stop
        eStopObjectCommand          'the server is commanding the stop object
        eStopObjectRequestDeny      'denying the stop object request
        eMoveObjectRequest          'the client is requesting the object to move
        eMoveObjectCommand          'this is the actual command
        eMoveObjectRequestDeny      'denying the move object request

        eUnitHPUpdate               'the specified unit's hitpoints have been updated

        eFireWeapon                 'Fire Weapon packet (Weapon hits)
        eFireWeaponMiss             'Fire weapon packet (weapon misses)

        eBurstEnvironmentRequest    'a client is requesting to receive an Environments data
        eBurstEnvironmentResponse   'the response with the environment data

        ePathfindingConnectionInfo  'the pathfinding server's connection information

        eAddObjectCommand           'add object thru a BurstMessage, default Add Object... Undock and Created use this too
        eAddObjectCommand_CE        'Add Object thru Change environments ONLY as the client will display it differently
        eRemoveObject               'object removal

        ePlayerLoginRequest         'player wishes to login with Username/Password
        ePlayerLoginResponse        'response to the login message, a negative value indicates failure, otherwise, it is the PlayerID

        eRequestPlayerList          'requests the ENTIRE player list. To be used only for region server initialization
        eRequestPlayerRelList       'requests all player rels. To be used only for region server initialization
        eDomainRequestEnvirObjects  'domain server requesting all objects from the primary server for an environment
        eRegisterDomain             'the domain server is telling the primary server what domain it controls
        eDomainServerReady          'the domain server is fully initialized and is ready

        eAddPlayerRelObject

        eRequestPlayerEnvironment   'the client is asking what environment the player passed in (ID) was last looking at
        ePlayerCurrentEnvironment   'the response to the RequestPlayerEnvironment message
        eRequestEnvironmentDomain   'requests the Domain Credentials for the passed in Environment
        eEnvironmentDomain          'the response to the RequestEnvironmentDomain message
        eChangingEnvironment        'the player is notifying a server that they are changing environments, contains new ID's

        eRequestGalaxyAndSystems    'the client is asking for the galaxy and system objects from the Primary
        eResponseGalaxyAndSystems   'the response from the Primary to the Client containing the Galaxy and System objects

        eRequestStarTypes           'the client is asking for the star type definitions from the Primary
        eResponseStarTypes          'the response from the Primary to the Client containing the Star Types

        eRequestSystemDetails       'the client needs the planet data for the items in the system
        eResponseSystemDetails      'the response from the primary to the client containing the details of the system
        eResponseSystemDetailsFail  'the response from the primary to the client indicating that the request failed

        eSetEntityProduction        'a request to the Primary Server to set the production of an entity
        eSetEntityProdFailed        'response to eSetEntityProduction request as a failure
        eSetEntityProdSucceed       'response to eSetEntityProduction request as a success

        eGetEntityProduction        'request and response to the Primary Server

        eRequestPlayerDetails       'client is requesting their player details

        eAddWaypointMsg             'a client is requesting to add a waypoint to a unit's destination list
        eDecelerationImminent       'domain server is notifying the pathfinding server that an entity is decelerating

        eEntityChangeEnvironment    'Change environment message indicating an entity is moving or has moved between environments

        eGetEntityDetails           'communication for getting/responding entity details...

        eRequestEntityContents      'used for request and response of an entity's contents (both cargo and hangar)

        eSetPrimaryTarget           'Client to Server, informs server to set primary target
        eRequestDock                'docking request
        eRequestDockFail            'docking request failed
        eRequestUndock              'request to undock unit
        eUpdatePlayerCredits        'Primary server telling client the player's current credits

        eSetPlayerRel               'Indicating a relationship change

        eEntityProductionStatus     'used to request and respond regarding the entity production status

        eSetMiningLoc               'client is telling unit to mine somewhere... 

        eSetEntityStatus            'to region server
        eMoveLockViolate            'move lock was violated

        eProductionCompleteNotice   'msg sent to clients that production was finished

        eHangarCargoBayDestroyed    'msg sent from Region to Primary to indicate either (or both) the Hangar/Cargo Bay was destroyed

        eGetAvailableResources      'msg for getting the available resources of an entity

        eDockCommand                'msg from primary to region, special

        eBombardFireMsg             'a special type of fire weapon event
        eRequestOrbitalBombard      'for requesting orbital bombardment
        eStopOrbitalBombard         'requesting to stop orbital bombardment

        eSetEntityAI                'sets the entity's AI settings

        eChatMessage                'a chat message

        eRequestSkillList           'get te skill list for agents
        eRequestGoalList            'get the goal list for agents

        eReloadWpnMsg               'sent to and from primary/domain

        eGetEntityName              'for client caching
        eSetEntityName

        eTransferContents           'for transferring cargo, ammo, fuel, minerals, ships between entities

        eGetColonyDetails           'for getting the colony details

        eRequestPlayerStartLoc      'from primary to region and back... gets a valid starting location for a player's first unit

        eSetEntityPersonnel         'from primary to client

        eServerShutdown             'from primary to region indicating a command to shut it down... from region to primary indicating domain is shut down... from primary to client, indicates server shutdown
        eUpdateEntityAndSave        'from region to primary

        eColonyLowResources         'some resource is low on the colony...

        eBugList                    'from client to primary indicates request for bug list... from primary to client indicates a bug list item
        eBugSubmission              'from client to primary indicates a new bug

        ePFRequestEntitys           'both way...

        eSubmitComponentPrototype       'from client to server
        eComponentPrototypeResponse     'from server to client

        eClientVersion
        eMineralDetailsMsg

        eSetColonyTaxRate

        eGetColonyChildList         'Get colony child list (facilities in the colony)

        eClearDestList              'from region to primary

        eUpdateEntityAttrs          'from region to client

        eUpdateCommandPoints        'from region to client, also serves as a keepalive

        eSetRallyPoint              'from Client to Primary

        eCreateFleet                'from client to primary
        eRemoveFromFleet            'from client to primary
        eAddToFleet                 'from client to primary
        eDeleteFleet                'from client to primary

        eUpdateFleetSpeed
        eSetFleetDest
        eFleetInterSystemMoving

        eUpdateEntityParent

        'Email related
        eAddPlayerCommFolder
        eSendEmail
        eSaveEmailDraft
        eDeleteEmailItem
        eMoveEmailToFolder
        eMarkEmailReadStatus

        ePlayerAlert

        eSubmitTrade

        eGetColonyList
        eGetTradePostTradeables
        eGetNonOwnerDetails         'gets an item's details based on the fact that the player is not the owner

        eRequestEntityDefenses
        eRepairOrder

        eRequestPlayerBudget

        eGetPirateStartLoc
        ePlacePirateAssets

        eUpdateEntityLoc

        eGetPlayerScores

        ePlayerTitleChange

        eGetEnvirConstructionObjects

        eMsgMonitorData

        ePlayerInitialSetup

        eDeathBudgetDeposit

        ePlayerIsDead

        eUpdateResearcherCnt

        eUpdateUnitExpLevel

        eMoveEngineer
        eSetDismantleTarget
        eSetRepairTarget
        eRepairCompleted
        eDeleteDesign

        eMissileFired
        eMissileDetonated
        eMissileImpact

        eRequestEntityAmmo
        eRequestLoadAmmo

        eSendOutMailMsg
        eEmailSettings

        eGetGTCList
        eGetTradeHistory
        eSubmitSellOrder
        eSubmitBuyOrder
        ePurchaseSellOrder
        eGetTradePostList
        eGetOrderSpecifics
        eAcceptBuyOrder
        eDeliverBuyOrder
        eGetTradeDeliveries

        eUpdatePlayerTechValue

        eFinalizeStopEvent  'Sent to and from Region/PF srvrs to determine final stop movements

        eAddPlayerTechKnow

        eSetPlayerSpecialAttribute

        eEntityDefCriticalHitChances

        eSetCounterAttack

        ePlayerDiscoversWormhole
        eJumpTarget
        eSetFleetReinforcer
        ePlayerAliasConfig
        eActOfWarNotice

        eFinalMoveCommand

        eDeinfiltrateAgent
        eSetInfiltrateSettings
        eSubmitMission
        eSetAgentStatus
        eGetAgentStatus
        eGetPMUpdate
        eSetSkipStatus
        eCaptureKillAgent

        eNewsItem

        eCameraPosUpdate
        ePlayerDomainAssignment
        eRemovePlayerRel

        eRequestPlayerDetailsPlayerMin
        eRequestPlayerDetailsAlert
        eRequestMineral
        eRequestEmailSummarys
        eRequestEmail
        eRequestAliasConfigs

        eRequestDXDiag

        eRouteMoveCommand
        eGetRouteList
        eSetRouteMineral
        eRemoveRouteItem
        eUpdateRouteStatus

        eGetSkillList
        eRebuildAISetting

        eAddFormation
        eRemoveFormation
        eMoveFormation
        eSetPrimaryTargetFormation

        eUpdateSlotStates
        eSetIronCurtain
        eGetItemIntelDetail
        eAlertDestinationReached
        eGetIntelSellOrderDetail

        eSetArchiveState
        eGetArchivedItems
        eGetColonyResearchList
        eSetColonyResearchQueue

        eAgentMissionCompleted

        eTutorialProdFinish
        eCreatePlanetInstance
        eSaveAndUnloadInstance
        eUpdatePlayerTutorialStep
        eAILaunchAll

        eBackupOperatorSyncMsg

        eUpdatePlayerTimer
        ePirateWaveSpawn

        eForcedMoveSpeedMove

        eUpdateMOTD
        eGuildBankTransaction
        eGuildRequestDetails
        eUpdatePlayerVote
        eUpdateGuildRank
        eUpdateRankPermission
        eGuildMemberStatus
        eInvitePlayerToGuild
        eUpdateGuildEvent
        eProposeGuildVote
        eSearchGuilds
        eRequestContactWithGuild
        eSetGuildRel
        eAddEventAttachment
        eUpdateGuildRecruitment
        eRemoveEventAttachment
        eRequestGuildVoteProposals
        eAdvancedEventConfig
        eCheckGuildName
        eCreateGuild
        eInviteFormGuild
        eGetGuildInvites
        eGetMyVoteValue
        eRequestGuildEvents
        eUpdateGuildTreasury

        eSetEntityTarget

        eGetGuildAssets
        ePlayerRestart

        eAddSenateProposalMessage
        eAddSenateProposal
        eGetSenateObjectDetails
        eGuildCreationUpdate
        eGuildCreationAcceptance

        eRequestChannelList
        eRequestChannelDetails
        eDeleteTradeHistoryItem
        eSetFleetFormation
        eGetImposedAgentEffects

        eUpdatePlayerDetails
        eUpdatePlayerCustomTitle

        eClearBudgetAlert
        eAbandonColony

        ePDS_AtMissileHit
        ePDS_AtMissileMiss
        ePDS_AtUnitHit
        ePDS_AtUnitMiss

        eTutorialGiveCredits

        eChatChannelCommand

        eForcePrimarySync
        eOperatorRequestPassThru
        eOperatorResponsePassThru
        eUpdatePlanetOwnership

        ePlayerConnectedPrimary
        ePlayerPrimaryOwner
        eForwardToPlayerAtPrimary
        eEntityChangingPrimary
        ePrimaryLoadSharedPlayerData
        eCrossPrimaryBudgetAdjust
        eBudgetResponse

        eRequestServerDetails
        eRequestGlobalPlayerScores
        eMineralBid
        eRequestBidStatus
        eSetMineralBid
        eUndockCommandFinished
        eOutBidAlert
        eRestartTutorial

        eCheckPrimaryReady
        eShiftClickAddProduction
        eClearEntityProdQueue

        ePlayerActivityReport
        eSetSpecialTechThrowback

        ePlaceGuildBillboardBid
        eGetGuildBillboards
        eGetSpecTechGuaranteeList
        eSetSpecTechGuarantee
        eForcefulDismantle

        eClaimItem
        eGetCPPenaltyList
        eSetTetherPoint

        '===================================================
        eLastMsgCode                'MUST ALWAYS BE LAST!!!!
    End Enum

    Public Enum TechBuilderErrorReason As Byte
        eNoError = 0
        eInvalidMaterialsSelected = 1
        eInvalidValuesEntered = 2

        eBeamWeapon_ConcReflectionZero = 10
        eBeamWeapon_CompReflectionZero = 11
        eBeamWeapon_CompQuantumZero = 12
        eBeamWeapon_CompCompressZero = 13
        eBeamWeapon_HousingDensityZero = 14
        eBeamWeapon_ConcReflectionKnowZero = 15
        eBeamWeapon_CatalystCompositeKnowFail = 16
        eBeamWeapon_HousingThermCondZero = 17

        eEngine_StructureBodyDensityZero = 25
        eEngine_StructureBodyMeltingPtZero = 26
        eEngine_StructureFrameMalleableZero = 27
        eEngine_StructureFrameHardnessZero = 28
        eEngine_StructureMeldMeltingPtZero = 29
        eEngine_DriveBodyMagReactZero = 30
        eEngine_DriveFrameSuperCZero = 31
        eEngine_DriveFrameHardnessZero = 32
        eEngine_DriveMeldChemReactZero = 33

        eHull_HardnessZero = 35

        eShield_CoilDensityZero = 45
        eShield_CoilMagProdZero = 46
        eShield_AccQuantumZero = 47
        eShield_AccMagReactZero = 48
        eShield_CasingThermCondZero = 49

        ePulse_CoilDensityZero = 55
        ePulse_CoilSuperCZero = 56
        ePulse_CoilMagReactZero = 57
        ePulse_AccQuantumZero = 58
        ePulse_AccSuperCZero = 59
        ePulse_AccMagReactZero = 60
        ePulse_CasingDensityZero = 61
        ePulse_CasingThermCondZero = 62
        ePulse_CompressChamberDensityZero = 63
        ePulse_CompressChmaberCompressExceedDensity = 64
        ePulse_CoilMagProdExceedCoilMagReact = 65

        eUnknownReason = 255
    End Enum
End Module

Public Class Base_GUID
    'this is a unique id for an object... should be inherited
    ' Every object has a Integer ID as its specific instance ID
    ' every object has a int object type ID as its object type identifier
    'Private mlObjectID As Integer = -1
    Public ObjectID As Int32 = -1
    Public ObjTypeID As Int16
    Protected mbStringReady As Boolean

    Public Sub DataChanged()
        mbStringReady = False
    End Sub

    'Public Property ObjectID() As Integer
    '    Get
    '        Return mlObjectID
    '    End Get
    '    Set(ByVal Value As Integer)
    '        mlObjectID = Value
    '        mbStringReady = False
    '    End Set
    'End Property

    'Public Property ObjTypeID() As Short
    '    Get
    '        Return miObjTypeID
    '    End Get
    '    Set(ByVal Value As Short)
    '        miObjTypeID = Value
    '        mbStringReady = False
    '    End Set
    'End Property

    Public Function GetGUIDAsString() As Byte()
        Dim yMsg(5) As Byte
        System.BitConverter.GetBytes(ObjectID).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(ObjTypeID).CopyTo(yMsg, 4)
        Return yMsg
    End Function

    'Public Sub PopulateFromMsg(ByVal sMsg As String)
    '    'here, we take the contents of sMsg and place them into our values
    '    'sMsg SHOULD be 6 characters Integer (4 bytes for the object id, 2 bytes for the object type id)
    '    Dim sObjID As String
    '    Dim sObjTypeID As String
    '    Dim cData() As Char

    '    If Len(sMsg) = 6 Then
    '        'if it doesn't equal 6, we don't want to assume anything or break out, or anything...
    '        'sObjID = Mid$(sMsg, 1, 4)
    '        'sObjTypeID = Mid$(sMsg, 5, 2)

    '        '
    '        'Now, take those strings and get their ascii vals
    '        'ok, 4 characters representing our numerical value... each one is a byte...

    '        '
    '    End If
    'End Sub
End Class

Public Class InitFile
    ' API functions
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
    Private strFilename As String

    ' Constructor, accepting a filename
    Public Sub New(Optional ByVal Filename As String = "")
        If Filename = "" Then
            'Ok, use the app.path
            strFilename = System.AppDomain.CurrentDomain.BaseDirectory()
            If Right$(strFilename, 1) <> "\" Then strFilename = strFilename & "\"
            strFilename = strFilename & Replace$(System.AppDomain.CurrentDomain.FriendlyName().ToLower, ".exe", ".ini")
        Else
            strFilename = Filename
        End If
    End Sub

    ' Read-only filename property
    ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property

    Public Function GetString(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As String) As String
        ' Returns a string from your INI file
        Dim intCharCount As Integer
        Dim objResult As New System.Text.StringBuilder(2048)
        intCharCount = GetPrivateProfileString(Section, Key, _
           [Default], objResult, objResult.Capacity, strFilename)
        If intCharCount > 0 Then Return Left(objResult.ToString, intCharCount) Else Return ""
    End Function

    Public Function GetInteger(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Integer) As Integer
        ' Returns an integer from your INI file
        Return GetPrivateProfileInt(Section, Key, _
           [Default], strFilename)
    End Function

    Public Function GetBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Boolean) As Boolean
        ' Returns a boolean from your INI file
        Return (GetPrivateProfileInt(Section, Key, _
           CInt([Default]), strFilename) = 1)
    End Function

    Public Sub WriteString(ByVal Section As String, _
      ByVal Key As String, ByVal Value As String)
        ' Writes a string to your INI file
        WritePrivateProfileString(Section, Key, Value, strFilename)
        Flush()
    End Sub

    Public Sub WriteInteger(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Integer)
        ' Writes an integer to your INI file
        WriteString(Section, Key, CStr(Value))
        Flush()
    End Sub

    Public Sub WriteBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Boolean)
        ' Writes a boolean to your INI file
        WriteString(Section, Key, CStr(CInt(Value)))
        Flush()
    End Sub

    Private Sub Flush()
        ' Stores all the cached changes to your INI file
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub

End Class

Public Class StateObject
    'Public workSocket As Socket = Nothing
    Public BufferSize As Int32 = 32767
    Public Buffer(32767) As Byte
    'Public sb As New StringBuilder()
End Class

Public Class NetSock
    'This class is inherited by NetSockClient and NetSockServer
    Public Event onConnect(ByVal Index As Int32)
    Public Event onError(ByVal Index As Int32, ByVal Description As String)
    Public Event onDataArrival(ByVal Index As Int32, ByVal Data As Byte(), ByVal TotalBytes As Int32)
    Public Event onDisconnect(ByVal Index As Int32)
    Public Event onSendComplete(ByVal Index As Int32, ByVal DataSize As Int32)
    Public Event onConnectionRequest(ByVal Index As Int32, ByVal oClient As Socket)

    Public SocketIndex As Int32 = 0     'this is important for when doing arrays of these classes

    Public PortNumber As Int32
    Private moListenThread As Threading.Thread
    Private Const ml_LISTEN_CHECK_INTERVAL As Int32 = 2000      'milliseconds for sleeper
    Private mbListening As Boolean

    Private moSocket As Socket      'the actual socket containment
    Private mbClosing As Boolean = False

    Public ReadOnly Property IsConnected() As Boolean
        Get
            If moSocket Is Nothing = False Then Return moSocket.Connected
            Return False
        End Get
    End Property

    Public Sub SetLocalBinding(ByVal iPort As Int16)
        ''#If Debug = True Then
        'Exit Sub
        ''#Else
        'If moSocket Is Nothing Then moSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

        'Dim oLocalhost As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName)
        'Dim oIP As IPAddress = oLocalhost.AddressList(0)

        ''moSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ExclusiveAddressUse, False)
        'moSocket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, 1)
        'moSocket.Bind(New System.Net.IPEndPoint(oIP, iPort))

        'oIP = Nothing
        'oLocalhost = Nothing
        ''#End If

    End Sub

    Public Sub Connect(ByVal sHostIP As String, ByVal lPortNumber As Int32)
        'this will connect...
        Dim sLoc As String = "Start"
        Try
            Dim oEndPoint As IPEndPoint
            Dim oHost As IPHostEntry
            Dim oAddress As IPAddress = Nothing

            'If moSocket Is Nothing Then
            sLoc = "New"
            moSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            'End If

            'DoLogEvent("Recv: " & moSocket.ReceiveBufferSize & ", Send: " & moSocket.SendBufferSize)
            'moSocket.ReceiveBufferSize = 32767
            'moSocket.SendBufferSize = 32767

            sLoc = "Port"
            PortNumber = lPortNumber
            'oHost = Dns.GetHostByAddress(sHostIP)
            Try
                sLoc = "Host/Addr1"
                oHost = Dns.GetHostByName(sHostIP)
            Catch
                sLoc = "Host/Addr2"
                oHost = Dns.GetHostEntry(sHostIP)
            End Try

            For Each oTmpAddr As IPAddress In oHost.AddressList
                If oTmpAddr.AddressFamily = AddressFamily.InterNetwork Then
                    oAddress = oTmpAddr
                    Exit For
                End If
            Next
            If oAddress Is Nothing Then oAddress = oHost.AddressList(0)

            sLoc = "EP"
            oEndPoint = New IPEndPoint(oAddress, PortNumber)

            sLoc = "BC"
            moSocket.BeginConnect(oEndPoint, AddressOf ConnectedCallback, moSocket)
        Catch ex As SocketException
            Dim sMsg As String = ""
            If ex.InnerException Is Nothing = False Then
                sMsg = ex.InnerException.Message
            End If
            RaiseEvent onError(SocketIndex, ex.Message & ": " & ex.SocketErrorCode & ", I: " & sMsg)
        Catch oldEx As Exception
            Dim sMsg As String = ""
            If oldEx.InnerException Is Nothing = False Then
                sMsg = oldEx.InnerException.Message
            End If
            RaiseEvent onError(SocketIndex, oldEx.Message & ": " & sMsg)
        End Try
    End Sub

    Public Sub SendData(ByVal Data() As Byte)
        Try
            Dim yTemp(Data.Length + 1) As Byte
            System.BitConverter.GetBytes(CShort(Data.Length)).CopyTo(yTemp, 0)
            Data.CopyTo(yTemp, 2)
            moSocket.BeginSend(yTemp, 0, yTemp.Length, 0, AddressOf SendDataCallback, moSocket)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub SendLenAppendedData(ByVal Data() As Byte)
        'Now, this sends the data... but we don't put the length in front of it... it is assumed that Data()
        '  already has the message length in front of it
        Try
            moSocket.BeginSend(Data, 0, Data.Length, 0, AddressOf SendDataCallback, moSocket)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub Disconnect()
        On Error Resume Next

        moSocket.Shutdown(SocketShutdown.Both)
        moSocket.BeginDisconnect(False, AddressOf DisconnectCallBack, moSocket)
    End Sub

    Private Sub DisconnectCallBack(ByVal ar As IAsyncResult)
        On Error Resume Next
        mbClosing = True
        RaiseEvent onDisconnect(SocketIndex)
        moSocket.Close()
        'RaiseEvent onSocketClosed(SocketIndex)
    End Sub

    Private Sub ConnectedCallback(ByVal ar As IAsyncResult)
        Try
            If moSocket.Connected = False Then RaiseEvent onError(SocketIndex, "Connection Refused by Remote Host.") : Exit Sub

            'Dim sTemp As String
            'Dim iEP As IPEndPoint
            'iEP = moSocket.RemoteEndPoint
            'sTemp = "Connected Callback, Remote: " & iEP.Address.ToString() & ":" & iEP.Port & ", Local: "
            'iEP = moSocket.LocalEndPoint
            'sTemp &= iEP.Address.ToString() & ":" & iEP.Port & vbCrLf
            'Debug.Write(sTemp)

            Dim oState As New StateObject()
            moSocket.BeginReceive(oState.Buffer, 0, oState.BufferSize, 0, AddressOf DataArrivalCallback, oState)
            RaiseEvent onConnect(SocketIndex)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Private Sub DataArrivalCallback(ByVal ar As IAsyncResult)
        Dim oState As StateObject = CType(ar.AsyncState, StateObject)
        Dim yTemp() As Byte = Nothing
        Dim lBytesRead As Int32
        Dim lPos As Int32
        Dim iLen As Int16
        Dim iOverrideLen As Int16 = -1

        Static xyLeftover() As Byte
        Static xlLeftoverSize As Int32
        Static xlExpectedLen As Int32
        Static xbPartialLenVal As Boolean

        If mbClosing = True Then Return

        Try
            'Tell the socket to end our receive so we can work, thank you mr. socket for letting us know. have a nice day
			lBytesRead = moSocket.EndReceive(ar)

            'now, that we are here, get our data
            Dim yData() As Byte = oState.Buffer
            If lBytesRead = 0 Then
                moSocket.Shutdown(SocketShutdown.Both)
                moSocket.Close()
                RaiseEvent onDisconnect(SocketIndex)
                Exit Sub
            End If
            ReDim oState.Buffer(32767)

			'DoLogEvent("NetSock.DataArrivalCallBack: " & lBytesRead & " bytes")
            lPos = 0
            While lPos < lBytesRead
				If xlExpectedLen <> 0 Then

					'DoLogEvent("NetSock.DataArrivalCallBack Partial Msg Handled")
					'ok, gotta take care of carryover first
					'Check if it is a partial length
					If xbPartialLenVal = True Then
						'Ok, this is a specific scenario... xlExpectedLen should equal 1 and we should have 1 byte in xyLeftover
						Dim yTmp(1) As Byte
						yTmp(0) = xyLeftover(0)
						yTmp(1) = yData(lPos)
						iOverrideLen = System.BitConverter.ToInt16(yTmp, 0)
					Else
						ReDim yTemp(xyLeftover.Length + xlExpectedLen)
						xyLeftover.CopyTo(yTemp, 0)
						Array.Copy(yData, 0, yTemp, xlLeftoverSize, xlExpectedLen)
					End If
					lPos += xlExpectedLen
					xlExpectedLen = 0
					xbPartialLenVal = False
				Else
					If iOverrideLen < 1 AndAlso lBytesRead - (lPos + 2) < 0 Then
						'Ok, what this is... this is a packet that got cut off at the LEN portion of the message... the length of the message takes two
						'  bytes, if the len gets cut in half (1 byte in one packet, 1 byte in the other) then this occurs
						'So what we do is basically the leftover packet deal...

						'DoLogEvent("NetSock.DataArrivalCallBack Half Len")

						xlLeftoverSize = 1
						ReDim xyLeftover(xlLeftoverSize)
						xlExpectedLen = 1
						xyLeftover(0) = yData(lPos)

						'and we set our boolean indicating partial length
						xbPartialLenVal = True

						lPos += 1
					Else
						'Check for our overridelen
						If iOverrideLen < 1 Then
							iLen = System.BitConverter.ToInt16(yData, lPos)
                            lPos += 2
                        Else
                            iLen = iOverrideLen
                            iOverrideLen = -1
						End If

						If (lBytesRead - lPos) < iLen AndAlso iLen <> 0 Then
							'not enough bytes in this message, set up to wait for next message
							xlLeftoverSize = lBytesRead - lPos
							ReDim xyLeftover(xlLeftoverSize)
							Array.Copy(yData, lPos, xyLeftover, 0, xlLeftoverSize)
							xlExpectedLen = iLen - xlLeftoverSize
							lPos += xlLeftoverSize
							'DoLogEvent("NetSock.DataArrivalCallBack MsgLen: " & iLen & " bytes (partial, leftover: " & xlLeftoverSize)
						Else
							'DoLogEvent("NetSock.DataArrivalCallBack MsgLen: " & iLen & " bytes (all here)")
							'nope, all is here, let's process like normal
							ReDim yTemp(iLen)
							Array.Copy(yData, lPos, yTemp, 0, iLen)
							lPos += iLen
						End If
					End If
				End If

                If xlExpectedLen = 0 AndAlso iOverrideLen < 1 Then RaiseEvent onDataArrival(SocketIndex, yTemp, iLen)
            End While
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
        End Try

        'This new Try...Catch block may hurt performance...
        Try
            'Set up our socket to receive the next data arrival
            If moSocket.Connected = True Then moSocket.BeginReceive(oState.Buffer, 0, oState.BufferSize, 0, AddressOf DataArrivalCallback, oState)
        Catch
            RaiseEvent onError(SocketIndex, "Socket Begin Receive Failed: " & Err.Description)
        End Try

    End Sub

    Private Sub SendDataCallback(ByVal ar As IAsyncResult)
        Try
            'Dim oTmpSock As Socket = CType(ar.AsyncState, Socket)
            If mbClosing = True Then Return
            Dim lBytesSent As Int32 = moSocket.EndSend(ar)
            RaiseEvent onSendComplete(SocketIndex, lBytesSent)
        Catch
            RaiseEvent onError(SocketIndex, Err.Description)
            Exit Sub
        End Try
    End Sub

    Public Sub Listen()
        'ok, here, we will start a new thread to handle listening for clients...
        mbListening = True
        moListenThread = New Threading.Thread(AddressOf BeginListening)
        moListenThread.Start()
    End Sub

    Public Sub StopListening()
        mbListening = False
    End Sub

    'NOTE: This sub should be called from Listen() as a new thread
    Private Sub BeginListening()
        'Create the listener socket
        'Dim oListener As New TcpListener(Dns.GetHostEntry("localhost").AddressList(0), PortNumber)
        Dim oListener As New TcpListener(PortNumber)

        'start listening
        oListener.Start()
        'Now, we will loop while we are listening
        While mbListening
            'Ok, we are listening... check our listener
            While oListener.Pending
                'Ok, we have pending connections...
                RaiseEvent onConnectionRequest(SocketIndex, oListener.AcceptSocket())
            End While

            'Now, that we have connected everyone, let's sleep a while
            Threading.Thread.Sleep(ml_LISTEN_CHECK_INTERVAL)
        End While
    End Sub

    Public Sub MakeReadyToReceive()
        If mbClosing = True Then Return
        Dim oState As New StateObject()
        moSocket.BeginReceive(oState.Buffer, 0, oState.BufferSize, 0, AddressOf DataArrivalCallback, oState)
    End Sub

    Public Sub New()
        'Ok, just your basic construct, no biggy
    End Sub

    Public Sub New(ByVal oSocket As Socket)
        'Ok, now we received a socket that the program already has...
        moSocket = oSocket
    End Sub

    Public Function GetLocalDetails() As EndPoint
        If moSocket Is Nothing = False Then Return moSocket.LocalEndPoint Else Return Nothing
    End Function
    Public Function GetRemoteDetails() As EndPoint
        If moSocket Is Nothing = False Then Return moSocket.RemoteEndPoint Else Return Nothing
    End Function
End Class