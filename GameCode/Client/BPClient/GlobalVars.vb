Option Strict On

Module GlobalVars

	'Public gsPerfString(1) As String
	'Public glPerfStringNum As Int32 = 0
    Public gsDEBUGString As String

    Public goFullScreenBackground As Microsoft.DirectX.Direct3D.Texture

    'Should only be one!!!
    Public goModelDefs As ModelDefs

    'Should only be one!!!
	Public goAvailableResources As AvailResources = Nothing

	'for camera loc sorting via the region
	Public glCameraSectorX As Int32
	Public glCameraSectorZ As Int32

	Public lTotalCache As Int32 = 0
    Public lTotalCacheType1 As Int32 = 0

    Public gfrmMain As frmMain

    'Private moFSLogger As IO.FileStream
    'Private moLogWrite As IO.StreamWriter
    'Public Sub DoLogEvent(ByVal sLine As String)

    '	If moFSLogger Is Nothing Then
    '		Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
    '		If sPath.EndsWith("\") = False Then sPath &= "\"
    '		moFSLogger = New IO.FileStream(sPath & "FullLog.txt", IO.FileMode.Append)
    '	End If
    '	If moLogWrite Is Nothing Then
    '		moLogWrite = New IO.StreamWriter(moFSLogger)
    '		moLogWrite.AutoFlush = True
    '	End If
    '	moLogWrite.WriteLine(sLine)
    'End Sub
    'Public Sub CloseLogForcefully()
    '	If moLogWrite Is Nothing = False Then
    '		moLogWrite.Close()
    '		moLogWrite.Dispose()
    '	End If
    '	moLogWrite = Nothing
    '	If moFSLogger Is Nothing = False Then
    '		moFSLogger.Close()
    '		moFSLogger.Dispose()
    '	End If
    '	moLogWrite = Nothing
    'End Sub

    'Should only be one!!!
    Public goTutorial As TutorialManager = Nothing

#Region "  Current View Management  "
    Public Enum CurrentView As Integer
        'The following are Environment dependant... they are visible based on an environment (or its parents)
        eGalaxyMapView = 0      'the galaxy... can see stars as systems... can scroll around, click on stars, etc...
        eSystemMapView1         'zoom level 1 (1/10000th)
        eSystemMapView2         'zoom level 2 (1/30th but reduced map area)
        eSystemView             'can see the system in 3D with units, facilities, and all
        ePlanetMapView          'can see a map of the planet... not as 3D
        ePlanetView             'can see the planet in 3D with units and facilities, trees, etc...

        'The following are Environment independant... they do not change the current environment or require a specific
        '  environment to be able to be displayed
        eFullScreenInterface    'This must preceed all full screen interfaces

        eDiplomacyScreen        'diplomatic relations interface
        'eSenateScreen           'Galactic Senate options
        'eContractScreen         'player to player contract management
        'eEspionageScreen        'Agents, Agent Missons, Training, etc...
        'eMissionScreen          'A pop-up (IMO) of a current mission... or designing a mission
        'eBudgetScreen           'Left Half is Income, Right Half is Expenses
        'eCorporationsScreen     'List of corps... like looking at a financials database
        'eColonyScreen           'Lists colonies and colony details
        'eProductionScreen       'lists production possibilities and facilities
        'eResearchScreen         'Lists Research Facilities and their projects, empire-wide research points, etc...
        'eMineralDetails         'displays mineral details based on the player's understanding of the mineral
        'eSpecialResearch        'displays a list of special technology researchs available to research or have been
        'eComponentDesign        'allows the user to design componenets
        'eUnitDesign             'allows the user to design units
        'eMilitaryScreen         'shows a list of Unit Groups and values of those groups
        'eGNSScreen              'shows the GNS news coverages
        'eEpicaMailScreen        'shows emails of the Epica game

        eAlloyResearch          'Alloy Research Designer
        eArmorResearch          'Armor Research Designer
        eEngineResearch         'Engine Research Designer
        eRadarResearch          'Radar Research Designer
        eShieldResearch         'Shield Research Designer
        eHullResearch           'Hull Research Designer
        ePrototypeBuilder       'Prototype Designer
        eWeaponBuilder          'Weapon Builder
        eGTCScreen              'trade screen

        eStartupDSELogo         'startup screen that displays the DSE logo
        eStartupLogin           'startup screen that displays the Login screen
    End Enum

    Private mlCurrentEnvirView As Int32    'CurrentView enum
    'Private mlPreviousEnvirView() As Int32
    Private mstckPreviousView As Stack
    Public Property glCurrentEnvirView() As Int32
        Get
            Return mlCurrentEnvirView
        End Get
        Set(ByVal value As Int32)
            Try
                If mlCurrentEnvirView <> CurrentView.ePlanetMapView AndAlso value = CurrentView.ePlanetMapView Then
                    If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                        Dim bSendRequest As Boolean = False
                        If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                            With CType(goCurrentEnvir.oGeoObject, Planet)
                                If .ParentSystem Is Nothing = False AndAlso .ParentSystem.ObjectID = 36 Then
                                    bSendRequest = True
                                End If
                            End With
                        Else
                            With CType(goCurrentEnvir.oGeoObject, Base_GUID)
                                If .ObjectID = 36 AndAlso .ObjTypeID = ObjectType.eSolarSystem Then bSendRequest = True
                            End With
                        End If
                        If bSendRequest = True Then
                            Dim yMsg(7) As Byte
                            Dim lPos As Int32 = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildBillboards).CopyTo(yMsg, lPos) : lPos += 2
                            goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                        End If
                    End If
                  
                End If
            Catch
            End Try

            goCamera.TrackingIndex = -1
            If mstckPreviousView Is Nothing Then mstckPreviousView = New Stack()
            mstckPreviousView.Push(mlCurrentEnvirView)
            'mlPreviousEnvirView = mlCurrentEnvirView
            If mlCurrentEnvirView < CurrentView.eFullScreenInterface AndAlso value < CurrentView.eFullScreenInterface Then mstckPreviousView.Clear()

            BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eChangeCurrentEnvirView, value, mlCurrentEnvirView)

            If mlCurrentEnvirView <> value Then
                mlCurrentEnvirView = value

                Dim sMsg As String = ""
                Select Case mlCurrentEnvirView
                    Case CurrentView.eGalaxyMapView
                        'frmQuickHelp.TriggerFired(frmQuickHelp.elTriggerRequirement.eSwitchToGalaxyMap)
                        sMsg = "Galaxy Map View"
                    Case CurrentView.eSystemMapView1
                        'frmQuickHelp.TriggerFired(frmQuickHelp.elTriggerRequirement.eSwitchToSystemStrategicView)
                        sMsg = "System Strategic View"
                    Case CurrentView.eSystemMapView2
                        'frmQuickHelp.TriggerFired(frmQuickHelp.elTriggerRequirement.eSwitchToSystemTacticalView)
                        sMsg = "System Tactical View"
                    Case CurrentView.eSystemView
                        'frmQuickHelp.TriggerFired(frmQuickHelp.elTriggerRequirement.eSwitchToSystemView)
                        sMsg = "System View"
                    Case CurrentView.ePlanetMapView
                        'frmQuickHelp.TriggerFired(frmQuickHelp.elTriggerRequirement.eSwitchToPlanetMapView)
                        sMsg = "Planet Map View"
                    Case CurrentView.ePlanetView
                        'frmQuickHelp.TriggerFired(frmQuickHelp.elTriggerRequirement.eSwitchToPlanetView)
                        sMsg = "Planet View"
                End Select
                If muSettings.bShowViewMessages = True AndAlso sMsg <> "" Then
                    If goUILib Is Nothing = False Then
                        Dim ofrm As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
                        If ofrm Is Nothing = False Then
                            ofrm.AddImportantMsg(sMsg, System.Drawing.Color.FromArgb(255, 255, 255, 255), 320)
                        End If
                    End If
                End If
            End If
        End Set
    End Property

    Public Sub ReturnToPreviousView()
        Dim lTemp As Int32 = mlCurrentEnvirView
        'mlCurrentEnvirView = mlPreviousEnvirView
        If mstckPreviousView Is Nothing = False AndAlso mstckPreviousView.Count > 0 Then mlCurrentEnvirView = CInt(mstckPreviousView.Pop())
        'mlPreviousEnvirView = lTemp
        If mlCurrentEnvirView < CurrentView.eFullScreenInterface Then mstckPreviousView.Clear()
    End Sub
#End Region

#Region " CONSTANTS AND ENUMERATIONS "
    Public Const gl_MAX_YAW As Int32 = 450
    Public Const gl_MIN_YAW As Int32 = -450

    'For Geometric Calculations
    Public Const gdPi As Single = 3.14159265358979
    Public Const gdHalfPie As Single = gdPi / 2.0F
    Public Const gdPieAndAHalf As Single = gdPi * 1.5F
    Public Const gdTwoPie As Single = gdPi * 2.0F
    Public Const gdDegreePerRad As Single = 180.0F / gdPi
    Public Const gdRadPerDegree As Single = Math.PI / 180.0F

    Public Enum ge_Light_Idx As Integer
        eStar1Light = 0
        eStar2Light
        eStar3Light
    End Enum
    Public Enum eVisibilityType As Byte
        NoVisibility = 0        'do not render at all
        InMaxRange              'render as a sphere? box? something
        Visible                 'render as normal
        FacilityIntel           'was seen at some point but not sure of its status
    End Enum
#End Region

#Region " Settings Management "
    Public Structure EngineSettings
        Public Enum LightQualitySetting As Int32
            VSPS1 = 0
            PerPixel = 1
            BumpMap = 2
        End Enum

#Region " Notification Configuration "
        Public MsgPersonnelProductionComplete As Int32
        Public MsgLandProductionComplete As Int32
        Public MsgNavalProductionComplete As Int32
        Public MsgAerialProductionComplete As Int32
        Public MsgCommandCenterProductionComplete As Int32
        Public MsgSpaceStationProductionComplete As Int32
        Public MsgRefineryProductionComplete As Int32
#End Region

#Region "  Graphics Configuration  "
        Public AmbientLevel As Int32
        Public BumpMapTerrain As Boolean
        Public BumpMapPlanetModel As Boolean
        Public BurnFXParticles As Int32
        Public DepthBuffer As Int32
        Public Dither As Boolean
        Public DrawGrid As Boolean
        Public EngineFXParticles As Int32
        Public EntityClipPlane As Int32
        Public FarClippingPlane As Int32
        Public ScreenshotFormat As Int32
        Public FullScreenRefreshRate As Int32
        Public FullScreenResX As Int32
        Public FullScreenResY As Int32
        Public FOWTextureResolution As Int32
        Public HiResPlanetTexture As Boolean
        Public IlluminationMap As Int32
        Public IlluminationMapTerrain As Boolean
        Public LightQuality As LightQualitySetting
        Public MaxLights As Int32
        Public ModelTextureResolution As eTextureResOptions
        Public NearClippingPlane As Int32
        Public PlanetFXParticles As Int32
        Public PlanetModelTextureWH As Int32
        Public PostGlowAmt As Single
        Public RenderCosmos As Boolean
        Public RenderPlanetRings As Boolean
        Public RenderMineralCaches As Boolean
        'Public RenderPlayerCustomization As Boolean
        Public RenderPulseBolts As Boolean
        Public RenderSpecularMap As Boolean
        Public RenderShieldFX As Boolean
        Public RenderExplosionFX As Boolean
        Public ScreenShakeEnabled As Boolean
        Public ShowFOWTerrainShading As Boolean
        Public ShowMiniMap As Boolean
        Public SmoothFOW As Boolean
        Public SpecularEnabled As Boolean
        Public StarfieldParticlesPlanet As Int32
        Public StarfieldParticlesSpace As Int32
        Public TerrainFarClippingPlane As Int32
        Public TerrainNearClippingPlane As Int32
        Public TerrainTextureResolution As Int32
        Public TripleBuffer As Boolean
        Public SavedTripleBuffer As Boolean
        Public VertexProcessing As Int32
        Public VSync As Boolean
        Public SavedVSync As Boolean
        Public WaterRenderMethod As Int32
        Public WaterTextureRes As Int32
        Public Windowed As Boolean
        Public WormholeAura As Boolean
        Public ExportedDataFormat As Int32
        Public CachePlanetTextures As Boolean
#End Region

#Region "  User Interface Specifics  "
        Public InterfaceBorderColor As System.Drawing.Color '= System.Drawing.Color.FromArgb(255, 255, 255, 255)
        Public InterfaceButtonColor As System.Drawing.Color
        Public InterfaceFillColor As System.Drawing.Color '= System.Drawing.Color.FromArgb(128, 32, 64, 128)
        Public InterfaceTextBoxFillColor As System.Drawing.Color
        Public InterfaceTextBoxForeColor As System.Drawing.Color
        Public NotificationDisplayTime As Int32

        Public FilterBadWords As Boolean
        Public ChatTimeStamps As Boolean
        Public bFormationMoveThenForm As Boolean
#End Region

#Region "  Window Position and Other Settings  "
        Public AdvanceDisplayLocX As Int32
        Public AdvanceDisplayLocY As Int32
        Public BehaviorLocX As Int32
        Public BehaviorLocY As Int32
        Public BuildWindowLeft As Int32
        Public BuildWindowTop As Int32
        Public ChatWindowLocX As Int32
        Public ChatWindowLocY As Int32
        Public ChatWindowWidth As Int32
        Public ChatWindowHeight As Int32
        Public ChatWindowState As Byte
        Public ContentsLocX As Int32
        Public ContentsLocY As Int32
        Public EnvirDisplayLocX As Int32
        Public EnvirDisplayLocY As Int32
        Public MiniMapLocX As Int32
        Public MiniMapLocY As Int32
        Public MiniMapWidthHeight As Int32
        Public MiningWindowX As Int32
        Public MiningWindowY As Int32
        Public PlanetMinimapZoomLevel As Byte
        Public ProdStatusLocX As Int32
        Public ProdStatusLocY As Int32
        Public ProdStatusWidth As Int32
        Public ResearchWindowLeft As Int32
        Public ResearchWindowTop As Int32
        Public ResearchWindowStaticExpand As Boolean
        Public ColonialManagementLeft As Int32
        Public ColonialManagementTop As Int32
        Public ColonialManagementHeight As Int32
        Public SpecialTechLeft As Int32
        Public SpecialTechTop As Int32
        Public ShowConnectionStatus As Boolean


        Public SelectLocX As Int32
        Public SelectLocY As Int32

        Public MultiSelectLocX As Int32
        Public MultiSelectLocY As Int32
        Public MultiSelectWidth As Int32
        Public MultiSelectHeight As Int32
        Public ShowMultiHealthBars As Boolean

        Public TutorialTOCX As Int32
        Public TutorialTOCY As Int32
        Public EmailMainX As Int32
        Public EmailMainY As Int32
        Public FleetX As Int32
        Public FleetY As Int32
        Public TradeMainX As Int32
        Public TradeMainY As Int32
        Public DiplomacyX As Int32
        Public DiplomacyY As Int32
        Public ColonyStatsX As Int32
        Public ColonyStatsY As Int32
        Public BudgetX As Int32
        Public BudgetY As Int32
        Public BudgetWindowStaticExpand As Boolean
        Public AgentMainX As Int32
        Public AgentMainY As Int32
        Public AgentMissionCreateX As Int32
        Public AgentMissionCreateY As Int32
        Public AgentMissionDetailsX As Int32
        Public AgentMissionDetailsY As Int32
        Public AgentDetailsX As Int32
        Public AgentDetailsY As Int32
        Public FormationsX As Int32
        Public FormationsY As Int32
        Public GuildMainX As Int32
        Public GuildMainY As Int32
        Public AvailResX As Int32
        Public AvailResY As Int32
        Public CommandX As Int32
        Public CommandY As Int32
        Public GalaxyControlX As Int32
        Public GalaxyControlY As Int32
        Public ArenaMainX As Int32
        Public ArenaMainY As Int32
        Public ArenaConfigX As Int32
        Public ArenaConfigY As Int32
        Public ArenaWaitX As Int32
        Public ArenaWaitY As Int32
        Public WarCalendarX As Int32
        Public WarCalendarY As Int32
        Public UnitGotoLocX As Int32
        Public UnitGotoLocY As Int32
        Public BuilderWorksheetX As Int32
        Public BuilderWorksheetY As Int32
        Public ObserveX As Int32
        Public ObserveY As Int32
        Public EnvirRelationsX As Int32
        Public EnvirRelationsY As Int32
        Public TransportManagementX As Int32
        Public TransportManagementY As Int32
        Public TransportOrdersX As Int32
        Public TransportOrdersY As Int32


#End Region

#Region "  Chat Window Colors  "
        Public DefaultChatColor As System.Drawing.Color
        Public AlertChatColor As System.Drawing.Color
        Public StatusChatColor As System.Drawing.Color
        Public LocalChatColor As System.Drawing.Color
        Public GuildChatColor As System.Drawing.Color
        Public SenateChatColor As System.Drawing.Color
        Public PMChatColor As System.Drawing.Color
        Public ChannelChatColor As System.Drawing.Color
        'Public Ch2ChatColor As System.Drawing.Color
        'Public Ch3ChatColor As System.Drawing.Color
        'Public Ch4ChatColor As System.Drawing.Color
        'Public Ch5ChatColor As System.Drawing.Color
        'Public Ch6ChatColor As System.Drawing.Color
        'Public Ch7ChatColor As System.Drawing.Color
        'Public Ch8ChatColor As System.Drawing.Color
        'Public Ch9ChatColor As System.Drawing.Color
        'Public Ch10ChatColor As System.Drawing.Color
#End Region

#Region "  Planet Type Color Config  "
        Public AcidBuildGhost As System.Drawing.Color
        Public AcidMineralCache As System.Drawing.Color
        Public AcidMinimapAngle As System.Drawing.Color
        Public AdaptableBuildGhost As System.Drawing.Color
        Public AdaptableMineralCache As System.Drawing.Color
        Public AdaptableMinimapAngle As System.Drawing.Color
        Public BarrenBuildGhost As System.Drawing.Color
        Public BarrenMineralCache As System.Drawing.Color
        Public BarrenMinimapAngle As System.Drawing.Color
        Public DesertBuildGhost As System.Drawing.Color
        Public DesertMineralCache As System.Drawing.Color
        Public DesertMinimapAngle As System.Drawing.Color
        Public IceBuildGhost As System.Drawing.Color
        Public IceMineralCache As System.Drawing.Color
        Public IceMinimapAngle As System.Drawing.Color
        Public LavaBuildGhost As System.Drawing.Color
        Public LavaMineralCache As System.Drawing.Color
        Public LavaMinimapAngle As System.Drawing.Color
        Public TerranBuildGhost As System.Drawing.Color
        Public TerranMineralCache As System.Drawing.Color
        Public TerranMinimapAngle As System.Drawing.Color
        Public WaterworldBuildGhost As System.Drawing.Color
        Public WaterworldMineralCache As System.Drawing.Color
        Public WaterworldMinimapAngle As System.Drawing.Color
#End Region

#Region "  Asset Colorization  "
        Public NeutralAssetColor As Microsoft.DirectX.Vector4
        Public MyAssetColor As Microsoft.DirectX.Vector4
        Public EnemyAssetColor As Microsoft.DirectX.Vector4
        Public AllyAssetColor As Microsoft.DirectX.Vector4
        Public GuildAssetColor As Microsoft.DirectX.Vector4
        Public TacticalAssetColor As Microsoft.DirectX.Vector4
#End Region

        Public RewardLabelFadeRate As Single
        Public RewardLabelRiseRate As Single

        Public AudioOn As Boolean
        Public PositionalSound As Boolean
        Public bRanBefore As Boolean

        Public ShowIntro As Boolean

        Public ZoomChangesView As Boolean
        Public CtrlQExits As Boolean
        Public ShowTargetBoxes As Boolean
        Public RenderHPBars As Boolean
        Public RenderBurnMarks As Boolean

        Public FlashSelections As Boolean
        Public FlashRate As Int32

        Public bDoNotShowEngineerCancelAlert As Boolean

        Public bShowViewMessages As Boolean

        Public lDYK_Frequency As Int32

        Public lCurrentResearchFilter As Int32          'NOT SAVED OR LOADED
        Public lPlanetViewCameraX As Int32              'NOT SAVED OR LOADED
        Public lPlanetViewCameraY As Int32              'NOT SAVED OR LOADED
        Public lPlanetViewCameraZ As Int32              'NOT SAVED OR LOADED
        Public lGalaxyViewCameraX As Int32              'NOT SAVED OR LOADED
        Public lGalaxyViewCameraY As Int32              'NOT SAVED OR LOADED
        Public lGalaxyViewCameraZ As Int32              'NOT SAVED OR LOADED
        Public yCurrentContentsView As Byte             'NOT SAVED OR LOADED
        Public ExpandedColonyStatsScreen As Boolean     'Not saved or loaded

        Public gbGalaxyControlHideWormholes As Boolean      'NOT SAVED OR LOADED
        Public gbGalaxyControlHideFleetMovement As Boolean  'NOT SAVED OR LOADED
        Public gbGalaxyControlStarLabelFalloff As Boolean   'Saved

        Public gbSolarMapDontShowPlanets As Boolean         'NOT SAVED OR LOADED
        Public gbSolarMapDontShowStars As Boolean          'NOT SAVED OR LOADED
        Public gbSolarMapDontShowWormholes As Boolean       'NOT SAVED OR LOADED

        Public gbPlanetMapDontRenderTerrain As Boolean      'NOT SAVED OR LOADED
        Public gbPlanetMapBlinkUnits As Boolean             'Saved
        Public gbPlanetMapDontShowMyUnits As Boolean        'NOT SAVED OR LOADED
        Public gbPlanetMapDontShowGuilds As Boolean         'NOT SAVED OR LOADED
        Public gbPlanetMapDontShowAllied As Boolean         'NOT SAVED OR LOADED
        Public gbPlanetMapDontShowNeutral As Boolean        'NOT SAVED OR LOADED
        Public gbPlanetMapDontShowEnemy As Boolean          'NOT SAVED OR LOADED
        Public gbPlanetMapDontShowUnknown As Boolean        'NOT SAVED OR LOADED


        Public gbDisableExclamations As Boolean 'Saved

        Public Enum eTextureResOptions As Integer
            eHiResTextures = 512
            eNormResTextures = 256
            eLowResTextures = 128
            eVeryLowResTextures = 64
        End Enum

        Public Sub LoadSettings()
            Dim oINI As New InitFile()

            Dim lA As Int32
            Dim lR As Int32
            Dim lG As Int32
            Dim lB As Int32

            lDYK_Frequency = CInt(Val(oINI.GetString("SETTINGS", "DYK_Frequency", "1")))

            bFormationMoveThenForm = CInt(Val(oINI.GetString("SETTINGS", "MoveThenForm", "0"))) <> 0

            bShowViewMessages = CBool(Val(oINI.GetString("SETTINGS", "ShowViewMessages", "1")) <> 0)

            Windowed = CBool(Val(oINI.GetString("GRAPHICS", "Windowed", "1")) <> 0)
            VertexProcessing = CInt(Val(oINI.GetString("GRAPHICS", "VertexProcessing", "64")))   'hardware
            TripleBuffer = CBool(Val(oINI.GetString("GRAPHICS", "TripleBuffer", "0")) <> 0)
            SavedTripleBuffer = TripleBuffer
            VSync = CBool(Val(oINI.GetString("GRAPHICS", "VSync", "0")) <> 0)
            SavedVSync = VSync
            FullScreenResX = CInt(oINI.GetString("GRAPHICS", "FullScreenResX", "-1"))
            FullScreenResY = CInt(oINI.GetString("GRAPHICS", "FullScreenResY", "-1"))
            FullScreenRefreshRate = CInt(oINI.GetString("GRAPHICS", "FullScreenRefreshRate", "-1"))

            ZoomChangesView = CBool(Val(oINI.GetString("SETTINGS", "ZoomChangesView", "0")) <> 0)

            RewardLabelFadeRate = CSng(oINI.GetString("SETTINGS", "RewardLabelFadeRate", "5"))
            RewardLabelRiseRate = CSng(oINI.GetString("SETTINGS", "RewardLabelRiseRate", "15"))

            FlashSelections = CBool(CInt(oINI.GetString("SETTINGS", "FlashSelections", "1")) <> 0)
            FlashRate = CInt(oINI.GetString("SETTINGS", "FlashRate", "5"))

            WormholeAura = CInt(oINI.GetString("GRAPHICS", "WormholeAura", "1")) <> 0

            AmbientLevel = Math.Max(0, Math.Min(50, CInt(oINI.GetString("GRAPHICS", "AmbientLevel", "19"))))

            ChatTimeStamps = CInt(Val(oINI.GetString("INTERFACE", "ChatTimeStamps", "1"))) <> 0
            FilterBadWords = CInt(Val(oINI.GetString("INTERFACE", "FilterBadWords", "1"))) <> 0
            RenderHPBars = CInt(Val(oINI.GetString("SETTINGS", "RenderHPBars", "1"))) <> 0
            RenderBurnMarks = CInt(Val(oINI.GetString("GRAPHICS", "RenderBurnMarks", "1"))) <> 0
            RenderPulseBolts = CInt(Val(oINI.GetString("GRAPHICS", "RenderPulseBolts", "1"))) <> 0

            bDoNotShowEngineerCancelAlert = CInt(Val(oINI.GetString("INTERFACE", "DoNotShowEngineerCancelAlert", "0"))) <> 0
            CtrlQExits = CInt(Val(oINI.GetString("SETTINGS", "CtrlQExits", "1"))) <> 0

            LightQuality = CType(CInt(Val(oINI.GetString("LIGHTING", "Quality", CInt(LightQualitySetting.BumpMap).ToString))), LightQualitySetting)
            RenderSpecularMap = CBool(CInt(Val(oINI.GetString("GRAPHICS", "SpecularMap", "1"))) <> 0)
            IlluminationMap = Math.Min(Math.Max(64, CInt(Val(oINI.GetString("LIGHTING", "IlluminationMap", "512")))), 1024)
            RenderCosmos = CInt(Val(oINI.GetString("GRAPHICS", "RenderCosmos", "1"))) <> 0
            RenderPlanetRings = CInt(Val(oINI.GetString("GRAPHICS", "RenderPlanetRings", "1"))) <> 0
            StarfieldParticlesPlanet = Math.Max(100, Math.Min(4000, CInt(Val(oINI.GetString("GRAPHICS", "StarfieldParticlesPlanet", 500.ToString)))))
            StarfieldParticlesSpace = Math.Max(1000, Math.Min(40000, CInt(Val(oINI.GetString("GRAPHICS", "StarfieldParticlesSpace", 10000.ToString)))))
            BumpMapTerrain = CInt(Val(oINI.GetString("LIGHTING", "BumpMapTerrain", "1"))) <> 0
            BumpMapPlanetModel = CInt(Val(oINI.GetString("LIGHTING", "BumpMapPlanetModel", "1"))) <> 0
            IlluminationMapTerrain = CInt(Val(oINI.GetString("LIGHTING", "IllumMapTerrain", "1"))) <> 0

            'NeutralAssetColor = New Microsoft.DirectX.Vector4(1, 1, 1, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "NeutralR", "255")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "NeutralG", "255")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "NeutralB", "255")))
            lA = 0
            NeutralAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'MyAssetColor = New Microsoft.DirectX.Vector4(0, 1, 0, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "PlayerR", "0")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "PlayerG", "255")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "PlayerB", "0")))
            lA = 0
            MyAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'EnemyAssetColor = New Microsoft.DirectX.Vector4(0, 1, 0, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "EnemyR", "255")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "EnemyG", "0")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "EnemyB", "0")))
            lA = 0
            EnemyAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'AllyAssetColor = New Microsoft.DirectX.Vector4(0, 1, 1, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "AllyR", "0")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "AllyG", "255")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "AllyB", "255")))
            lA = 0
            AllyAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "GuildR", "192")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "GuildG", "32")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "GuildB", "128")))
            lA = 0
            GuildAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "TacticalR", "192")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "TacticalG", "192")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "TacticalB", "32")))
            lA = 0
            TacticalAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'DefaultChatColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultB", "255")))
            DefaultChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'AlertChatColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "AlertR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "AlertG", "0")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "AlertB", "0")))
            AlertChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'StatusChatColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "StatusR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "StatusG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "StatusB", "0")))
            StatusChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'LocalChatColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "LocalR", "0")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "LocalG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "LocalB", "255")))
            LocalChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'GuildChatColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "GuildR", "0")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "GuildG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "GuildB", "0")))
            GuildChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'PMChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "PMR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "PMG", "128")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "PMB", "0")))
            PMChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'SenateChatColor = System.Drawing.Color.FromArgb(255, 192, 192, 192)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "SenateR", "192")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "SenateG", "192")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "SenateB", "192")))
            SenateChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'ChannelChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelG", "128")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelB", "255")))
            ChannelChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'Ch2ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch3ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch4ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch5ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch6ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch7ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch8ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch9ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch10ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "AcidR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "AcidG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "AcidB", "255")))
            AcidBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "AcidR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "AcidG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "AcidB", "255")))
            AcidMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "AcidR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "AcidG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "AcidB", "255")))
            AcidMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "AdaptableR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "AdaptableG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "AdaptableB", "255")))
            AdaptableBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "AdaptableR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "AdaptableG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "AdaptableB", "255")))
            AdaptableMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "AdaptableR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "AdaptableG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "AdaptableB", "255")))
            AdaptableMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "BarrenR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "BarrenG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "BarrenB", "255")))
            BarrenBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "BarrenR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "BarrenG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "BarrenB", "255")))
            BarrenMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "BarrenR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "BarrenG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "BarrenB", "255")))
            BarrenMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "DesertR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "DesertG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "DesertB", "255")))
            DesertBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "DesertR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "DesertG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "DesertB", "255")))
            DesertMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "DesertR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "DesertG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "DesertB", "255")))
            DesertMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "IceR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "IceG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "IceB", "255")))
            IceBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "IceR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "IceG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "IceB", "255")))
            IceMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "IceR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "IceG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "IceB", "255")))
            IceMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "LavaR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "LavaG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "LavaB", "255")))
            LavaBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "LavaR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "LavaG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "LavaB", "255")))
            LavaMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "LavaR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "LavaG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "LavaB", "255")))
            LavaMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "TerranR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "TerranG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "TerranB", "255")))
            TerranBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "TerranR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "TerranG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "TerranB", "255")))
            TerranMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "TerranR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "TerranG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "TerranB", "255")))
            TerranMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "WaterworldR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "WaterworldG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "WaterworldB", "255")))
            WaterworldBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "WaterworldR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "WaterworldG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "WaterworldB", "255")))
            WaterworldMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "WaterworldR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "WaterworldG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "WaterworldB", "255")))
            WaterworldMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)


            'ALWAYS SET TO THIS ON EVERY STARTUP
            lCurrentResearchFilter = -1
            lPlanetViewCameraX = 0
            lPlanetViewCameraY = 1700
            lPlanetViewCameraZ = 1000
            ExpandedColonyStatsScreen = True
            '====== end of same startup =======

            ShowIntro = CBool(Val(oINI.GetString("SETTINGS", "ShowIntroScreen", "1")) <> 0)
            ShowTargetBoxes = CBool(Val(oINI.GetString("SETTINGS", "ShowTargetBoxes", "1")) <> 0)
            SmoothFOW = CBool(Val(oINI.GetString("GRAPHICS", "SmoothFOW", "1")) <> 0)
            WaterRenderMethod = CInt(Val(oINI.GetString("GRAPHICS", "WaterRenderMethod", "1")))
            WaterTextureRes = Math.Min(256, Math.Max(64, CInt(Val(oINI.GetString("GRAPHICS", "WaterTextureRes", "256")))))
            PlanetMinimapZoomLevel = CByte(Math.Min(3, Math.Max(0, Val(oINI.GetString("SETTINGS", "PlanetMinimapZoomLevel", "3")))))
            MaxLights = CInt(Val(oINI.GetString("GRAPHICS", "MaxLights", "0")))
            FarClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "FarClippingPlane", "1000000000")))
            NearClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "NearClippingPlane", "1")))
            TerrainNearClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "TerrainNearClippingPlane", "100")))
            TerrainFarClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "TerrainFarClippingPlane", "25000")))
            ShowMiniMap = CBool(Val(oINI.GetString("PLANET_VIEW", "ShowMiniMap", "1")))
            gbPlanetMapBlinkUnits = CBool(Val(oINI.GetString("PLANET_VIEW", "BlinkUnits", "1")))
            gbGalaxyControlStarLabelFalloff = CBool(Val(oINI.GetString("GALAXY_VIEW", "LabelFalloff", "1")))
            EntityClipPlane = CInt(Val(oINI.GetString("GRAPHICS", "EntityClipPlane", "20000")))
            ModelTextureResolution = CType((Math.Min(512, Math.Max(64, Val(oINI.GetString("GRAPHICS", "ModelTextureResolution", CInt(eTextureResOptions.eNormResTextures).ToString()))))), eTextureResOptions)
            If CInt(ModelTextureResolution) = 0 Then ModelTextureResolution = eTextureResOptions.eNormResTextures
            DrawGrid = CBool(Val(oINI.GetString("GRAPHICS", "DrawGrid", "0")))
            'DrawProceduralWater = CBool(Val(oINI.GetString("GRAPHICS", "DrawProceduralWater", "0")))
            ShowFOWTerrainShading = CBool(Val(oINI.GetString("GRAPHICS", "ShowFOWTerrainShading", "1")))
            FOWTextureResolution = CInt(Val(oINI.GetString("GRAPHICS", "FOWTextureResolution", "512")))
            ScreenshotFormat = CInt(Val(oINI.GetString("GRAPHICS", "ScreenshotFormat", "0")))
            If FOWTextureResolution < 256 Then FOWTextureResolution = 256
            MiniMapLocX = CInt(Val(oINI.GetString("HUD", "MiniMapLocX", "0")))
            MiniMapLocY = CInt(Val(oINI.GetString("HUD", "MiniMapLocY", "0")))
            MiniMapWidthHeight = CInt(Val(oINI.GetString("HUD", "MiniMapWidthHeight", "120")))
            RenderMineralCaches = CBool(Val(oINI.GetString("GRAPHICS", "RenderMineralCaches", "1")))
            Dither = CBool(Val(oINI.GetString("GRAPHICS", "DITHER", "0")))
            SpecularEnabled = CBool(Val(oINI.GetString("GRAPHICS", "SpecularEnabled", "1")))
            NotificationDisplayTime = CInt(Val(oINI.GetString("SETTINGS", "NotificationDisplayTime", "5000")))
            AudioOn = CBool(Val(oINI.GetString("AUDIO", "AudioEnabled", "1")))
            PositionalSound = CBool(Val(oINI.GetString("AUDIO", "PositionalAudio", "0")))
            'WaterQuadWidthHeight = CInt(Val(oINI.GetString("GRAPHICS", "WaterQuadRes", "32")))
            ScreenShakeEnabled = CBool(Val(oINI.GetString("SETTINGS", "ScreenShakeEnable", "1")) <> 0)
            TerrainTextureResolution = Math.Max(1, CInt(Val(oINI.GetString("GRAPHICS", "TerrainTextureResolution", "3"))))
            PlanetModelTextureWH = Math.Max(64, CInt(Val(oINI.GetString("GRAPHICS", "PlanetModelTextureResolution", "256"))))

            BuildWindowLeft = CInt(Val(oINI.GetString("HUD", "BuildWindowLeft", "-1")))
            BuildWindowTop = CInt(Val(oINI.GetString("HUD", "BuildWindowTop", "-1")))
            ResearchWindowLeft = CInt(Val(oINI.GetString("HUD", "ResearchWindowLeft", "-1")))
            ResearchWindowTop = CInt(Val(oINI.GetString("HUD", "ResearchWindowTop", "-1")))
            ColonialManagementLeft = CInt(Val(oINI.GetString("HUD", "ColonialManagementLeft", "-1")))
            ColonialManagementTop = CInt(Val(oINI.GetString("HUD", "ColonialManagementTop", "-1")))
            ColonialManagementHeight = CInt(Val(oINI.GetString("HUD", "ColonialManagementHeight", "-1")))
            SpecialTechLeft = CInt(Val(oINI.GetString("HUD", "SpecialTechLeft", "-1")))
            SpecialTechTop = CInt(Val(oINI.GetString("HUD", "SpecialTechTop", "-1")))

            EngineFXParticles = CInt(Val(oINI.GetString("GRAPHICS", "EngineFXParticles", "4")))
            If EngineFXParticles > 4 OrElse EngineFXParticles < 0 Then EngineFXParticles = 4
            BurnFXParticles = CInt(Val(oINI.GetString("GRAPHICS", "BurnFXParticles", "4")))
            If BurnFXParticles > 4 OrElse BurnFXParticles < 0 Then BurnFXParticles = 4
            RenderShieldFX = CBool(Val(oINI.GetString("GRAPHICS", "RenderShieldFX", "1")) <> 0)
            'StarfieldParticlesSpace = CInt(Val(oINI.GetString("GRAPHICS", "StarfieldParticles", "10000")))
            RenderExplosionFX = CBool(Val(oINI.GetString("GRAPHICS", "RenderExplosionFX", "1")) <> 0)
            If StarfieldParticlesSpace < 1000 Then StarfieldParticlesSpace = 1000
            If StarfieldParticlesSpace > 40000 Then StarfieldParticlesSpace = 40000
            PlanetFXParticles = CInt(Val(oINI.GetString("GRAPHICS", "PlanetFXParticles", "4")))
            If PlanetFXParticles > 4 OrElse PlanetFXParticles < 0 Then PlanetFXParticles = 4

            PostGlowAmt = CSng(Val(oINI.GetString("GRAPHICS", "PostGlowAmount", "0.5")))
            HiResPlanetTexture = CBool(Val(oINI.GetString("GRAPHICS", "HiResPlanetTexture", "1")) <> 0)

            bRanBefore = CBool(Val(oINI.GetString("SETTINGS", "RanBefore", "0")) <> 0)
            gbDisableExclamations = CBool(Val(oINI.GetString("SETTINGS", "DisableExclamations", "0")) <> 0)
            ExportedDataFormat = CInt(Val(oINI.GetString("SETTINGS", "ExportedDataFormat", "1")))
            '(Disabled for now) CachePlanetTextures = CBool(Val(oINI.GetString("SETTINGS", "CachePlanetTextures", "0")) <> 0)

            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_R", "255"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_G", "255"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_B", "255"))), 0), 255)
            InterfaceBorderColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_A", "128"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_R", "32"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_G", "64"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_B", "128"))), 0), 255)
            InterfaceFillColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_R", "32"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_G", "64"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_B", "92"))), 0), 255)
            InterfaceTextBoxFillColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_R", "255"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_G", "255"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_B", "255"))), 0), 255)
            InterfaceTextBoxForeColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_R", "192"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_G", "220"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_B", "255"))), 0), 255)
            InterfaceButtonColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)

            'Hud Position Settings
            AdvanceDisplayLocX = CInt(Val(oINI.GetString("INTERFACE", "AdvanceDisplayLocX", "-1")))
            AdvanceDisplayLocY = CInt(Val(oINI.GetString("INTERFACE", "AdvanceDisplayLocY", "-1")))
            BehaviorLocX = CInt(Val(oINI.GetString("INTERFACE", "BehaviorLocX", "-1")))
            BehaviorLocY = CInt(Val(oINI.GetString("INTERFACE", "BehaviorLocY", "-1")))
            ContentsLocX = CInt(Val(oINI.GetString("INTERFACE", "ContentsLocX", "-1")))
            ContentsLocY = CInt(Val(oINI.GetString("INTERFACE", "ContentsLocY", "-1")))
            EnvirDisplayLocX = CInt(Val(oINI.GetString("INTERFACE", "EnvirDisplayLocX", "-1")))
            EnvirDisplayLocY = CInt(Val(oINI.GetString("INTERFACE", "EnvirDisplayLocY", "-1")))
            SelectLocX = CInt(Val(oINI.GetString("INTERFACE", "SelectLocX", "-1")))
            SelectLocY = CInt(Val(oINI.GetString("INTERFACE", "SelectLocY", "-1")))
            MultiSelectLocX = CInt(Val(oINI.GetString("INTERFACE", "MultiSelectLocX", SelectLocX.ToString)))
            MultiSelectLocY = CInt(Val(oINI.GetString("INTERFACE", "MultiSelectLocY", SelectLocY.ToString)))
            MultiSelectWidth = CInt(Val(oINI.GetString("INTERFACE", "MultiSelectWidth", "-1")))
            MultiSelectHeight = CInt(Val(oINI.GetString("INTERFACE", "MultiSelectHeight", "-1")))
            ProdStatusLocX = CInt(Val(oINI.GetString("INTERFACE", "ProdStatusLocX", "-1")))
            ProdStatusLocY = CInt(Val(oINI.GetString("INTERFACE", "ProdStatusLocY", "-1")))
            ProdStatusWidth = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ProdStatusWidth", "-1"))), 200), 500)
            ChatWindowLocX = CInt(Val(oINI.GetString("INTERFACE", "ChatWindowLocX", "-1")))
            ChatWindowLocY = CInt(Val(oINI.GetString("INTERFACE", "ChatWindowLocY", "-1")))
            ChatWindowWidth = CInt(Val(oINI.GetString("INTERFACE", "ChatWindowWidth", "-1")))
            ChatWindowHeight = CInt(Val(oINI.GetString("INTERFACE", "ChatWindowHeight", "-1")))
            ChatWindowState = CByte(oINI.GetString("INTERFACE", "ChatWindowState", "1"))

            TutorialTOCX = CInt(Val(oINI.GetString("INTERFACE", "TutorialTOCX", "-1")))
            TutorialTOCY = CInt(Val(oINI.GetString("INTERFACE", "TutorialTOCY", "-1")))
            EmailMainX = CInt(Val(oINI.GetString("INTERFACE", "EmailMainX", "-1")))
            EmailMainY = CInt(Val(oINI.GetString("INTERFACE", "EmailMainY", "-1")))
            FleetX = CInt(Val(oINI.GetString("INTERFACE", "FleetX", "-1")))
            FleetY = CInt(Val(oINI.GetString("INTERFACE", "FleetY", "-1")))
            TradeMainX = CInt(Val(oINI.GetString("INTERFACE", "TradeMainX", "-1")))
            TradeMainY = CInt(Val(oINI.GetString("INTERFACE", "TradeMainY", "-1")))
            DiplomacyX = CInt(Val(oINI.GetString("INTERFACE", "DiplomacyX", "-1")))
            DiplomacyY = CInt(Val(oINI.GetString("INTERFACE", "DiplomacyY", "-1")))
            ColonyStatsX = CInt(Val(oINI.GetString("INTERFACE", "ColonyStatsX", "-1")))
            ColonyStatsY = CInt(Val(oINI.GetString("INTERFACE", "ColonyStatsY", "-1")))
            BudgetX = CInt(Val(oINI.GetString("INTERFACE", "BudgetX", "-1")))
            BudgetY = CInt(Val(oINI.GetString("INTERFACE", "BudgetY", "-1")))
            AgentMainX = CInt(Val(oINI.GetString("INTERFACE", "AgentMainX", "-1")))
            AgentMainY = CInt(Val(oINI.GetString("INTERFACE", "AgentMainY", "-1")))
            ArenaMainX = CInt(Val(oINI.GetString("INTERFACE", "ArenaMainX", "-1")))
            ArenaMainY = CInt(Val(oINI.GetString("INTERFACE", "ArenaMainY", "-1")))
            ArenaConfigX = CInt(Val(oINI.GetString("INTERFACE", "ArenaConfigX", "-1")))
            ArenaConfigY = CInt(Val(oINI.GetString("INTERFACE", "ArenaConfigX", "-1")))
            ArenaWaitX = CInt(Val(oINI.GetString("INTERFACE", "ArenaWaitX", "-1")))
            ArenaWaitY = CInt(Val(oINI.GetString("INTERFACE", "ArenaWaitY", "-1")))
            AgentMissionCreateX = CInt(Val(oINI.GetString("INTERFACE", "AgentMissionCreateX", "-1")))
            AgentMissionCreateY = CInt(Val(oINI.GetString("INTERFACE", "AgentMissionCreateY", "-1")))
            AgentMissionDetailsX = CInt(Val(oINI.GetString("INTERFACE", "AgentMissionDetailsX", "-1")))
            AgentMissionDetailsY = CInt(Val(oINI.GetString("INTERFACE", "AgentMissionDetailsY", "-1")))
            AgentDetailsX = CInt(Val(oINI.GetString("INTERFACE", "AgentDetailsX", "-1")))
            AgentDetailsY = CInt(Val(oINI.GetString("INTERFACE", "AgentDetailsY", "-1")))
            FormationsX = CInt(Val(oINI.GetString("INTERFACE", "FormationsX", "-1")))
            FormationsY = CInt(Val(oINI.GetString("INTERFACE", "FormationsY", "-1")))
            GuildMainX = CInt(Val(oINI.GetString("INTERFACE", "GuildMainX", "-1")))
            GuildMainY = CInt(Val(oINI.GetString("INTERFACE", "GuildMainY", "-1")))
            AvailResX = CInt(Val(oINI.GetString("INTERFACE", "AvailResX", "-1")))
            AvailResY = CInt(Val(oINI.GetString("INTERFACE", "AvailResY", "-1")))
            CommandX = CInt(Val(oINI.GetString("INTERFACE", "CommandX", "-1")))
            CommandY = CInt(Val(oINI.GetString("INTERFACE", "CommandY", "-1")))
            GalaxyControlX = CInt(Val(oINI.GetString("INTERFACE", "GalaxyControlX", "-1")))
            GalaxyControlY = CInt(Val(oINI.GetString("INTERFACE", "GalaxyControlY", "-1")))
            MiningWindowX = CInt(Val(oINI.GetString("INTERFACE", "MiningWindowX", "-1")))
            MiningWindowY = CInt(Val(oINI.GetString("INTERFACE", "MiningWindowY", "-1")))
            WarCalendarX = CInt(Val(oINI.GetString("INTERFACE", "WarCalendarX", "-1")))
            WarCalendarY = CInt(Val(oINI.GetString("INTERFACE", "WarCalendarY", "-1")))
            UnitGotoLocX = CInt(Val(oINI.GetString("INTERFACE", "UnitGotoLocX", "-1")))
            UnitGotoLocY = CInt(Val(oINI.GetString("INTERFACE", "UnitGotoLocy", "-1")))
            BuilderWorksheetX = CInt(Val(oINI.GetString("INTERFACE", "BuilderWorksheetX", "-1")))
            BuilderWorksheetY = CInt(Val(oINI.GetString("INTERFACE", "BuilderWorksheetY", "-1")))
            ObserveX = CInt(Val(oINI.GetString("INTERFACE", "ObserveX", "-1")))
            ObserveY = CInt(Val(oINI.GetString("INTERFACE", "ObserveY", "-1")))
            EnvirRelationsX = CInt(Val(oINI.GetString("INTERFACE", "EnvirRelationsX", "-1")))
            EnvirRelationsY = CInt(Val(oINI.GetString("INTERFACE", "EnvirRelationsY", "-1")))
            TransportManagementX = CInt(Val(oINI.GetString("INTERFACE", "TransportManagementX", "-1")))
            TransportManagementY = CInt(Val(oINI.GetString("INTERFACE", "TransportManagementY", "-1")))
            TransportOrdersX = CInt(Val(oINI.GetString("INTERFACE", "TransportOrdersX", "-1")))
            TransportOrdersY = CInt(Val(oINI.GetString("INTERFACE", "TransportOrdersY", "-1")))

            'End of HUD Position settings

            'Begin HUD Options settings
            ResearchWindowStaticExpand = CInt(Val(oINI.GetString("INTERFACE", "ResearchWindowStaticExpand", "0"))) <> 0
            BudgetWindowStaticExpand = CInt(Val(oINI.GetString("INTERFACE", "BudgetWindowStaticExpand", "1"))) <> 0
            ShowConnectionStatus = CInt(Val(oINI.GetString("INTERFACE", "ShowConnectionStatus", "0"))) <> 0
            ShowMultiHealthBars = CInt(Val(oINI.GetString("INTERFACE", "ShowMultiHealthBars", "1"))) <> 0

            'End HUD Options settings

            'Begin Notification settings 
            MsgPersonnelProductionComplete = CInt(Val(oINI.GetString("NOTIFICATIONS", "PersonnelProductionComplete", "1")))
            MsgLandProductionComplete = CInt(Val(oINI.GetString("NOTIFICATIONS", "LandProductionComplete", "1")))
            MsgNavalProductionComplete = CInt(Val(oINI.GetString("NOTIFICATIONS", "NavalProductionComplete", "1")))
            MsgAerialProductionComplete = CInt(Val(oINI.GetString("NOTIFICATIONS", "AerialProductionComplete", "1")))
            MsgCommandCenterProductionComplete = CInt(Val(oINI.GetString("NOTIFICATIONS", "CommandCenterProductionComplete", "1")))
            MsgSpaceStationProductionComplete = CInt(Val(oINI.GetString("NOTIFICATIONS", "SpaceStationProductionComplete", "1")))
            MsgRefineryProductionComplete = CInt(Val(oINI.GetString("NOTIFICATIONS", "RefineryProductionComplete", "1")))

            'End Notification settings

            oINI = Nothing
        End Sub

        Public Sub SaveSettings()
            Dim oINI As New InitFile()

            oINI.WriteString("SETTINGS", "DYK_Frequency", lDYK_Frequency.ToString)

            If Windowed = True Then
                oINI.WriteString("GRAPHICS", "Windowed", "1")
            Else : oINI.WriteString("GRAPHICS", "Windowed", "0")
            End If
            If bFormationMoveThenForm = True Then
                oINI.WriteString("SETTINGS", "MoveThenForm", "1")
            Else : oINI.WriteString("SETTINGS", "MoveThenForm", "0")
            End If
            oINI.WriteString("GRAPHICS", "VertexProcessing", VertexProcessing.ToString)
            If SavedTripleBuffer = True Then
                oINI.WriteString("GRAPHICS", "TripleBuffer", "1")
            Else : oINI.WriteString("GRAPHICS", "TripleBuffer", "0")
            End If
            If SavedVSync = True Then
                oINI.WriteString("GRAPHICS", "VSync", "1")
            Else : oINI.WriteString("GRAPHICS", "VSync", "0")
            End If
            If ZoomChangesView = True Then
                oINI.WriteString("SETTINGS", "ZoomChangesView", "1")
            Else : oINI.WriteString("SETTINGS", "ZoomChangesView", "0")
            End If
            oINI.WriteString("GRAPHICS", "FullScreenResX", FullScreenResX.ToString)
            oINI.WriteString("GRAPHICS", "FullScreenResY", FullScreenResY.ToString)
            oINI.WriteString("GRAPHICS", "FullScreenRefreshRate", FullScreenRefreshRate.ToString)

            oINI.WriteString("SETTINGS", "RewardLabelFadeRate", RewardLabelFadeRate.ToString)
            oINI.WriteString("SETTINGS", "RewardLabelRiseRate", RewardLabelRiseRate.ToString)

            If bShowViewMessages = True Then
                oINI.WriteString("SETTINGS", "ShowViewMessages", "1")
            Else : oINI.WriteString("SETTINGS", "ShowViewMessages", "0")
            End If

            If RenderBurnMarks = True Then
                oINI.WriteString("GRAPHICS", "RenderBurnMarks", "1")
            Else : oINI.WriteString("GRAPHICS", "RenderBurnMarks", "0")
            End If

            If RenderPulseBolts = True Then
                oINI.WriteString("GRAPHICS", "RenderPulseBolts", "1")
            Else : oINI.WriteString("GRAPHICS", "RenderPulseBolts", "0")
            End If

            If WormholeAura = True Then
                oINI.WriteString("GRAPHICS", "WormholeAura", "1")
            Else : oINI.WriteString("GRAPHICS", "WormholeAura", "0")
            End If

            oINI.WriteString("SETTINGS", "FlashRate", FlashRate.ToString)
            If FlashSelections = True Then
                oINI.WriteString("SETTINGS", "FlashSelections", "1")
            Else : oINI.WriteString("SETTINGS", "FlashSelections", "0")
            End If

            oINI.WriteString("LIGHTING", "Quality", CInt(LightQuality).ToString)
            If RenderSpecularMap = True Then
                oINI.WriteString("GRAPHICS", "SpecularMap", "1")
            Else : oINI.WriteString("GRAPHICS", "SpecularMap", "0")
            End If
            oINI.WriteString("LIGHTING", "IlluminationMap", IlluminationMap.ToString)
            If RenderCosmos = True Then
                oINI.WriteString("GRAPHICS", "RenderCosmos", "1")
            Else : oINI.WriteString("GRAPHICS", "RenderCosmos", "0")
            End If
            If RenderPlanetRings = True Then
                oINI.WriteString("GRAPHICS", "RenderPlanetRings", "1")
            Else : oINI.WriteString("GRAPHICS", "RenderPlanetRings", "0")
            End If
            oINI.WriteString("GRAPHICS", "StarfieldParticlesPlanet", StarfieldParticlesPlanet.ToString)
            oINI.WriteString("GRAPHICS", "StarfieldParticlesSpace", StarfieldParticlesSpace.ToString)
            If BumpMapTerrain = True Then
                oINI.WriteString("LIGHTING", "BumpMapTerrain", "1")
            Else : oINI.WriteString("LIGHTING", "BumpMapTerrain", "0")
            End If
            'BumpMapPlanetModel = CInt(Val(oINI.GetString("LIGHTING", "BumpMapPlanetModel", "0"))) <> 0
            If BumpMapPlanetModel = True Then
                oINI.WriteString("LIGHTING", "BumpMapPlanetModel", "1")
            Else : oINI.WriteString("LIGHTING", "BumpMapPlanetModel", "0")
            End If
            If IlluminationMapTerrain = True Then
                oINI.WriteString("LIGHTING", "IllumMapTerrain", "1")
            Else : oINI.WriteString("LIGHTING", "IllumMapTerrain", "0")
            End If

            Dim lR As Int32
            Dim lG As Int32
            Dim lB As Int32
            lR = Math.Max(Math.Min(255, CInt(NeutralAssetColor.X * 255)), 0)
            lG = Math.Max(Math.Min(255, CInt(NeutralAssetColor.Y * 255)), 0)
            lB = Math.Max(Math.Min(255, CInt(NeutralAssetColor.Z * 255)), 0)
            oINI.WriteString("IDENTIFICATION", "NeutralR", lR.ToString)
            oINI.WriteString("IDENTIFICATION", "NeutralG", lG.ToString)
            oINI.WriteString("IDENTIFICATION", "NeutralB", lB.ToString)

            lR = Math.Max(Math.Min(255, CInt(MyAssetColor.X * 255)), 0)
            lG = Math.Max(Math.Min(255, CInt(MyAssetColor.Y * 255)), 0)
            lB = Math.Max(Math.Min(255, CInt(MyAssetColor.Z * 255)), 0)
            oINI.WriteString("IDENTIFICATION", "PlayerR", lR.ToString)
            oINI.WriteString("IDENTIFICATION", "PlayerG", lG.ToString)
            oINI.WriteString("IDENTIFICATION", "PlayerB", lB.ToString)

            lR = Math.Max(Math.Min(255, CInt(EnemyAssetColor.X * 255)), 0)
            lG = Math.Max(Math.Min(255, CInt(EnemyAssetColor.Y * 255)), 0)
            lB = Math.Max(Math.Min(255, CInt(EnemyAssetColor.Z * 255)), 0)
            oINI.WriteString("IDENTIFICATION", "EnemyR", lR.ToString)
            oINI.WriteString("IDENTIFICATION", "EnemyG", lG.ToString)
            oINI.WriteString("IDENTIFICATION", "EnemyB", lB.ToString)

            lR = Math.Max(Math.Min(255, CInt(AllyAssetColor.X * 255)), 0)
            lG = Math.Max(Math.Min(255, CInt(AllyAssetColor.Y * 255)), 0)
            lB = Math.Max(Math.Min(255, CInt(AllyAssetColor.Z * 255)), 0)
            oINI.WriteString("IDENTIFICATION", "AllyR", lR.ToString)
            oINI.WriteString("IDENTIFICATION", "AllyG", lG.ToString)
            oINI.WriteString("IDENTIFICATION", "AllyB", lB.ToString)

            lR = Math.Max(Math.Min(255, CInt(GuildAssetColor.X * 255)), 0)
            lG = Math.Max(Math.Min(255, CInt(GuildAssetColor.Y * 255)), 0)
            lB = Math.Max(Math.Min(255, CInt(GuildAssetColor.Z * 255)), 0)
            oINI.WriteString("IDENTIFICATION", "GuildR", lR.ToString)
            oINI.WriteString("IDENTIFICATION", "GuildG", lG.ToString)
            oINI.WriteString("IDENTIFICATION", "GuildB", lB.ToString)

            lR = Math.Max(Math.Min(255, CInt(TacticalAssetColor.X * 255)), 0)
            lG = Math.Max(Math.Min(255, CInt(TacticalAssetColor.Y * 255)), 0)
            lB = Math.Max(Math.Min(255, CInt(TacticalAssetColor.Z * 255)), 0)
            oINI.WriteString("IDENTIFICATION", "TacticalR", lR.ToString)
            oINI.WriteString("IDENTIFICATION", "TacticalG", lG.ToString)
            oINI.WriteString("IDENTIFICATION", "TacticalB", lB.ToString)

            With DefaultChatColor
                oINI.WriteString("CHATCOLOR", "DefaultR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "DefaultG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "DefaultB", .B.ToString)
            End With
            With AlertChatColor
                oINI.WriteString("CHATCOLOR", "AlertR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "AlertG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "AlertB", .B.ToString)
            End With
            With StatusChatColor
                oINI.WriteString("CHATCOLOR", "StatusR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "StatusG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "StatusB", .B.ToString)
            End With
            With LocalChatColor
                oINI.WriteString("CHATCOLOR", "LocalR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "LocalG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "LocalB", .B.ToString)
            End With
            With GuildChatColor
                oINI.WriteString("CHATCOLOR", "GuildR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "GuildG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "GuildB", .B.ToString)
            End With
            With PMChatColor
                oINI.WriteString("CHATCOLOR", "PMR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "PMG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "PMB", .B.ToString)
            End With
            With SenateChatColor
                oINI.WriteString("CHATCOLOR", "SenateR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "SenateG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "SenateB", .B.ToString)
            End With
            With ChannelChatColor
                oINI.WriteString("CHATCOLOR", "ChannelR", .R.ToString)
                oINI.WriteString("CHATCOLOR", "ChannelG", .G.ToString)
                oINI.WriteString("CHATCOLOR", "ChannelB", .B.ToString)
            End With

            With AcidBuildGhost
                oINI.WriteString("BUILDGHOST", "AcidR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "AcidG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "AcidB", .B.ToString)
            End With
            With AcidMineralCache
                oINI.WriteString("MINERALCACHE", "AcidR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "AcidG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "AcidB", .B.ToString)
            End With
            With AcidMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "AcidR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "AcidG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "AcidB", .B.ToString)
            End With


            With AdaptableBuildGhost
                oINI.WriteString("BUILDGHOST", "AdaptableR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "AdaptableG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "AdaptableB", .B.ToString)
            End With
            With AdaptableMineralCache
                oINI.WriteString("MINERALCACHE", "AdaptableR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "AdaptableG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "AdaptableB", .B.ToString)
            End With
            With AdaptableMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "AdaptableR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "AdaptableG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "AdaptableB", .B.ToString)
            End With

            With BarrenBuildGhost
                oINI.WriteString("BUILDGHOST", "BarrenR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "BarrenG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "BarrenB", .B.ToString)
            End With
            With BarrenMineralCache
                oINI.WriteString("MINERALCACHE", "BarrenR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "BarrenG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "BarrenB", .B.ToString)
            End With
            With BarrenMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "BarrenR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "BarrenG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "BarrenB", .B.ToString)
            End With

            With DesertBuildGhost
                oINI.WriteString("BUILDGHOST", "DesertR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "DesertG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "DesertB", .B.ToString)
            End With
            With DesertMineralCache
                oINI.WriteString("MINERALCACHE", "DesertR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "DesertG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "DesertB", .B.ToString)
            End With
            With DesertMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "DesertR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "DesertG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "DesertB", .B.ToString)
            End With

            With IceBuildGhost
                oINI.WriteString("BUILDGHOST", "IceR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "IceG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "IceB", .B.ToString)
            End With
            With IceMineralCache
                oINI.WriteString("MINERALCACHE", "IceR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "IceG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "IceB", .B.ToString)
            End With
            With IceMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "IceR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "IceG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "IceB", .B.ToString)
            End With

            With LavaBuildGhost
                oINI.WriteString("BUILDGHOST", "LavaR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "LavaG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "LavaB", .B.ToString)
            End With
            With LavaMineralCache
                oINI.WriteString("MINERALCACHE", "LavaR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "LavaG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "LavaB", .B.ToString)
            End With
            With LavaMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "LavaR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "LavaG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "LavaB", .B.ToString)
            End With

            With TerranBuildGhost
                oINI.WriteString("BUILDGHOST", "TerranR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "TerranG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "TerranB", .B.ToString)
            End With
            With TerranMineralCache
                oINI.WriteString("MINERALCACHE", "TerranR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "TerranG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "TerranB", .B.ToString)
            End With
            With TerranMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "TerranR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "TerranG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "TerranB", .B.ToString)
            End With

            With WaterworldBuildGhost
                oINI.WriteString("BUILDGHOST", "WaterworldR", .R.ToString)
                oINI.WriteString("BUILDGHOST", "WaterworldG", .G.ToString)
                oINI.WriteString("BUILDGHOST", "WaterworldB", .B.ToString)
            End With
            With WaterworldMineralCache
                oINI.WriteString("MINERALCACHE", "WaterworldR", .R.ToString)
                oINI.WriteString("MINERALCACHE", "WaterworldG", .G.ToString)
                oINI.WriteString("MINERALCACHE", "WaterworldB", .B.ToString)
            End With
            With WaterworldMinimapAngle
                oINI.WriteString("MINIMAPANGLE", "WaterworldR", .R.ToString)
                oINI.WriteString("MINIMAPANGLE", "WaterworldG", .G.ToString)
                oINI.WriteString("MINIMAPANGLE", "WaterworldB", .B.ToString)
            End With

            If ShowIntro = True Then
                oINI.WriteString("SETTINGS", "ShowIntroScreen", "1")
            Else : oINI.WriteString("SETTINGS", "ShowIntroScreen", "0")
            End If
            If ShowTargetBoxes = True Then
                oINI.WriteString("SETTINGS", "ShowTargetBoxes", "1")
            Else : oINI.WriteString("SETTINGS", "ShowTargetBoxes", "0")
            End If

            oINI.WriteString("GRAPHICS", "WaterRenderMethod", WaterRenderMethod.ToString)
            oINI.WriteString("GRAPHICS", "WaterTextureRes", WaterTextureRes.ToString)
            oINI.WriteString("GRAPHICS", "SmoothFOW", CInt(SmoothFOW).ToString)
            oINI.WriteString("GRAPHICS", "MaxLights", MaxLights.ToString)
            oINI.WriteString("GRAPHICS", "FarClippingPlane", FarClippingPlane.ToString)
            oINI.WriteString("GRAPHICS", "NearClippingPlane", NearClippingPlane.ToString)
            oINI.WriteString("GRAPHICS", "TerrainNearClippingPlane", TerrainNearClippingPlane.ToString)
            oINI.WriteString("GRAPHICS", "TerrainFarClippingPlane", TerrainFarClippingPlane.ToString)
            oINI.WriteString("PLANET_VIEW", "ShowMiniMap", CInt(ShowMiniMap).ToString)
            oINI.WriteString("PLANET_VIEW", "BlinkUnits", CInt(gbPlanetMapBlinkUnits).ToString)
            oINI.WriteString("GALAXY_VIEW", "LabelFalloff", CInt(gbGalaxyControlStarLabelFalloff).ToString)
            oINI.WriteString("GRAPHICS", "EntityClipPlane", EntityClipPlane.ToString)
            oINI.WriteString("GRAPHICS", "ModelTextureResolution", CInt(ModelTextureResolution).ToString)
            oINI.WriteString("GRAPHICS", "DrawGrid", CInt(DrawGrid).ToString)
            'oINI.WriteString("GRAPHICS", "DrawProceduralWater", CInt(DrawProceduralWater).ToString)
            oINI.WriteString("GRAPHICS", "ShowFOWTerrainShading", CInt(ShowFOWTerrainShading).ToString)
            oINI.WriteString("GRAPHICS", "FOWTextureResolution", FOWTextureResolution.ToString)
            oINI.WriteString("GRAPHICS", "ScreenshotFormat", ScreenshotFormat.ToString)
            oINI.WriteString("HUD", "MiniMapLocX", MiniMapLocX.ToString)
            oINI.WriteString("HUD", "MiniMapLocY", MiniMapLocY.ToString)
            oINI.WriteString("HUD", "MiniMapWidthHeight", MiniMapWidthHeight.ToString)
            oINI.WriteString("GRAPHICS", "RenderMineralCaches", CInt(RenderMineralCaches).ToString)
            oINI.WriteString("GRAPHICS", "DITHER", CInt(Dither).ToString)
            oINI.WriteString("GRAPHICS", "SpecularEnabled", CInt(SpecularEnabled).ToString)
            oINI.WriteString("SETTINGS", "NotificationDisplayTime", NotificationDisplayTime.ToString)
            oINI.WriteString("AUDIO", "AudioEnabled", CInt(AudioOn).ToString)
            oINI.WriteString("AUDIO", "PositionalAudio", CInt(PositionalSound).ToString)

            'oINI.WriteString("GRAPHICS", "WaterQuadRes", WaterQuadWidthHeight.ToString)
            oINI.WriteString("SETTINGS", "ScreenShakeEnable", CInt(ScreenShakeEnabled).ToString)
            oINI.WriteString("SETTINGS", "PlanetMinimapZoomLevel", CInt(PlanetMinimapZoomLevel).ToString)
            oINI.WriteString("GRAPHICS", "TerrainTextureResolution", TerrainTextureResolution.ToString)

            oINI.WriteString("GRAPHICS", "EngineFXParticles", EngineFXParticles.ToString)
            oINI.WriteString("GRAPHICS", "BurnFXParticles", BurnFXParticles.ToString)
            oINI.WriteString("GRAPHICS", "RenderShieldFX", CInt(RenderShieldFX).ToString)
            oINI.WriteString("GRAPHICS", "StarfieldParticles", StarfieldParticlesSpace.ToString)
            oINI.WriteString("GRAPHICS", "PlanetFXParticles", PlanetFXParticles.ToString)
            oINI.WriteString("GRAPHICS", "PostGlowAmount", PostGlowAmt.ToString("0.#"))
            oINI.WriteString("GRAPHICS", "HiResPlanetTexture", CInt(HiResPlanetTexture).ToString)
            oINI.WriteString("GRAPHICS", "PlanetModelTextureResolution", PlanetModelTextureWH.ToString)
            oINI.WriteString("GRAPHICS", "RenderExplosionFX", CInt(RenderExplosionFX).ToString)

            oINI.WriteString("HUD", "BuildWindowLeft", BuildWindowLeft.ToString)
            oINI.WriteString("HUD", "BuildWindowTop", BuildWindowTop.ToString)
            oINI.WriteString("HUD", "ResearchWindowLeft", ResearchWindowLeft.ToString)
            oINI.WriteString("HUD", "ResearchWindowTop", ResearchWindowTop.ToString)
            oINI.WriteString("HUD", "ColonialManagementLeft", ColonialManagementLeft.ToString)
            oINI.WriteString("HUD", "ColonialManagementTop", ColonialManagementTop.ToString)
            oINI.WriteString("HUD", "ColonialManagementHeight", ColonialManagementHeight.ToString)
            oINI.WriteString("HUD", "SpecialTechLeft", SpecialTechLeft.ToString)
            oINI.WriteString("HUD", "SpecialTechTop", SpecialTechTop.ToString)

            oINI.WriteString("INTERFACE", "ResearchWindowStaticExpand", CInt(ResearchWindowStaticExpand).ToString)
            oINI.WriteString("INTERFACE", "BudgetWindowStaticExpand", CInt(BudgetWindowStaticExpand).ToString)
            oINI.WriteString("INTERFACE", "ShowConnectionStatus", CInt(ShowConnectionStatus).ToString)
            oINI.WriteString("INTERFACE", "ShowMultiHealthBars", CInt(ShowMultiHealthBars).ToString)

            oINI.WriteString("SETTINGS", "CtrlQExits", CInt(CtrlQExits).ToString)
            oINI.WriteString("SETTINGS", "RenderHPBars", CInt(RenderHPBars).ToString)
            oINI.WriteString("SETTINGS", "DisableExclamations", CInt(gbDisableExclamations).ToString)
            oINI.WriteString("SETTINGS", "ExportedDataFormat", ExportedDataFormat.ToString)
            oINI.WriteString("SETTINGS", "CachePlanetTextures", CachePlanetTextures.ToString)
            '(Disabled for now) oINI.WriteString("INTERFACE", "DoNotShowEngineerCancelAlert", CInt(bDoNotShowEngineerCancelAlert).ToString)

            If bRanBefore = True Then oINI.WriteString("SETTINGS", "RanBefore", "1")

            oINI.WriteString("INTERFACE", "BorderColor_A", InterfaceBorderColor.A.ToString)
            oINI.WriteString("INTERFACE", "BorderColor_R", InterfaceBorderColor.R.ToString)
            oINI.WriteString("INTERFACE", "BorderColor_G", InterfaceBorderColor.G.ToString)
            oINI.WriteString("INTERFACE", "BorderColor_B", InterfaceBorderColor.B.ToString)

            oINI.WriteString("INTERFACE", "FillColor_A", InterfaceFillColor.A.ToString)
            oINI.WriteString("INTERFACE", "FillColor_R", InterfaceFillColor.R.ToString)
            oINI.WriteString("INTERFACE", "FillColor_G", InterfaceFillColor.G.ToString)
            oINI.WriteString("INTERFACE", "FillColor_B", InterfaceFillColor.B.ToString)

            oINI.WriteString("INTERFACE", "TextBoxFillColor_A", InterfaceTextBoxFillColor.A.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_R", InterfaceTextBoxFillColor.R.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_G", InterfaceTextBoxFillColor.G.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_B", InterfaceTextBoxFillColor.B.ToString)

            oINI.WriteString("INTERFACE", "TextBoxForeColor_A", InterfaceTextBoxForeColor.A.ToString)
            oINI.WriteString("INTERFACE", "TextBoxForeColor_R", InterfaceTextBoxForeColor.R.ToString)
            oINI.WriteString("INTERFACE", "TextBoxForeColor_G", InterfaceTextBoxForeColor.G.ToString)
            oINI.WriteString("INTERFACE", "TextBoxForeColor_B", InterfaceTextBoxForeColor.B.ToString)

            oINI.WriteString("INTERFACE", "ButtonColor_A", InterfaceButtonColor.A.ToString)
            oINI.WriteString("INTERFACE", "ButtonColor_R", InterfaceButtonColor.R.ToString)
            oINI.WriteString("INTERFACE", "ButtonColor_G", InterfaceButtonColor.G.ToString)
            oINI.WriteString("INTERFACE", "ButtonColor_B", InterfaceButtonColor.B.ToString)

            oINI.WriteString("GRAPHICS", "AmbientLevel", AmbientLevel.ToString)

            'Hud Position Settings
            If AdvanceDisplayLocX <> -1 Then oINI.WriteString("INTERFACE", "AdvanceDisplayLocX", AdvanceDisplayLocX.ToString)
            If AdvanceDisplayLocY <> -1 Then oINI.WriteString("INTERFACE", "AdvanceDisplayLocY", AdvanceDisplayLocY.ToString)
            If BehaviorLocX <> -1 Then oINI.WriteString("INTERFACE", "BehaviorLocX", BehaviorLocX.ToString)
            If BehaviorLocY <> -1 Then oINI.WriteString("INTERFACE", "BehaviorLocY", BehaviorLocY.ToString)
            If ContentsLocX <> -1 Then oINI.WriteString("INTERFACE", "ContentsLocX", ContentsLocX.ToString)
            If ContentsLocY <> -1 Then oINI.WriteString("INTERFACE", "ContentsLocY", ContentsLocY.ToString)
            If EnvirDisplayLocX <> -1 Then oINI.WriteString("INTERFACE", "EnvirDisplayLocX", EnvirDisplayLocX.ToString)
            If EnvirDisplayLocY <> -1 Then oINI.WriteString("INTERFACE", "EnvirDisplayLocY", EnvirDisplayLocY.ToString)
            If SelectLocX <> -1 Then oINI.WriteString("INTERFACE", "SelectLocX", SelectLocX.ToString)
            If SelectLocY <> -1 Then oINI.WriteString("INTERFACE", "SelectLocY", SelectLocY.ToString)
            If MultiSelectLocX <> -1 Then oINI.WriteString("INTERFACE", "MultiSelectLocX", MultiSelectLocX.ToString)
            If MultiSelectLocY <> -1 Then oINI.WriteString("INTERFACE", "MultiSelectLocY", MultiSelectLocY.ToString)
            If MultiSelectWidth <> -1 Then oINI.WriteString("INTERFACE", "MultiSelectWidth", MultiSelectWidth.ToString)
            If MultiSelectHeight <> -1 Then oINI.WriteString("INTERFACE", "MultiSelectHeight", MultiSelectHeight.ToString)
            If ProdStatusLocX <> -1 Then oINI.WriteString("INTERFACE", "ProdStatusLocX", ProdStatusLocX.ToString)
            If ProdStatusLocY <> -1 Then oINI.WriteString("INTERFACE", "ProdStatusLocY", ProdStatusLocY.ToString)
            If ProdStatusWidth <> -1 Then oINI.WriteString("INTERFACE", "ProdStatusWidth", ProdStatusWidth.ToString)
            If ChatWindowLocX <> -1 Then oINI.WriteString("INTERFACE", "ChatWindowLocX", ChatWindowLocX.ToString)
            If ChatWindowLocY <> -1 Then oINI.WriteString("INTERFACE", "ChatWindowLocY", ChatWindowLocY.ToString)
            If ChatWindowWidth <> -1 Then oINI.WriteString("INTERFACE", "ChatWindowWidth", ChatWindowWidth.ToString)
            If ChatWindowHeight <> -1 Then oINI.WriteString("INTERFACE", "ChatWindowHeight", ChatWindowHeight.ToString)
            oINI.WriteString("INTERFACE", "ChatWindowState", ChatWindowState.ToString)

            oINI.WriteString("INTERFACE", "TutorialTOCX", TutorialTOCX.ToString)
            oINI.WriteString("INTERFACE", "TutorialTOCY", TutorialTOCY.ToString)
            oINI.WriteString("INTERFACE", "EmailMainX", EmailMainX.ToString)
            oINI.WriteString("INTERFACE", "EmailMainY", EmailMainY.ToString)
            oINI.WriteString("INTERFACE", "FleetX", FleetX.ToString)
            oINI.WriteString("INTERFACE", "FleetY", FleetY.ToString)
            oINI.WriteString("INTERFACE", "TradeMainX", TradeMainX.ToString)
            oINI.WriteString("INTERFACE", "TradeMainY", TradeMainY.ToString)
            oINI.WriteString("INTERFACE", "DiplomacyX", DiplomacyX.ToString)
            oINI.WriteString("INTERFACE", "DiplomacyY", DiplomacyY.ToString)
            oINI.WriteString("INTERFACE", "ColonyStatsX", ColonyStatsX.ToString)
            oINI.WriteString("INTERFACE", "ColonyStatsY", ColonyStatsY.ToString)
            oINI.WriteString("INTERFACE", "BudgetX", BudgetX.ToString)
            oINI.WriteString("INTERFACE", "BudgetY", BudgetY.ToString)
            oINI.WriteString("INTERFACE", "AgentMainX", AgentMainX.ToString)
            oINI.WriteString("INTERFACE", "AgentMainY", AgentMainY.ToString)
            oINI.WriteString("INTERFACE", "ArenaMainX", ArenaMainX.ToString)
            oINI.WriteString("INTERFACE", "ArenaMainY", ArenaMainY.ToString)
            oINI.WriteString("INTERFACE", "ArenaConfigX", ArenaConfigX.ToString)
            oINI.WriteString("INTERFACE", "ArenaConfigY", ArenaConfigY.ToString)
            oINI.WriteString("INTERFACE", "ArenaWaitX", ArenaWaitX.ToString)
            oINI.WriteString("INTERFACE", "ArenaWaitY", ArenaWaitY.ToString)
            oINI.WriteString("INTERFACE", "AgentMissionCreateX", AgentMissionCreateX.ToString)
            oINI.WriteString("INTERFACE", "AgentMissionCreateY", AgentMissionCreateY.ToString)
            oINI.WriteString("INTERFACE", "AgentMissionDetailsX", AgentMissionDetailsX.ToString)
            oINI.WriteString("INTERFACE", "AgentMissionDetailsY", AgentMissionDetailsY.ToString)
            oINI.WriteString("INTERFACE", "AgentDetailsX", AgentDetailsX.ToString)
            oINI.WriteString("INTERFACE", "AgentDetailsY", AgentDetailsY.ToString)
            oINI.WriteString("INTERFACE", "FormationsX", FormationsX.ToString)
            oINI.WriteString("INTERFACE", "FormationsY", FormationsY.ToString)
            oINI.WriteString("INTERFACE", "GuildMainX", GuildMainX.ToString)
            oINI.WriteString("INTERFACE", "GuildMainY", GuildMainY.ToString)
            oINI.WriteString("INTERFACE", "AvailResX", AvailResX.ToString)
            oINI.WriteString("INTERFACE", "AvailResY", AvailResY.ToString)
            oINI.WriteString("INTERFACE", "CommandX", CommandX.ToString)
            oINI.WriteString("INTERFACE", "CommandY", CommandY.ToString)
            oINI.WriteString("INTERFACE", "GalaxyControlX", GalaxyControlX.ToString)
            oINI.WriteString("INTERFACE", "GalaxyControlY", GalaxyControlY.ToString)
            oINI.WriteString("INTERFACE", "MiningWindowX", MiningWindowX.ToString)
            oINI.WriteString("INTERFACE", "MiningWindowY", MiningWindowY.ToString)
            oINI.WriteString("INTERFACE", "WarCalendarX", WarCalendarX.ToString)
            oINI.WriteString("INTERFACE", "WarCalendarY", WarCalendarY.ToString)
            oINI.WriteString("INTERFACE", "UnitGotoLocX", UnitGotoLocX.ToString)
            oINI.WriteString("INTERFACE", "UnitGotoLocY", UnitGotoLocY.ToString)
            oINI.WriteString("INTERFACE", "BuilderWorksheetX", BuilderWorksheetX.ToString)
            oINI.WriteString("INTERFACE", "BuilderWorksheetY", BuilderWorksheetY.ToString)
            oINI.WriteString("INTERFACE", "ObserveX", ObserveX.ToString)
            oINI.WriteString("INTERFACE", "ObserveY", ObserveY.ToString)
            oINI.WriteString("INTERFACE", "EnvirRelationsX", EnvirRelationsX.ToString)
            oINI.WriteString("INTERFACE", "EnvirRelationsY", EnvirRelationsY.ToString)
            oINI.WriteString("INTERFACE", "TransportManagementX", TransportManagementX.ToString)
            oINI.WriteString("INTERFACE", "TransportManagementY", TransportManagementY.ToString)
            oINI.WriteString("INTERFACE", "TransportOrdersX", TransportOrdersX.ToString)
            oINI.WriteString("INTERFACE", "TransportOrdersY", TransportOrdersY.ToString)

            oINI.WriteString("NOTIFICATIONS", "PersonnelProductionComplete", MsgPersonnelProductionComplete.ToString)
            oINI.WriteString("NOTIFICATIONS", "LandProductionComplete", MsgLandProductionComplete.ToString)
            oINI.WriteString("NOTIFICATIONS", "NavalProductionComplete", MsgNavalProductionComplete.ToString)
            oINI.WriteString("NOTIFICATIONS", "AerialProductionComplete", MsgAerialProductionComplete.ToString)
            oINI.WriteString("NOTIFICATIONS", "CommandCenterProductionComplete", MsgCommandCenterProductionComplete.ToString)
            oINI.WriteString("NOTIFICATIONS", "SpaceStationProductionComplete", MsgSpaceStationProductionComplete.ToString)
            oINI.WriteString("NOTIFICATIONS", "RefineryProductionComplete", MsgRefineryProductionComplete.ToString)


            'End of HUD Position settings
            If FilterBadWords = True Then oINI.WriteString("INTERFACE", "FilterBadWords", "1") Else oINI.WriteString("INTERFACE", "FilterBadWords", "0")
            If ChatTimeStamps = True Then oINI.WriteString("INTERFACE", "ChatTimeStamps", "1") Else oINI.WriteString("INTERFACE", "ChatTimeStamps", "0")

            oINI = Nothing
        End Sub
    End Structure
    Public muSettings As EngineSettings
#End Region

#Region " Client-Specific Data (Player, Environment) And Engine Code "
    Public goControlGroups As ControlGroups

    Public glPlayerID As Int32
    Public goCurrentPlayer As Player
    Public gsUserName As String
    Public gsPassword As String
    Public gsUserNameProper As String
    Public glActualPlayerID As Int32        'for aliased situations
    Public gsAliasUserName As String
    Public gsAliasPassword As String
    Public glAliasRights As Int32           'not very secure
    Public gbAliased As Boolean = False
    Public gyTriggerFired() As Byte

    Public Function HasAliasedRights(ByVal lRights As AliasingRights) As Boolean
        Return (gbAliased = False) OrElse (glAliasRights And lRights) <> 0 ' = lRights
    End Function

    Public Enum eyClaimFlags As Byte
        eClaimed = 1
    End Enum
    Public Structure uClaimable
        Public sName As String
        Public lID As Int32
        Public iTypeID As Int16
        Public lOfferCode As Int32
        Public yClaimFlag As Byte
    End Structure
    Public guClaimables() As uClaimable

    'Manages player intelligence... specifically, what the player knows about the players in this list
    Public goPlayerIntel() As PlayerIntel
    Public glPlayerIntelIdx() As Int32
    Public glPlayerIntelUB As Int32 = -1

    Public goPlayerTechKnowledge() As PlayerTechKnowledge
    Public glPlayerTechKnowledgeIdx() As Int32
    Public glPlayerTechKnowledgeUB As Int32 = -1

    Public goItemIntel() As PlayerItemIntel
    Public glItemIntelIdx() As Int32
    Public glItemIntelUB As Int32 = -1
    'End of Player Intel

    Public goCurrentEnvir As BaseEnvironment
    Public glCurrentCycle As Int32
    Public gfCurrentCyclePrecise As Single

    Public goEntityDefs() As EntityDef
    Public glEntityDefIdx() As Int32
    Public glEntityDefUB As Int32 = -1

    Public goMineralProperty() As MineralProperty
    Public glMineralPropertyIdx() As Int32
    Public glMineralPropertyUB As Int32 = -1

    Public goMinerals() As Mineral
    Public glMineralIdx() As Int32
    Public glMineralUB As Int32 = -1

    Public goSkills() As Skill
    Public glSkillIdx() As Int32
    Public glSkillUB As Int32 = -1

    Public goGoals() As Goal
    Public glGoalIdx() As Int32
    Public glGoalUB As Int32 = -1

    Public goMissions() As Mission
    Public glMissionIdx() As Int32
    Public glMissionUB As Int32 = -1

    Public gsMissionMethods() As String
    Public gsMissionDescs() As String
    Public glMissionMethodIdx() As Int32
    Public glMissionMethodUB As Int32 = -1

    Public gb_InHandleMovement As Boolean = False

    Public Sub HandleMovement()
        Dim X As Int32
        Dim bTurned As Boolean
        Dim iTemp As Int16
        Dim iTurnAmt As Int16
        Dim bTurnedCCW As Boolean
        Dim lCyclesToStop As Int32
        Dim fDistToStop As Single
        Dim fTotDist As Single
        Dim bAdjustYaw As Boolean

        'For dealing with Skipping...
        Dim lCyclesLapsed As Int32

        Dim fTempVelX As Single
        Dim fTempVelZ As Single
        Dim dVecAngleRads As Single

        Dim lFar As Int32 = muSettings.EntityClipPlane

        Dim fMaxSpeed As Single

        Dim oCurrEntity As BaseEntity
        Dim oCurrEnvir As BaseEnvironment = goCurrentEnvir

        gb_InHandleMovement = True

        If oCurrEnvir Is Nothing Then Return

        Dim lCurrTime As Int32 = RenderObject.GetCurrTime()

        Dim lCurUB As Int32 = -1

        If oCurrEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oCurrEnvir.lEntityUB, oCurrEnvir.lEntityIdx.GetUpperBound(0))
        For X = 0 To lCurUB
            'validate the unit object
            If oCurrEnvir.lEntityIdx(X) <> -1 Then

                oCurrEntity = oCurrEnvir.oEntity(X)
                If oCurrEntity Is Nothing Then
                    oCurrEnvir.lEntityIdx(X) = -1
                    Continue For
                End If


                With oCurrEntity
                    lCyclesLapsed = glCurrentCycle - .LastUpdateCycle

                    If (.oUnitDef.Acceleration = 0 OrElse .oUnitDef.MaxSpeed = 0) AndAlso (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                        .CurrentStatus -= elUnitStatus.eUnitMoving
                    End If

                    If .oUnitDef.yFormationManeuver <> 0 Then iTurnAmt = .oUnitDef.yFormationManeuver Else iTurnAmt = .oUnitDef.TurnAmount

                    If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso (lCyclesLapsed > 0) Then

                        Dim lCyclesAccelerate As Int32 = 0
                        Dim lCyclesDecelerate As Int32 = 0
                        Dim iTAx100 As Int16

                        bAdjustYaw = True
                        If .LocAngle <> .DestAngle Then
                            bTurned = True

                            'Change our Angle
                            iTemp = .DestAngle - .LocAngle

                            'Ok, we turn, while we turn, we decelerate... unless our turn amount x100 > Abs(LocAngle - DestAngle), 
                            'at which point, we accelerate

                            'we decelerate unless we are in the x100 area, at which point, we accelerate
                            'so, we cannot turn more than 3600, the difference of dest and loc, or my turn amount * cycles
                            Dim iMaxTurnPotential As Int16 = CShort(Math.Min(3600, Math.Min(iTurnAmt * lCyclesLapsed, Math.Abs(iTemp))))
                            'calculate turn amount x100
                            iTAx100 = iTurnAmt * 100S
                            Dim lTurnInDecelerate As Int32 = Math.Min(iMaxTurnPotential, Math.Abs(iTemp))
                            Dim lTurnInAccelerate As Int32 = 0 '= Math.Min(lTurnInDecelerate, Math.Max(0, iTAx100 - Math.Abs(iTemp) + iMaxTurnPotential))
                            If Math.Abs(iTemp) - lTurnInDecelerate < iTAx100 Then
                                lTurnInAccelerate = Math.Min(iMaxTurnPotential, Math.Abs(iTAx100 - Math.Abs(iTemp) - lTurnInDecelerate))
                            End If
                            lTurnInDecelerate -= lTurnInAccelerate
                            lCyclesDecelerate = lTurnInDecelerate \ iTurnAmt
                            lCyclesAccelerate = lCyclesLapsed - lCyclesDecelerate  'lTurnInAccelerate \ iTurnAmt

                            If Math.Abs(iTemp) < iMaxTurnPotential Then bAdjustYaw = False

                            'Ok, apply our turn amount
                            If Math.Abs(iTemp) = 1800S Then
                                .LocAngle -= iMaxTurnPotential
                            Else
                                If iTemp < 0 Then
                                    If iTemp > -1800S Then
                                        'CCW
                                        .LocAngle -= iMaxTurnPotential
                                        bTurnedCCW = True
                                    Else
                                        'CW
                                        .LocAngle += iMaxTurnPotential
                                        bTurnedCCW = False
                                    End If
                                Else
                                    If iTemp > 1800S Then
                                        'CCW
                                        .LocAngle -= iMaxTurnPotential
                                        bTurnedCCW = True
                                    Else
                                        'CW
                                        .LocAngle += iMaxTurnPotential
                                        bTurnedCCW = False
                                    End If
                                End If  'if iTemp < 0
                                If .LocAngle < 0 Then
                                    .LocAngle += 3600S
                                ElseIf .LocAngle > 3599S Then
                                    .LocAngle -= 3600S
                                End If
                            End If
                        Else
                            bTurned = False
                            lCyclesAccelerate = lCyclesLapsed
                            lCyclesDecelerate = 0
                        End If

                        If bTurned = True AndAlso bAdjustYaw = True Then
                            If bTurnedCCW = True Then
                                'lean left... negative yaw
                                .LocYaw -= iTurnAmt
                            Else
                                'lean right... positive yaw
                                .LocYaw += iTurnAmt
                            End If
                            If .LocYaw > gl_MAX_YAW Then
                                .LocYaw = gl_MAX_YAW
                            ElseIf .LocYaw < gl_MIN_YAW Then
                                .LocYaw = gl_MIN_YAW
                            End If
                        ElseIf .LocYaw <> 0 Then
                            If iTurnAmt > Math.Abs(.LocYaw) Then
                                'just set to 0
                                .LocYaw = 0
                            Else
                                If .LocYaw < 0 Then
                                    .LocYaw += iTurnAmt
                                Else
                                    .LocYaw -= iTurnAmt
                                End If
                            End If
                        End If

                        'Caculate our MaxSpeed
                        If .oUnitDef.yFormationMaxSpeed <> 0 Then fMaxSpeed = .oUnitDef.yFormationMaxSpeed Else fMaxSpeed = .oUnitDef.MaxSpeed
                        If .oMesh Is Nothing = False AndAlso .oMesh.bLandBased = True Then
                            'If .fMapWrapLocX = Single.MinValue Then .fMapWrapLocX = .LocX
                            fMaxSpeed *= oCurrEnvir.GetEnvirSpeedMod(CInt(.LocX), CInt(.LocZ), .LocAngle)
                            If fMaxSpeed < 1 Then fMaxSpeed = 1
                            'Else : fMaxSpeed = .oUnitDef.MaxSpeed
                        End If

                        Dim fAcc As Single
                        If .oUnitDef.yFormationManeuver <> 0 Then fAcc = .oUnitDef.fFormationAcceleration Else fAcc = .oUnitDef.Acceleration
                        'Ok... determine our distance to begin slowing down 
                        If fAcc = 0 Then lCyclesToStop = 0 Else lCyclesToStop = CInt(System.Math.Abs(.TotalVelocity / fAcc))
                        fDistToStop = System.Math.Abs(.TotalVelocity * lCyclesToStop) + (0.5F * fAcc * lCyclesToStop)

                        'Total distance from dest - approx
                        Dim fAbsDiffX As Single = System.Math.Abs(.DestX - .LocX)
                        Dim fAbsDiffZ As Single = System.Math.Abs(.DestZ - .LocZ)
                        fTotDist = fAbsDiffX + fAbsDiffZ

                        Dim fAdditionalTravel As Single = 0.0F
                        Dim bDoTurnCalc As Boolean = False

                        'Ok, we may need to burn some cycles decelerating (for example, if we are in a turn)
                        If lCyclesDecelerate <> 0 Then
                            'ok, decelerate because we are turning... no DI, maneuver change or stop event
                            If .TotalVelocity <> 0 Then
                                'ok, we are moving while decelerating, determine by how much we are moving
                                Dim lActualCyclesDecelerate As Int32 = Math.Min(lCyclesDecelerate, lCyclesToStop)
                                Dim fTmpAcc As Single = fAcc * 0.25F
                                If lActualCyclesDecelerate > 43000 Then lActualCyclesDecelerate = 43000
                                fAdditionalTravel += (.TotalVelocity * lActualCyclesDecelerate) + (-fTmpAcc * (lActualCyclesDecelerate * lActualCyclesDecelerate) * 0.5F) ' ((fTmpAcc * lActualCyclesDecelerate) * (fTmpAcc * lActualCyclesDecelerate)) * 0.5F
                                .TotalVelocity -= (fTmpAcc * lActualCyclesDecelerate)
                                If .TotalVelocity < 0 Then .TotalVelocity = 0
                                bDoTurnCalc = True
                            End If
                        End If

                        'Now, we are here, lCyclesAccelerate are the remaining movement cycles, while moving...
                        '  we will attempt to accelerate as long as we will have room remaining to decelerate
                        '  furthermore, we cannot exceed our maxspeed
                        If lCyclesAccelerate <> 0 Then
                            Dim lActualAcc As Int32 = 0
                            'I can accelerate to a point where my maxspeed and distance travelled will exceed the ability to stop
                            'm = (al - l) +/- sqrt ( 5l^2 - a^2l^2 + 4a^2d + 4ad )
                            '    -------------------------------------------------
                            '						2a+2
                            'a = acceleration
                            'l = vel start
                            'd = dist to target

                            Dim fM As Single
                            If .TotalVelocity <> fMaxSpeed Then
                                Dim fVelAtStart As Single = .TotalVelocity
                                Dim fASquared As Single = fAcc * fAcc
                                Dim fLSquared As Single = fVelAtStart * fVelAtStart
                                fM = 5.0F * fLSquared
                                fM -= fASquared * fLSquared
                                fM += 4.0F * fASquared * fTotDist
                                fM += 4.0F * fAcc * fTotDist
                                fM = CSng(Math.Sqrt(fM))
                                fM += (fAcc * fVelAtStart - fVelAtStart)
                                fM /= (2 * fAcc + 2)
                                fM = Math.Min(fM, fMaxSpeed)
                            Else : fM = fMaxSpeed
                            End If

                            'fM now has our max speed before we need to decelerate
                            If .TotalVelocity < fM Then lActualAcc = CInt(Math.Min(lCyclesAccelerate, (fM - .TotalVelocity) / fAcc))

                            If lActualAcc <> 0 Then
                                'ok, now reduce lActualAcc from lCyclesAccelerate
                                lCyclesAccelerate -= lActualAcc
                                'Next, we accelerate
                                If lActualAcc > 43000 Then lActualAcc = 43000
                                fAdditionalTravel += (.TotalVelocity * lActualAcc) + (fAcc * (lActualAcc * lActualAcc) * 0.5F)
                                .TotalVelocity += (lActualAcc * fAcc)
                            End If

                            'Now, do we have remaining cycles? if so, we need to calculate the distance travelled at present speed (max)
                            If lCyclesAccelerate <> 0 Then
                                'ok, did we actually accelerate?
                                .TotalVelocity = fM
                                If lActualAcc <> 0 Then
                                    'recalculate our dist to stop
                                    fDistToStop = System.Math.Abs(.TotalVelocity * lCyclesToStop) + (0.5F * fAcc * lCyclesToStop)
                                End If
                                'Ok, now, determine if we need to decelerate...
                                If fTotDist - fAdditionalTravel < fDistToStop Then
                                    'ok, we need to decelerate
                                    lCyclesDecelerate = lCyclesAccelerate
                                    lCyclesToStop = CInt(System.Math.Abs(.TotalVelocity / fAcc))
                                    Dim lActualCyclesDecelerate As Int32 = Math.Min(lCyclesDecelerate, lCyclesToStop)

                                    If lActualCyclesDecelerate > 43000 Then lActualCyclesDecelerate = 43000
                                    fAdditionalTravel += (.TotalVelocity * lActualCyclesDecelerate) + (fAcc * (lActualCyclesDecelerate * lActualCyclesDecelerate) * 0.5F)
                                    .TotalVelocity -= (fAcc * lActualCyclesDecelerate)
                                    If .TotalVelocity < 0 Then .TotalVelocity = 0
                                Else
                                    fAdditionalTravel += (.TotalVelocity * lCyclesAccelerate)
                                End If
                            End If
                        End If

                        If fAdditionalTravel > fTotDist Then fAdditionalTravel = fTotDist
                        If fAdditionalTravel < .TotalVelocity Then fAdditionalTravel = .TotalVelocity

                        If bDoTurnCalc = True Then
                            If fTotDist <> 0 Then
                                .VelX = .LocAngle Mod 8 ' (.LocAngle \ 10) Mod 8
                                .VelX *= 0.125F
                                .VelX *= fAdditionalTravel
                                .VelZ = fAdditionalTravel - .VelX
                            End If
                        Else
                            If fTotDist <> 0 Then
                                .VelX = fAbsDiffX
                                .VelX = .VelX / fTotDist
                                .VelX *= fAdditionalTravel
                                .VelZ = fAdditionalTravel - .VelX
                            End If
                        End If

                        'So, at this point we have a VelX and VelZ
                        If .VelX = 0 Then
                            dVecAngleRads = gdPieAndAHalf
                        ElseIf .VelZ = 0 Then
                            dVecAngleRads = 0 'gdTwoPie
                        Else
                            'angled
                            dVecAngleRads = CSng(Math.Atan(Math.Abs(.VelZ / .VelX)))
                            dVecAngleRads = gdTwoPie - dVecAngleRads
                        End If
                        Const fPiOver180 As Single = gdPi / 180.0F
                        dVecAngleRads = (.LocAngle * 0.1F) * fPiOver180 - dVecAngleRads
                        Dim fTmpCos As Single = CSng(Math.Cos(dVecAngleRads))
                        Dim fTmpSin As Single = CSng(Math.Sin(dVecAngleRads))
                        fTempVelX = (.VelX * fTmpCos) + (.VelZ * fTmpSin)
                        fTempVelZ = -((.VelX * fTmpSin) - (.VelZ * fTmpCos))
                        .VelX = fTempVelX
                        .VelZ = fTempVelZ

                        If .LocX <> .DestX Then
                            If fAbsDiffX > Math.Abs(.VelX) Then
                                .LocX += .VelX
                            Else
                                .LocX = .DestX
                            End If
                        Else
                            If .DestZ < .LocZ Then
                                .VelZ = -fAdditionalTravel
                            Else
                                .VelZ = fAdditionalTravel
                            End If
                            'If .VelZ < 0 Then .VelZ = -fAdditionalTravel Else .VelZ = fAdditionalTravel
                        End If
                        If .LocZ <> .DestZ Then
                            If fAbsDiffZ > Math.Abs(.VelZ) Then
                                .LocZ += .VelZ
                            Else
                                .LocZ = .DestZ
                            End If
                        Else
                            'If .VelX < 0 Then .VelX = -fAdditionalTravel Else .VelX = fAdditionalTravel
                            If .DestX < .LocX Then
                                .VelX = -fAdditionalTravel
                            Else
                                .VelX = fAdditionalTravel
                            End If
                        End If

                        'This is for map wrapping (east west)
                        If .LocX < oCurrEnvir.lMinXPos Then
                            '.LocX += oCurrEnvir.lMapWrapAdjustX
                            '.DestX += oCurrEnvir.lMapWrapAdjustX
                            .LocX = oCurrEnvir.lMinXPos
                        ElseIf .LocX > oCurrEnvir.lMaxXPos Then
                            '.LocX -= oCurrEnvir.lMapWrapAdjustX
                            '.DestX -= oCurrEnvir.lMapWrapAdjustX
                            .LocX = oCurrEnvir.lMaxXPos
                        End If
                        If .LocZ < oCurrEnvir.lMinZPos Then
                            '.LocZ = oCurrEnvir.lMinZPos
                            '.DestZ = oCurrEnvir.lMinZPos
                            .LocZ = oCurrEnvir.lMinZPos
                        ElseIf .LocZ > oCurrEnvir.lMaxZPos Then
                            '.LocZ = oCurrEnvir.lMaxZPos
                            '.DestZ = oCurrEnvir.lMaxZPos
                            .LocZ = oCurrEnvir.lMaxZPos
                        End If

                        'Now, check for total stop, if we are closer than the acceleration, then stop us
                        If System.Math.Abs(.DestX - .LocX) < fAcc Then
                            .LocX = .DestX
                        End If
                        If System.Math.Abs(.DestZ - .LocZ) < fAcc Then
                            .LocZ = .DestZ
                        End If

                        'Finally, get our destination angle
                        'Added If bTurned and bAdjustYaw to get rid of annoying 'wiggle'
                        If bTurned = True AndAlso bAdjustYaw = True Then
                            If .LocX <> .DestX Or .LocZ <> .DestZ Then
                                .DestAngle = CShort(LineAngleDegrees(CInt(.LocX), CInt(.LocZ), .DestX, .DestZ) * 10)
                            End If
                        End If

                        'check if we have reached our destination...
                        If .DestX = CInt(.LocX) AndAlso .LocZ = CInt(.DestZ) AndAlso .DestAngle = .LocAngle Then
                            If .DestAngle <> .TrueDestAngle AndAlso .TrueDestAngle <> -1 Then
                                .DestAngle = .TrueDestAngle
                            Else
                                'ok, we're not the server so simply change the unit's status to not moving
                                .CurrentStatus -= elUnitStatus.eUnitMoving
                                .ShutoffEngines()
                            End If
                        End If

                        'This differs from Server's Movement because we need the cell
                        .CellLocX = CInt(.LocX) \ lFar ' CInt(Math.Floor(.LocX / lFar))
                        .CellLocZ = CInt(.LocZ) \ lFar  'CInt(Math.Floor(.LocZ / lFar))

                        .LastUpdateCycle = glCurrentCycle

                        'Handle our y movement
                        '.HandleMoveY(iturnamt, .oUnitDef.Acceleration, lCurrTime)
                        .HandleMoveY(iTurnAmt, lCurrTime, .oUnitDef.Acceleration) '  Math.Max(1, fMaxSpeed))

                        'If it is not Planet to System or System to Planet, then the unit is either NOT doing
                        '  a change environment, or it is changing from system to system which doesn't require any
                        '  extra work on the movement engine's behalf.
                        'If .yChangeEnvironment = ChangeEnvironmentType.ePlanetToSystem Then
                        '    If oCurrEnvir.ObjTypeID = ObjectType.ePlanet Then
                        '        'Scale down
                        '        .LocY += CInt(.TotalVelocity)
                        '    Else
                        '        'Scale Up, Y should cap at 0
                        '        'fTotDist
                        '        '.LocY += CInt(.TotalVelocity)
                        '        'If .LocY > 0 Then .LocY = 0
                        '        If .fOriginalDist <> 0 Then
                        '            .LocY = CInt(-1000 * CSng(fTotDist / .fOriginalDist))
                        '        Else : .LocY = 0
                        '        End If
                        '    End If
                        'ElseIf .yChangeEnvironment = ChangeEnvironmentType.eSystemToPlanet Then
                        '    If oCurrEnvir.ObjTypeID = ObjectType.ePlanet Then
                        '        'scale up, Y should cap at 256 * ml_y_mult
                        '        'If oCurrEnvir.oGeoObject Is Nothing = False Then
                        '        '    If .LocY < 256 * oCurrEnvir.oGeoObject.ml_Y_Mult Then
                        '        '        .LocY = 256 * oCurrEnvir.oGeoObject.ml_Y_Mult
                        '        '    End If
                        '        'Else : .LocY = 4000
                        '        'End If
                        '    Else
                        '        'scale down
                        '        .LocY -= CInt(.TotalVelocity * 2)
                        '    End If
                        'ElseIf oCurrEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                        '    .LocY = 0       'TODO: Fail-safe, but not sure about it
                        'End If
                    ElseIf .LocYaw <> 0 AndAlso .LocAngle = .DestAngle Then
                        'straighten us out if we are not straight.
                        If .oUnitDef.TurnAmount > Math.Abs(.LocYaw) Then
                            'just set to 0
                            .LocYaw = 0
                        Else
                            If .LocYaw < 0 Then
                                .LocYaw += .oUnitDef.TurnAmount
                            Else
                                .LocYaw -= .oUnitDef.TurnAmount
                            End If
                        End If
                    End If      'are we moving?
                End With
            End If
        Next X

        gb_InHandleMovement = False

    End Sub

    Public Sub LoadAgentData()
        Dim sDATPak As String = "md.pak"
        Dim sFile As String = "Missions.dat"

        'ok, first, let's load up missions
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        Dim oINI As New InitFile(goResMgr.MeshPath & sFile)
        Dim bDone As Boolean = False
        Dim lIdx As Int32 = 0

        glMissionUB = -1
        While bDone = False
            Dim sHdr As String = "MISSION" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "MissionID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                glMissionUB += 1
                ReDim Preserve glMissionIdx(glMissionUB)
                ReDim Preserve goMissions(glMissionUB)
                glMissionIdx(glMissionUB) = lID
                goMissions(glMissionUB) = New Mission
                With goMissions(glMissionUB)
                    .ObjectID = lID
                    .ObjTypeID = ObjectType.eMission
                    .BaseEffect = CShort(oINI.GetString(sHdr, "BaseEffect", "0"))
                    .GoalUB = -1
                    .lInfiltrationType = CType(CInt(oINI.GetString(sHdr, "InfiltrationType", "0")), eInfiltrationType)
                    .Modifier = CShort(oINI.GetString(sHdr, "Modifier", "0"))
                    .ProgramControlID = CShort(oINI.GetString(sHdr, "ProgramControlID", "0"))
                    .sMissionDesc = oINI.GetString(sHdr, "MissionDesc", "").Replace("[VBCRLF]", vbCrLf)
                    .sMissionName = oINI.GetString(sHdr, "MissionName", "")
                End With
                lIdx += 1
            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)

        glSkillUB = -1
        lIdx = 0
        sFile = "Skills.dat"
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        oINI = New InitFile(goResMgr.MeshPath & sFile)
        bDone = False
        While bDone = False
            Dim sHdr As String = "SKILL" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "SkillID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                glSkillUB += 1
                ReDim Preserve glSkillIdx(glSkillUB)
                ReDim Preserve goSkills(glSkillUB)
                glSkillIdx(glSkillUB) = lID
                goSkills(glSkillUB) = New Skill
                With goSkills(glSkillUB)
                    .ObjectID = lID
                    .ObjTypeID = ObjectType.eSkill
                    .MaxVal = CByte(oINI.GetString(sHdr, "MaxVal", "98"))
                    .MinVal = CByte(oINI.GetString(sHdr, "MinVal", "0"))
                    .SkillDesc = oINI.GetString(sHdr, "SkillDesc", "").Replace("[VBCRLF]", vbCrLf)
                    .SkillName = oINI.GetString(sHdr, "SkillName", "")
                    .SkillType = CByte(oINI.GetString(sHdr, "SkillType", "0"))
                End With
                lIdx += 1
            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)

        'Now, Methods
        sFile = "Methods.dat"
        lIdx = 0
        glMissionMethodUB = -1
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        oINI = New InitFile(goResMgr.MeshPath & sFile)
        bDone = False
        While bDone = False
            Dim sHdr As String = "METHOD" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "MM_ID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                glMissionMethodUB += 1
                ReDim Preserve glMissionMethodIdx(glMissionMethodUB)
                ReDim Preserve gsMissionMethods(glMissionMethodUB)
                ReDim Preserve gsMissionDescs(glMissionMethodUB)
                glMissionMethodIdx(glMissionMethodUB) = lID
                gsMissionMethods(glMissionMethodUB) = oINI.GetString(sHdr, "MethodName", "").Replace("[VBCRLF]", vbCrLf)
                gsMissionDescs(glMissionMethodUB) = oINI.GetString(sHdr, "MethodDesc", "")
                lIdx += 1
            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)

        'Now, goals
        glGoalUB = -1
        sFile = "Goals.dat"
        lIdx = 0
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        oINI = New InitFile(goResMgr.MeshPath & sFile)
        bDone = False
        While bDone = False
            Dim sHdr As String = "GOAL" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "GoalID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                glGoalUB += 1
                ReDim Preserve glGoalIdx(glGoalUB)
                ReDim Preserve goGoals(glGoalUB)
                glGoalIdx(glGoalUB) = lID
                goGoals(glGoalUB) = New Goal
                With goGoals(glGoalUB)
                    .ObjectID = lID
                    .ObjTypeID = ObjectType.eGoal
                    .BaseTime = CInt(oINI.GetString(sHdr, "BaseTime", "0"))
                    .MissionPhase = CType(CInt(oINI.GetString(sHdr, "MissionPhase", "0")), eMissionPhase)
                    .SuccessProgCtrlID = CInt(oINI.GetString(sHdr, "SuccessProgCtrlID", "0"))
                    .FailureProgCtrlID = CInt(oINI.GetString(sHdr, "FailureProgCtrlID", "0"))
                    .RiskOfDetection = CByte(oINI.GetString(sHdr, "RiskOfDetection", "0"))
                    .sGoalDesc = oINI.GetString(sHdr, "GoalDesc", "").Replace("[VBCRLF]", vbCrLf)
                    .sGoalName = oINI.GetString(sHdr, "GoalName", "").Replace("[VBCRLF]", vbCrLf)
                    .SkillSetUB = -1
                End With
                lIdx += 1
            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)

        'Now, Mission Goals
        sFile = "MGoal.dat"
        lIdx = 0
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        oINI = New InitFile(goResMgr.MeshPath & sFile)
        bDone = False
        While bDone = False
            Dim sHdr As String = "MG" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "GoalID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                Dim lMissionID As Int32 = CInt(oINI.GetString(sHdr, "MissionID", "-1"))
                Dim lMethodID As Int32 = CInt(oINI.GetString(sHdr, "MethodID", "-1"))
                If lMissionID <> -1 Then
                    Dim oGoal As Goal = Nothing
                    For Y As Int32 = 0 To glGoalUB
                        If glGoalIdx(Y) = lID Then
                            oGoal = goGoals(Y)
                            Exit For
                        End If
                    Next Y
                    If oGoal Is Nothing = False Then
                        For X As Int32 = 0 To glMissionUB
                            If glMissionIdx(X) = lMissionID Then
                                goMissions(X).GoalUB += 1
                                ReDim Preserve goMissions(X).Goals(goMissions(X).GoalUB)
                                ReDim Preserve goMissions(X).MethodIDs(goMissions(X).GoalUB)

                                goMissions(X).Goals(goMissions(X).GoalUB) = oGoal
                                goMissions(X).MethodIDs(goMissions(X).GoalUB) = lMethodID
                                Exit For
                            End If
                        Next X
                    End If
                End If

                lIdx += 1
            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)

        'Now, Skillsets
        lIdx = 0
        sFile = "Skillsets.dat"
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        oINI = New InitFile(goResMgr.MeshPath & sFile)
        bDone = False
        While bDone = False
            Dim sHdr As String = "SKILLSET" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "SkillsetID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                Dim lGoalID As Int32 = CInt(oINI.GetString(sHdr, "GoalID", "-1"))
                Dim oSkillSet As New SkillSet

                With oSkillSet
                    .ProgramControlID = CInt(oINI.GetString(sHdr, "ProgramControlID", "0"))
                    .SkillSetID = lID
                    .SkillUB = -1
                    .sSkillSetName = oINI.GetString(sHdr, "SkillsetName", "")
                End With

                For X As Int32 = 0 To glGoalUB
                    If glGoalIdx(X) = lGoalID Then
                        oSkillSet.oGoal = goGoals(X)
                        goGoals(X).AddSkillSet(oSkillSet)
                        Exit For
                    End If
                Next X

                lIdx += 1
            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)

        'Now, Skillset_Skills
        lIdx = 0
        sFile = "SSSkills.dat"
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        oINI = New InitFile(goResMgr.MeshPath & sFile)
        bDone = False
        While bDone = False
            Dim sHdr As String = "SSSKILL" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "SkillsetID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                'increment our idx now because we will continue while inside the logic here
                lIdx += 1

                Dim oSkill As Skill = Nothing
                Dim lSkillID As Int32 = CInt(oINI.GetString(sHdr, "SkillID", "-1"))
                If lSkillID <> -1 Then
                    For X As Int32 = 0 To glSkillUB
                        If glSkillIdx(X) = lSkillID Then
                            oSkill = goSkills(X)
                            Exit For
                        End If
                    Next X
                End If
                If oSkill Is Nothing Then Continue While

                'Ok, need to find our skillset
                For X As Int32 = 0 To glGoalUB
                    If glGoalIdx(X) <> -1 Then
                        For Y As Int32 = 0 To goGoals(X).SkillSetUB
                            If goGoals(X).SkillSets(Y).SkillSetID = lID Then
                                'ok, found it
                                Dim oSSS As SkillSet_Skill = goGoals(X).SkillSets(Y).AddSkill(oSkill)
                                With oSSS
                                    .FailureProgCtrlID = CInt(oINI.GetString(sHdr, "FailureProgCtrlID", "0"))
                                    .SuccessProgCtrlID = CInt(oINI.GetString(sHdr, "SuccessProgCtrlID", "0"))
                                    .ToHitModifier = CShort(oINI.GetString(sHdr, "ToHitModifier", "0"))
                                    .AgentGroupID = CShort(oINI.GetString(sHdr, "AgentGroupingID", "0"))
                                    .PointRequirement = CInt(oINI.GetString(sHdr, "PointRequirement", "0"))
                                End With

                                Continue While
                            End If
                        Next Y
                    End If
                Next X

            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)
    End Sub

    Public Sub LoadStarTypes()
        Dim sDATPak As String = "md.pak"
        Dim sFile As String = "StarTypes.dat"

        'ok, first, let's load up missions
        If goResMgr.UnpackLocalDataFile(sDATPak, sFile) = False Then
            '?
        End If
        Dim oINI As New InitFile(goResMgr.MeshPath & sFile)
        Dim bDone As Boolean = False
        Dim lIdx As Int32 = 0

        glStarTypeUB = -1

        While bDone = False
            Dim sHdr As String = "STARTYPE" & lIdx
            Dim lID As Int32 = CInt(oINI.GetString(sHdr, "StarTypeID", "-1"))
            If lID = -1 Then
                bDone = True
            Else
                glStarTypeUB += 1
                ReDim Preserve goStarTypes(glStarTypeUB)
                goStarTypes(glStarTypeUB) = New StarType()
                With goStarTypes(glStarTypeUB)
                    .HeatIndex = CByte(Val(oINI.GetString(sHdr, "HeadIndex", "0")))
                    .LightAmbient = CInt(Val(oINI.GetString(sHdr, "LightAmbient", "0")))
                    .LightAtt0 = CSng(Val(oINI.GetString(sHdr, "LightAtt1", "1")))
                    .LightAtt1 = CSng(Val(oINI.GetString(sHdr, "LightAtt2", "0")))
                    .LightAtt2 = CSng(Val(oINI.GetString(sHdr, "LightAtt3", "0")))
                    .LightDiffuse = CInt(Val(oINI.GetString(sHdr, "LightDiffuse", "0")))
                    .LightRange = CInt(Val(oINI.GetString(sHdr, "LightRange", "0")))
                    .LightSpecular = CInt(Val(oINI.GetString(sHdr, "LightSpecular", "0")))
                    .MatDiffuse = CInt(Val(oINI.GetString(sHdr, "MatDiffuse", "0")))
                    .MatEmissive = CInt(Val(oINI.GetString(sHdr, "MatEmissive", "0")))
                    .StarAttributes = CInt(Val(oINI.GetString(sHdr, "StarTypeAttrs", "0")))
                    .StarTexture = oINI.GetString(sHdr, "StarTexture", "StarMain.bmp")
                    .StarTypeID = CByte(lID)
                    .StarTypeName = oINI.GetString(sHdr, "StarTypeName", "")
                    .StarRadius = CInt(Val(oINI.GetString(sHdr, "StarRadius", "")))
                    .lStarMapRectIdx = CInt(Val(oINI.GetString(sHdr, "StarMapRectIdx", "0")))
                End With

                lIdx += 1
            End If
        End While
        oINI = Nothing
        If Exists(goResMgr.MeshPath & sFile) = True Then Kill(goResMgr.MeshPath & sFile)
    End Sub

    Public Sub SortMineralProperties()
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        For X As Int32 = 0 To glMineralPropertyUB
            Dim lIdx As Int32 = -1
            If glMineralPropertyIdx(X) <> -1 Then
                For Y As Int32 = 0 To lSortedUB
                    If goMineralProperty(lSorted(Y)).MineralPropertyName.ToUpper > goMineralProperty(X).MineralPropertyName.ToUpper Then
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
            End If
        Next X

        Dim lNewIdx(glMineralPropertyUB) As Int32
        Dim oNewObj(glMineralPropertyUB) As MineralProperty

        For X As Int32 = 0 To glMineralPropertyUB
            lNewIdx(X) = glMineralPropertyIdx(lSorted(X))
            oNewObj(X) = goMineralProperty(lSorted(X))
        Next X

        glMineralPropertyIdx = lNewIdx
        goMineralProperty = oNewObj

    End Sub
#End Region

#Region " Geography Instantiations "
    Public goStarTypes() As StarType
    Public glStarTypeUB As Int32 = -1

    Public goGalaxy As Galaxy       'the one, the only...
#End Region

#Region " Cache Handlers and Definitions "
    'This structure is for caching detailed lists... for example, Player ID to Player Name or Unit Type ID to Unit Type Name
    Private Structure CacheEntry
        Public lID As Int32
        Public iTypeID As Int16
        Public iExtTypeID As Int16
        Public sValue As String
        Public bRequested As Boolean
    End Structure

    Private muCacheNewMethod As Boolean = False
    Private muCacheIDX As New Hashtable
    Private muCache() As CacheEntry
    Private mlCacheUB As Int32 = -1
    Private mlRequestsOutstanding As Int32 = 0
    Private mbInProcessGetCacheEntries As Boolean = False

    Public Function GetCacheSize() As Int32
        Return mlCacheUB + 1
    End Function

    Public Sub ProcessGetCacheEntries()
        mbInProcessGetCacheEntries = True
        'Ok, we go through our cache list and see if any need requested
        If mlRequestsOutstanding > 0 AndAlso goUILib Is Nothing = False Then
            Dim lCnt As Int32 = mlRequestsOutstanding
            If lCnt * 6 + 5 > 32000 Then lCnt = (32000 - 5) \ 6

            Dim yMsg(5 + (lCnt * 6)) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

            For X As Int32 = 0 To mlCacheUB
                If muCache(X).bRequested = False Then
                    mlRequestsOutstanding -= 1
                    lCnt -= 1
                    System.BitConverter.GetBytes(muCache(X).lID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(muCache(X).iTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    muCache(X).bRequested = True
                    If lCnt = 0 Then Exit For
                End If
            Next X
            goUILib.SendMsgToPrimary(yMsg)

            'If mlRequestsOutstanding Then
            mlRequestsOutstanding = 0
            If mlCacheUB > UBound(muCache) Then mlCacheUB = UBound(muCache)
            For X As Int32 = 0 To mlCacheUB
                If muCache(X).bRequested = False Then mlRequestsOutstanding += 1
            Next X
            'End If
        End If

        mbInProcessGetCacheEntries = False
    End Sub

    Public Function GetCacheObjectValueCheckUnknowns(ByVal lID As Int32, ByVal iTypeID As Int16, ByRef bUnknown As Boolean) As String
        Dim sReturn As String
        sReturn = GetCacheObjectValue(lID, iTypeID)
        If sReturn = "Unknown" Then bUnknown = True Else bUnknown = False
        Return sReturn
    End Function

    Public Function GetCacheObjectValue(ByVal lID As Int32, ByVal iTypeID As Int16) As String
        While mbInProcessGetCacheEntries = True
            Threading.Thread.Sleep(10)
        End While

        If muCacheNewMethod = True Then
            Return GetCacheObjectValue_Hashed(lID, iTypeID)
        End If

        Dim X As Int32
        Dim lIdx As Int32 = -1

        'Ok, this works differently... find our Idx
        Try
            For X = 0 To mlCacheUB
                If muCache(X).lID = lID AndAlso muCache(X).iTypeID = iTypeID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                'add the cache object
                mlCacheUB += 1
                ReDim Preserve muCache(mlCacheUB)
                lIdx = mlCacheUB
                muCache(lIdx).lID = lID
                muCache(lIdx).iTypeID = iTypeID
                muCache(lIdx).bRequested = False

                Select Case iTypeID
                    Case ObjectType.eColonists
                        muCache(lIdx).sValue = "Colonists"
                    Case ObjectType.eEnlisted
                        muCache(lIdx).sValue = "Enlisted"
                    Case ObjectType.eOfficers
                        muCache(lIdx).sValue = "Officers"
                    Case Else
                        muCache(lIdx).sValue = "Unknown"
                        mlRequestsOutstanding += 1
                End Select
            End If

            If muCache(lIdx).bRequested = False Then
                'Ok, first, see if it is a technology of the player's
                If goCurrentPlayer Is Nothing = False Then
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
                    If oTech Is Nothing = False Then
                        muCache(lIdx).bRequested = True
                        muCache(lIdx).sValue = oTech.GetComponentName()
                        mlRequestsOutstanding -= 1
                    End If
                End If
            End If

            Return muCache(lIdx).sValue
        Catch
        End Try
        Return "Unknown"
    End Function

    Public Function GetCacheObjectValue_Hashed(ByVal lID As Int32, ByVal iTypeID As Int16) As String
        Dim lIdx As Int32 = -1
        Try
            Dim sKey As String = iTypeID.ToString & "-" & lID.ToString
            If muCacheIDX.Contains(sKey) = True Then
                lIdx = CInt(muCacheIDX(sKey))
            Else
                'add the cache object
                mlCacheUB += 1
                ReDim Preserve muCache(mlCacheUB)
                lIdx = mlCacheUB
                muCache(lIdx).lID = lID
                muCache(lIdx).iTypeID = iTypeID
                muCache(lIdx).bRequested = False
                muCacheIDX.Add(sKey, lIdx)

                Select Case iTypeID
                    Case ObjectType.eColonists
                        muCache(lIdx).sValue = "Colonists"
                    Case ObjectType.eEnlisted
                        muCache(lIdx).sValue = "Enlisted"
                    Case ObjectType.eOfficers
                        muCache(lIdx).sValue = "Officers"
                    Case Else
                        muCache(lIdx).sValue = "Unknown"
                        mlRequestsOutstanding += 1
                End Select
            End If

            If muCache(lIdx).bRequested = False Then
                'Ok, first, see if it is a technology of the player's
                If goCurrentPlayer Is Nothing = False Then
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
                    If oTech Is Nothing = False Then
                        muCache(lIdx).bRequested = True
                        muCache(lIdx).sValue = oTech.GetComponentName()
                        mlRequestsOutstanding -= 1
                    End If
                End If
            ElseIf mlRequestsOutstanding = 0 Then
                'We have missing items but no outstanging.  Should not happen but does.
                'TODO: Test more before blinding adding this
                'mlRequestsOutstanding += 1
            End If
            Return muCache(lIdx).sValue
        Catch
        End Try
        Return "Unknown"
    End Function

    Public Sub SetCacheObjectValue(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal sValue As String)
        If muCacheNewMethod = True Then
            SetCacheObjectValue_Hashed(lID, iTypeID, sValue)
            Exit Sub
        End If

        Dim X As Int32
        Dim lIdx As Int32 = -1

        'set our value
        For X = 0 To mlCacheUB
            If muCache(X).lID = lID AndAlso muCache(X).iTypeID = iTypeID Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlCacheUB += 1
            ReDim Preserve muCache(mlCacheUB)
            lIdx = mlCacheUB

            muCache(lIdx).lID = lID
            muCache(lIdx).iTypeID = iTypeID
            muCache(lIdx).bRequested = True
        End If

        muCache(lIdx).sValue = sValue
    End Sub

    Public Sub SetCacheObjectValue_Hashed(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal sValue As String)
        Dim lIdx As Int32 = -1
        Try
            Dim sKey As String = iTypeID.ToString & "-" & lID.ToString
            If muCacheIDX.Contains(sKey) Then
                lIdx = CInt(muCacheIDX(sKey))
            Else
                mlCacheUB += 1
                ReDim Preserve muCache(mlCacheUB)
                lIdx = mlCacheUB
                muCache(lIdx).lID = lID
                muCache(lIdx).iTypeID = iTypeID
                muCache(lIdx).bRequested = True
                muCache(lIdx).sValue = sValue
                muCacheIDX.Add(sKey, lIdx)
            End If
            muCache(lIdx).sValue = sValue
        Catch
        End Try
    End Sub

    Public Sub SortCollection(ByVal col As Collection, ByVal psSortPropertyName As String, ByVal pbAscending As Boolean, Optional ByVal psKeyPropertyName As String = "")
        Try
            ' The Objects were originally declared as Variants. VB.Net has 
            'eliminated the
            ' Variant type so they must be declared as type Object. 
            'Also Objects cannot be
            'used with the Set keyword, so I had to remove the set 
            'keyword. Other than that
            'I did not have to make hardly any changes.

            Dim obj As Object
            Dim i As Integer
            Dim j As Integer
            Dim iMinMaxIndex As Integer
            Dim vMinMax As Object
            Dim vValue As Object
            Dim bSortCondition As Boolean
            Dim bUseKey As Boolean
            Dim sKey As String

            bUseKey = (psKeyPropertyName <> "")

            For i = 1 To col.Count - 1
                obj = col(i)
                ' the vbGet can be replaced with a 
                'CallType.Get if you
                ' want. See VB Language reference for CallByName

                vMinMax = CallByName(obj, psSortPropertyName, vbGet)
                iMinMaxIndex = i

                For j = i + 1 To col.Count
                    obj = col(j)
                    vValue = CallByName(obj, _
                        psSortPropertyName, vbGet)

                    If (pbAscending) Then
                        bSortCondition = (vValue.ToString < vMinMax.ToString)
                    Else
                        bSortCondition = (vValue.ToString > vMinMax.ToString)
                    End If

                    If (bSortCondition) Then
                        vMinMax = vValue
                        iMinMaxIndex = j
                    End If

                    obj = Nothing
                Next j

                If (iMinMaxIndex <> i) Then
                    obj = col(iMinMaxIndex)

                    col.Remove(iMinMaxIndex)
                    If (bUseKey) Then
                        sKey = CStr(CallByName(obj, _
                           psKeyPropertyName, vbGet))
                        col.Add(obj, sKey, i)
                    Else
                        col.Add(obj, , i)
                    End If

                    obj = Nothing
                End If

                obj = Nothing
            Next i
        Catch
        End Try
    End Sub

    'NonOwnerItemData here too
    Private muNonOwnerItemData() As CacheEntry
    Private mlNonOwnerItemDataUB As Int32 = -1

    Public Function GetNonOwnerItemData(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16) As String
        Dim X As Int32
        'Dim yData() As Byte
        Dim lIdx As Int32 = -1

        'Ok, this works differently... find our Idx
        For X = 0 To mlNonOwnerItemDataUB
            If muNonOwnerItemData(X).lID = lObjectID AndAlso muNonOwnerItemData(X).iTypeID = iObjTypeID Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            'add the cache object
            mlNonOwnerItemDataUB += 1
            ReDim Preserve muNonOwnerItemData(mlNonOwnerItemDataUB)
            lIdx = mlNonOwnerItemDataUB
            muNonOwnerItemData(lIdx).sValue = "Requesting Details..."
            muNonOwnerItemData(lIdx).lID = lObjectID
            muNonOwnerItemData(lIdx).iTypeID = iObjTypeID
        End If

        If muNonOwnerItemData(lIdx).bRequested = False Then
            'Ok, first, see if it is a technology of the player's
            'ok, request the value
            If goUILib Is Nothing = False Then
                Dim yRequest(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yRequest, 0)
                System.BitConverter.GetBytes(lObjectID).CopyTo(yRequest, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yRequest, 6)
                goUILib.SendMsgToPrimary(yRequest)
                muNonOwnerItemData(lIdx).bRequested = True
            End If
        End If

        Return muNonOwnerItemData(lIdx).sValue
    End Function

    Public Sub SetNonOwnerItemData(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal sValue As String)
        Dim X As Int32
        Dim lIdx As Int32 = -1

        'set our value
        Try
            For X = 0 To mlNonOwnerItemDataUB
                If muNonOwnerItemData(X).lID = lID AndAlso muNonOwnerItemData(X).iTypeID = iTypeID Then
                    lIdx = X
                    Exit For
                End If
            Next X
        Catch
        End Try

        If lIdx = -1 Then
            mlNonOwnerItemDataUB += 1
            ReDim Preserve muNonOwnerItemData(mlNonOwnerItemDataUB)
            lIdx = mlNonOwnerItemDataUB

            muNonOwnerItemData(lIdx).lID = lID
            muNonOwnerItemData(lIdx).iTypeID = iTypeID
            muNonOwnerItemData(lIdx).bRequested = True
        End If

        muNonOwnerItemData(lIdx).sValue = sValue
    End Sub

    Public Function GetNonOwnerIntelItemData(ByVal lTradepostID As Int32, ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal iExtTypeID As Int16) As String
        Dim X As Int32
        Dim lIdx As Int32 = -1

        'Ok, this works differently... find our Idx
        For X = 0 To mlNonOwnerItemDataUB
            If muNonOwnerItemData(X).lID = lObjectID AndAlso muNonOwnerItemData(X).iTypeID = iObjTypeID AndAlso muNonOwnerItemData(X).iExtTypeID = iExtTypeID Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            'add the cache object
            mlNonOwnerItemDataUB += 1
            ReDim Preserve muNonOwnerItemData(mlNonOwnerItemDataUB)
            lIdx = mlNonOwnerItemDataUB
            muNonOwnerItemData(lIdx).sValue = "Requesting Details..."
            muNonOwnerItemData(lIdx).lID = lObjectID
            muNonOwnerItemData(lIdx).iTypeID = iObjTypeID
            muNonOwnerItemData(lIdx).iExtTypeID = iExtTypeID
        End If

        If muNonOwnerItemData(lIdx).bRequested = False Then
            'Ok, first, see if it is a technology of the player's
            'ok, request the value
            If goUILib Is Nothing = False Then
                Dim yRequest(13) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetIntelSellOrderDetail).CopyTo(yRequest, 0)
                System.BitConverter.GetBytes(lTradepostID).CopyTo(yRequest, 2)
                System.BitConverter.GetBytes(lObjectID).CopyTo(yRequest, 6)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yRequest, 10)
                System.BitConverter.GetBytes(iExtTypeID).CopyTo(yRequest, 12)
                goUILib.SendMsgToPrimary(yRequest)
                muNonOwnerItemData(lIdx).bRequested = True
            End If
        End If

        Return muNonOwnerItemData(lIdx).sValue
    End Function

    Public Sub SetNonOwnerIntelItemData(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal iExtTypeID As Int16, ByVal sValue As String)
        Dim X As Int32
        Dim lIdx As Int32 = -1

        'set our value
        Try
            For X = 0 To mlNonOwnerItemDataUB
                If muNonOwnerItemData(X).lID = lID AndAlso muNonOwnerItemData(X).iTypeID = iTypeID AndAlso muNonOwnerItemData(X).iExtTypeID = iExtTypeID Then
                    lIdx = X
                    Exit For
                End If
            Next X
        Catch
        End Try

        If lIdx = -1 Then
            mlNonOwnerItemDataUB += 1
            ReDim Preserve muNonOwnerItemData(mlNonOwnerItemDataUB)
            lIdx = mlNonOwnerItemDataUB

            muNonOwnerItemData(lIdx).lID = lID
            muNonOwnerItemData(lIdx).iTypeID = iTypeID
            muNonOwnerItemData(lIdx).iExtTypeID = iExtTypeID
            muNonOwnerItemData(lIdx).bRequested = True
        End If

        muNonOwnerItemData(lIdx).sValue = sValue
    End Sub

    'Private Class UnitWarpointUpkeepItem
    '    Public lUnitID As Int32
    '    Public lValue As Int32
    '    Public bRequested As Boolean
    'End Class
    'Private mcolUnitWarpointUpkeep As New Collection
    'Public Function GetUnitWarpointUpkeep(ByVal lUnitID As Int32) As Int32
    '    Dim sKey As String = "U" & lUnitID.ToString
    '    If mcolUnitWarpointUpkeep.Contains(sKey) = True Then
    '        Dim oItem As UnitWarpointUpkeepItem = CType(mcolUnitWarpointUpkeep(sKey), UnitWarpointUpkeepItem)
    '        If oItem Is Nothing = False Then
    '            If oItem.bRequested = True Then Return oItem.lValue Else Return -1I
    '        End If
    '    End If

    '    Dim oNew As New UnitWarpointUpkeepItem()
    '    oNew.lUnitID = lUnitID
    '    oNew.lValue = -1I
    '    oNew.bRequested = True
    '    mcolUnitWarpointUpkeep.Add(oNew, sKey)

    '    'Dim yMsg(7) As Byte
    '    'Dim lPos As Int32 = 0
    '    'System.BitConverter.GetBytes(GlobalMessageCode.eRequestWarpointValue).CopyTo(yMsg, lPos) : lPos += 2
    '    'System.BitConverter.GetBytes(lUnitID).CopyTo(yMsg, lPos) : lPos += 4
    '    'System.BitConverter.GetBytes(ObjectType.eUnit).CopyTo(yMsg, lPos) : lPos += 2
    '    'goUILib.SendMsgToPrimary(yMsg)

    '    Return -1
    'End Function
    'Public Sub SetUnitWarpointUpkeep(ByVal lUnitID As Int32, ByVal lVal As Int32)
    '    Dim sKey As String = "U" & lUnitID.ToString
    '    If mcolUnitWarpointUpkeep.Contains(sKey) = True Then
    '        Dim oItem As UnitWarpointUpkeepItem = CType(mcolUnitWarpointUpkeep(sKey), UnitWarpointUpkeepItem)
    '        If oItem Is Nothing = False Then
    '            oItem.lValue = Math.Max(1, lVal)
    '            Return
    '        End If
    '    End If
    '    'mcolUnitWarpointUpkeep.Add(lVal, sKey)
    '    Dim oNew As New UnitWarpointUpkeepItem()
    '    oNew.lUnitID = lUnitID
    '    oNew.lValue = Math.Max(1, lVal)
    '    oNew.bRequested = True
    '    mcolUnitWarpointUpkeep.Add(oNew, sKey)
    'End Sub
#End Region

#Region " Trigonometric/Geometric Functions "

    Public Function RadianToDegree(ByVal fRads As Single) As Single
        Return fRads * gdDegreePerRad
    End Function

    Public Function DegreeToRadian(ByVal fDegree As Single) As Single
        Return fDegree * gdRadPerDegree
    End Function

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Int32, ByRef lEndY As Int32, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * CSng(Math.PI / 180.0F)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CInt((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CInt((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Single, ByRef lEndY As Single, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * CSng(Math.PI / 180.0F)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CSng((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CSng((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Sub RotatePoint(ByVal fAxisX As Single, ByVal fAxisY As Single, ByRef fEndX As Single, ByRef fEndY As Single, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * CSng(Math.PI / 180.0F)
        dDX = fEndX - fAxisX
        dDY = fEndY - fAxisY
        fEndX = fAxisX + CSng((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        fEndY = fAxisY + -CSng((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Function LineAngleDegrees(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
        Dim dDeltaX As Single
        Dim dDeltaY As Single
        Dim dAngle As Single
        'Return 0

        dDeltaX = lX2 - lX1
        dDeltaY = lY2 - lY1

        If dDeltaX = 0 Then     'vertical
            If dDeltaY < 0 Then
                dAngle = gdHalfPie
            Else
                dAngle = gdPieAndAHalf
            End If
        ElseIf dDeltaY = 0 Then     'horizontal
            If dDeltaX < 0 Then
                dAngle = gdPi
            Else
                dAngle = 0
            End If
        Else    'angled
            dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))
            'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
            If dDeltaX > -1 And dDeltaY > -1 Then       'VB Lower Right
                dAngle = gdTwoPie - dAngle
            ElseIf dDeltaX < 0 And dDeltaY > -1 Then    'VB Lower Left
                dAngle = gdPi + dAngle
            ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left
                dAngle = gdPi - dAngle
            End If
        End If

        'Not sure this is suppose to be CINT
        'Return CInt(dAngle * gdDegreePerRad)
        Return dAngle * gdDegreePerRad

    End Function

    Public Function LineAngleDegreesF(ByVal lX1 As Single, ByVal lY1 As Single, ByVal lX2 As Single, ByVal lY2 As Single) As Single
        Dim dDeltaX As Single
        Dim dDeltaY As Single
        Dim dAngle As Single

        dDeltaX = lX2 - lX1
        dDeltaY = lY2 - lY1

        If dDeltaX = 0 Then     'vertical
            If dDeltaY < 0 Then
                dAngle = gdHalfPie
            Else
                dAngle = gdPieAndAHalf
            End If
        ElseIf dDeltaY = 0 Then     'horizontal
            If dDeltaX < 0 Then
                dAngle = gdPi
            Else
                dAngle = 0
            End If
        Else    'angled
            dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))
            'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
            If dDeltaX > -1 And dDeltaY > -1 Then       'VB Lower Right
                dAngle = gdTwoPie - dAngle
            ElseIf dDeltaX < 0 And dDeltaY > -1 Then    'VB Lower Left
                dAngle = gdPi + dAngle
            ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left
                dAngle = gdPi - dAngle
            End If
        End If

        Return dAngle * gdDegreePerRad

    End Function

    Public Function Distance(ByVal lX1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lZ2 As Int32) As Single
        Dim fX As Single = lX2 - lX1
        Dim fZ As Single = lZ2 - lZ1
        fX *= fX
        fZ *= fZ
        Return CSng(Math.Sqrt(fX + fZ))
    End Function

    Public Function AngleToQuadrant(ByVal lAngle As Int32) As Byte
        'here, we will return the quadrant from the angle
        Select Case lAngle
            Case Is < 45, Is > 315
                Return UnitArcs.eForwardArc
            Case Is < 135
                'Return UnitArcs.eLeftArc
                Return UnitArcs.eRightArc
            Case Is < 225
                Return UnitArcs.eBackArc
            Case Else
                'Return UnitArcs.eRightArc
                Return UnitArcs.eLeftArc
        End Select
    End Function
#End Region

#Region " MEDIA INSTANTIATIONS "
    Public goResMgr As GFXResourceManager
    Public WithEvents goWpnMgr As WpnFXManager
    Public goShldMgr As ShieldFXManager
    Public goExplMgr As ExplosionManager 'ExplosionFXManager
    Public goPFXEngine32 As BurnFX.ParticleEngine
    Public goEntityDeath As DeathSequenceMgr
    Public goMissileMgr As MissileMgr
    Public goBurnMarkMgr As EntityBurnMarkManager
    'Public goRewards As WarpointRewards
    Public goFireworks As FireworksMgr

    Public goWormholeMgr As WormholeManager

    Public goCamera As Camera

    Public goSound As SoundMgr        'the global instance of the sound object

    Public goUILib As UILib

    Public gclrEngines() As Color
    Public gclrShields() As Color
    Public Sub InitializeFXColors()
        ReDim gclrEngines(15)
        ReDim gclrShields(15)

        'Set all to the default
        For X As Int32 = 0 To 15
            gclrEngines(X) = System.Drawing.Color.FromArgb(255, 64, 128, 255)
            gclrShields(X) = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        Next X

        'Now, set the Engines values
        gclrEngines(1) = System.Drawing.Color.FromArgb(255, 64, 255, 64)
        gclrEngines(2) = System.Drawing.Color.FromArgb(255, 255, 128, 64)
        gclrEngines(3) = System.Drawing.Color.FromArgb(255, 64, 64, 255)
        gclrEngines(4) = System.Drawing.Color.FromArgb(255, 192, 64, 255)
        gclrEngines(5) = System.Drawing.Color.FromArgb(255, 64, 64, 92)
        gclrEngines(6) = System.Drawing.Color.FromArgb(255, 255, 64, 64)
        gclrEngines(7) = System.Drawing.Color.FromArgb(255, 255, 255, 32)
        gclrEngines(8) = System.Drawing.Color.FromArgb(255, 32, 128, 64)

        'Now for shields
        gclrShields(1) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        gclrShields(2) = System.Drawing.Color.FromArgb(255, 255, 255, 0)
        gclrShields(3) = System.Drawing.Color.FromArgb(255, 255, 128, 0)
        gclrShields(4) = System.Drawing.Color.FromArgb(255, 0, 0, 255)
        gclrShields(5) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        gclrShields(6) = System.Drawing.Color.FromArgb(255, 255, 0, 255)

    End Sub

    Public Sub goWpnMgr_WeaponEnd(ByVal oTarget As BaseEntity, ByVal yWeaponType As Byte, ByVal bHit As Boolean, ByVal vecOnModelTo As Microsoft.DirectX.Vector3) Handles goWpnMgr.WeaponEnd
        'The wpn's SFX will end by default... we do not need to worry about stopping it here...
        '  HOWEVER, *IF* we change our minds, stopping a related SFX would be done here
        Dim lTemp As Int32
        Dim oTmpVec As Microsoft.DirectX.Vector3 = New Microsoft.DirectX.Vector3(oTarget.LocX, oTarget.LocY, oTarget.LocZ)

        If bHit = True Then
            If oTarget.yVisibility = eVisibilityType.Visible Then
                If oTarget.yShieldHP <> 0 Then
                    If goShldMgr Is Nothing = False Then
                        goShldMgr.AddNewEffect(CType(oTarget, RenderObject), oTarget.clrShieldFX)
                    End If
                    If goSound Is Nothing = False Then
                        If goSound.MusicOn = True AndAlso oTarget.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(20) 'hit shields +20

                        'Two shield sound FX files
                        lTemp = CInt(Int(Rnd() * 2) + 1)
                        'Once and done SFX, we dont care about retreiving the id
                        goSound.StartSound("Explosions\ShieldHit" & lTemp & ".wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, oTmpVec, New Microsoft.DirectX.Vector3(0, 0, 0))
                    End If
                Else
                    'Ok, if the object is visible

                    'If goExplMgr Is Nothing = False Then
                    '    'TODO: it would be real cool to pass in a special ID for critical hits... the last parm allows this
                    '    goExplMgr.AddNewEffect(CInt(oTarget.LocX), oTarget.LocY + oTarget.oMesh.YMidPoint, CInt(oTarget.LocZ), 0)
                    'End If

                    If goSound Is Nothing = False Then
                        If goSound.MusicOn = True AndAlso oTarget.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(30) 'hit armor +30

                        'TODO: There are various explosions, but right now, there's only 4 and are all of the same type
                        'lTemp = CInt(Int(GetNxtRnd() * 4) + 1)
                        'Once and done SFX, we don't care about the ID
                        'goSound.StartSound("Explosions\HullHit" & lTemp & ".wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, oTmpVec, New Microsoft.DirectX.Vector3(0, 0, 0))
                    End If
                End If

                If goExplMgr Is Nothing = False Then
                    Dim vecFinal As Microsoft.DirectX.Vector3 '= Microsoft.DirectX.Vector3.TransformCoordinate(vecOnModelTo, oTarget.GetWorldMatrix)
                    With oTarget.GetWorldMatrix()
                        Dim fX As Single = vecOnModelTo.X
                        Dim fY As Single = vecOnModelTo.Y
                        Dim fZ As Single = vecOnModelTo.Z
                        vecFinal.X = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                        vecFinal.Y = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                        vecFinal.Z = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                    End With
                    Dim clrBright As System.Drawing.Color = ExplosionManager.GetWeaponTypeBrightColor(yWeaponType)
                    Dim fSize As Single = Rnd() * 20 + 100
                    goExplMgr.Add(vecFinal, Rnd() * 360, Rnd(), CInt(Rnd() * 4), fSize, 0, Color.White, clrBright, 30, fSize * 2, True)
                End If

            Else
                If goSound Is Nothing = False AndAlso goSound.MusicOn = True AndAlso oTarget.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(10) 'hit invisible +10
                If goExplMgr Is Nothing = False Then
                    Dim clrBright As System.Drawing.Color = ExplosionManager.GetWeaponTypeBrightColor(yWeaponType)
                    Dim fSize As Single = Rnd() * 20 + 65
                    goExplMgr.Add(vecOnModelTo, Rnd() * 360, Rnd(), CInt(Rnd() * 4), fSize, 0, Color.White, clrBright, 30, fSize * 2, True)
                End If
            End If
        End If
    End Sub

#End Region

#Region "  Performance Monitoring  "
    Public gbMonitorPerformance As Boolean = False
    'Update Camera
    Public gsw_Camera As Stopwatch
    'Update Movement
    Public gsw_Movement As Stopwatch
    'Render Geography
    Public gsw_Geography As Stopwatch
    Public glQuadsRendered As Int32
    'Render Models
    Public gsw_Models As Stopwatch
    Public glModelsRendered As Int32
    'Render Mineral Caches
    Public gsw_Caches As Stopwatch
    Public glCachesRendered As Int32
    'Render System Minimap
    Public gsw_Minimap As Stopwatch
    Public glMinimapItemsRendered As Int32
    'Render Wpn FX
    Public gsw_WpnFX As Stopwatch
    Public glWpnFXRendered As Int32
    'Render Bomb FX
    Public gsw_BombFX As Stopwatch
    Public glBombFXRendered As Int32
    'Render Shield FX
    Public gsw_ShieldFX As Stopwatch
    Public glShieldFXRendered As Int32
    'Render Burn FX
    Public gsw_BurnFX As Stopwatch
    Public glBurnFXRendered As Int32
    'Render Death FX
    Public gsw_DeathFX As Stopwatch
    'Render HPBars and WPInds
    Public gsw_HPBars As Stopwatch
    'Render User Interfaces
    Public gsw_UI As Stopwatch
    'Post effects
    Public gsw_PostEffects As Stopwatch
    'Preset
    Public gsw_Present As Stopwatch
    'Render Fireworks
    Public gsw_FireworksFX As Stopwatch
    Public glFireworksRendered As Int32
    'explosions
    Public gsw_Explosions As Stopwatch
    Public glExplosionRendered As Int32
    'Entity sounds
    Public gsw_CommitEntitySoundChanges As Stopwatch
    Public glUI_Rendered As Int32 = 0
    'AOE Explosions
    Public gsw_AOEExplosions As Stopwatch
    Public glAOEExplRendered As Int32
#End Region

#Region "  Interface Texture Rectangles  "
    Public Enum elInterfaceRectangle As Int32
        eButton_Normal = 0
        eButton_Disabled = 1
        eButton_Down = 2
        eSmall_Button_Normal = 3
        eSmall_Button_Disabled = 4
        eSmall_Button_Down = 5
        eUpArrow_Button_Normal = 6
        eUpArrow_Button_Disabled = 7
        eUpArrow_Button_Down = 8
        eLeftArrow_Button_Normal = 9
        eLeftArrow_Button_Disabled = 10
        eLeftArrow_Button_Down = 11
        eRightArrow_Button_Normal = 12
        eRightArrow_Button_Disabled = 13
        eRightArrow_Button_Down = 14
        eDownArrow_Button_Normal = 15
        eDownArrow_Button_Disabled = 16
        eDownArrow_Button_Down = 17

        eDirectionalArrow = 18
        eLightning = 19
        eSingleDude = 20
        eHappySadFace = 21

        eSphere = 22
        ePlanetOrbit = 23

        eXPRank_EmptyDot = 24
        eXPRank_SolidDot = 25
        eXPRank_Arrow = 26
        eXPRank_Bar = 27

        eWrench = 28
        eDemolish = 29
        eAlarm = 30
        eKeyButton = 31
        eLock = 32

        eCheck_Unchecked = 33
        eCheck_Disabled = 34
        eCheck_Checked = 35
        eCheck_Xed = 36
        eCheck_Blocked = 37

        eOption_Normal = 38
        eOption_Disabled = 39
        eOption_Marked = 40

        eLeftExpander = 41
        eRightExpander = 42
        eUpExpander = 43
        eDownExpander = 44

        eColonyStatsMinimizer = 45

        eEmailIcon = 46
        ePlanetIcon = 47

        eMinBar_0 = 48
        eMinBar_1 = 49
        eMinBar_2 = 50
        eMinBar_3 = 51
        eMinBar_4 = 52
        eMinBar_5 = 53
        eMinBar_6 = 54
        eMinBar_7 = 55
        eMinBar_8 = 56
        eMinBar_9 = 57
        eMinBar_10 = 58

        eWhiteBox = 59

        eAgentScrollLeftButton = 60
        eAgentScrollRightButton = 61

        eQuickbar_Help = 62
        eQuickbar_Email = 63
        eQuickbar_Battlegroup = 64
        eQuickbar_Trade = 65
        eQuickbar_Diplomacy = 66
        eQuickbar_ColonyStats = 67
        eQuickbar_Budget = 68
        eQuickbar_Mining = 69
        eQuickbar_Agent = 70
        eQuickbar_Formations = 71
        eQuickbar_ColonyResearch = 72
        eQuickbar_AvailResources = 73
        eQuickbar_ChatConfig = 74
        eQuickbar_Command = 75
        eQuickBar_Senate = 76
        eQuickBar_Transports = 77
        eQuickBar_Guild = 78
        eQuickbar_RouteTemplate = 79
        eQuickbar_EnvironmentRelations = 80

        eLastElement
    End Enum
    Public grc_UI() As Rectangle
    Public Sub CreateGlobalRectangleList()

        ReDim grc_UI(elInterfaceRectangle.eLastElement - 1)

        grc_UI(elInterfaceRectangle.eAgentScrollLeftButton) = New Rectangle(214, 107, 21, 50)
        grc_UI(elInterfaceRectangle.eAgentScrollRightButton) = New Rectangle(235, 107, 21, 50)
        grc_UI(elInterfaceRectangle.eAlarm) = New Rectangle(192, 96, 16, 16)
        grc_UI(elInterfaceRectangle.eButton_Disabled) = New Rectangle(0, 33, 120, 32)
        grc_UI(elInterfaceRectangle.eButton_Down) = New Rectangle(1, 64, 118, 32)
        grc_UI(elInterfaceRectangle.eButton_Normal) = New Rectangle(1, 0, 118, 32)
        grc_UI(elInterfaceRectangle.eCheck_Blocked) = New Rectangle(82, 235, 10, 10) 'New Rectangle(40, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Checked) = New Rectangle(62, 235, 10, 10) 'New Rectangle(20, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Disabled) = New Rectangle(52, 235, 10, 10) 'New Rectangle(10, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Unchecked) = New Rectangle(42, 235, 10, 10) 'New Rectangle(0, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eCheck_Xed) = New Rectangle(72, 235, 10, 10) 'New Rectangle(30, 215, 10, 10)
        grc_UI(elInterfaceRectangle.eColonyStatsMinimizer) = New Rectangle(0, 246, 159, 10)
        grc_UI(elInterfaceRectangle.eDemolish) = New Rectangle(192, 80, 16, 16)
        grc_UI(elInterfaceRectangle.eDirectionalArrow) = New Rectangle(100, 98, 16, 16)
        grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled) = New Rectangle(120, 96, 24, 24)
        grc_UI(elInterfaceRectangle.eDownArrow_Button_Down) = New Rectangle(168, 96, 24, 24)
        grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal) = New Rectangle(144, 96, 24, 24)
        grc_UI(elInterfaceRectangle.eDownExpander) = New Rectangle(27, 236, 10, 10)
        grc_UI(elInterfaceRectangle.eEmailIcon) = New Rectangle(161, 239, 32, 17)
        grc_UI(elInterfaceRectangle.eHappySadFace) = New Rectangle(110, 173, 16, 16)
        grc_UI(elInterfaceRectangle.eKeyButton) = New Rectangle(208, 80, 16, 16)
        grc_UI(elInterfaceRectangle.eLeftArrow_Button_Disabled) = New Rectangle(120, 48, 24, 24)
        grc_UI(elInterfaceRectangle.eLeftArrow_Button_Down) = New Rectangle(168, 48, 24, 24)
        grc_UI(elInterfaceRectangle.eLeftArrow_Button_Normal) = New Rectangle(144, 48, 24, 24)

        grc_UI(elInterfaceRectangle.eLeftExpander) = New Rectangle(0, 235, 8, 10)

        grc_UI(elInterfaceRectangle.eLightning) = New Rectangle(111, 144, 16, 16)
        grc_UI(elInterfaceRectangle.eLock) = New Rectangle(208, 64, 16, 16)

        For X As Int32 = 0 To 10
            grc_UI(elInterfaceRectangle.eMinBar_0 + X) = New Rectangle(193, 157 + (X * 9), 64, 9)
        Next X

        grc_UI(elInterfaceRectangle.eOption_Disabled) = New Rectangle(10, 225, 10, 10)
        grc_UI(elInterfaceRectangle.eOption_Marked) = New Rectangle(20, 225, 10, 10)
        grc_UI(elInterfaceRectangle.eOption_Normal) = New Rectangle(0, 225, 10, 10)
        grc_UI(elInterfaceRectangle.ePlanetIcon) = New Rectangle(161, 208, 32, 32)
        grc_UI(elInterfaceRectangle.ePlanetOrbit) = New Rectangle(144, 144, 16, 16)
        grc_UI(elInterfaceRectangle.eQuickbar_Agent) = New Rectangle(64, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_AvailResources) = New Rectangle(128, 192, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Battlegroup) = New Rectangle(0, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Budget) = New Rectangle(64, 96, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_ChatConfig) = New Rectangle(129, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_ColonyResearch) = New Rectangle(96, 192, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_ColonyStats) = New Rectangle(32, 160, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Command) = New Rectangle(224, 64, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Diplomacy) = New Rectangle(32, 128, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Email) = New Rectangle(0, 128, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Formations) = New Rectangle(64, 192, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Help) = New Rectangle(0, 96, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Mining) = New Rectangle(64, 128, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickBar_Senate) = New Rectangle(182, 120, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickbar_Trade) = New Rectangle(32, 96, 32, 32)
        grc_UI(elInterfaceRectangle.eQuickBar_Transports) = New Rectangle(0, 191, 32, 32)
        grc_UI(elInterfaceRectangle.eRightArrow_Button_Disabled) = New Rectangle(120, 72, 24, 24)
        grc_UI(elInterfaceRectangle.eRightArrow_Button_Down) = New Rectangle(168, 72, 24, 24)
        grc_UI(elInterfaceRectangle.eRightArrow_Button_Normal) = New Rectangle(144, 72, 24, 24)
        grc_UI(elInterfaceRectangle.eQuickBar_Guild) = New Rectangle(182, 120, 32, 32) 'Temp using Senate
        grc_UI(elInterfaceRectangle.eQuickbar_RouteTemplate) = New Rectangle(64, 192, 32, 32) 'Temp using Formations
        grc_UI(elInterfaceRectangle.eQuickbar_EnvironmentRelations) = New Rectangle(32, 160, 32, 32) 'Temp using Colony icon


        grc_UI(elInterfaceRectangle.eRightExpander) = New Rectangle(7, 235, 8, 10)

        grc_UI(elInterfaceRectangle.eSingleDude) = New Rectangle(111, 158, 15, 15)
        grc_UI(elInterfaceRectangle.eSmall_Button_Disabled) = New Rectangle(120, 0, 24, 24)
        grc_UI(elInterfaceRectangle.eSmall_Button_Down) = New Rectangle(168, 0, 24, 24)
        grc_UI(elInterfaceRectangle.eSmall_Button_Normal) = New Rectangle(144, 0, 24, 24)
        grc_UI(elInterfaceRectangle.eSphere) = New Rectangle(128, 144, 16, 16)
        grc_UI(elInterfaceRectangle.eUpArrow_Button_Disabled) = New Rectangle(120, 24, 24, 24)
        grc_UI(elInterfaceRectangle.eUpArrow_Button_Down) = New Rectangle(168, 24, 24, 24)
        grc_UI(elInterfaceRectangle.eUpArrow_Button_Normal) = New Rectangle(144, 24, 24, 24)

        grc_UI(elInterfaceRectangle.eUpExpander) = New Rectangle(15, 236, 10, 10)

        grc_UI(elInterfaceRectangle.eWhiteBox) = New Rectangle(192, 0, 62, 64)
        grc_UI(elInterfaceRectangle.eWrench) = New Rectangle(192, 64, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_Arrow) = New Rectangle(161, 192, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_Bar) = New Rectangle(177, 191, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_EmptyDot) = New Rectangle(161, 176, 16, 16)
        grc_UI(elInterfaceRectangle.eXPRank_SolidDot) = New Rectangle(177, 176, 16, 16)


    End Sub

#End Region

    Public Function Exists(ByVal sFilename As String) As Boolean
        If Trim(sFilename).Length > 0 Then
            On Error Resume Next
            sFilename = Dir$(sFilename, FileAttribute.Archive Or FileAttribute.Directory Or FileAttribute.Hidden Or FileAttribute.Normal Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
            Return Err.Number = 0 And sFilename.Length > 0
        Else
            Return False
        End If

    End Function

    Public Function GetStringFromBytes(ByRef yData() As Byte, ByVal lPos As Int32, ByVal lDataLen As Int32) As String
        Dim yTemp(lDataLen - 1) As Byte
        Array.Copy(yData, lPos, yTemp, 0, lDataLen)
        Dim lLen As Int32 = yTemp.Length
        For Y As Int32 = 0 To yTemp.Length - 1
            If yTemp(Y) = 0 Then
                lLen = Y
                Exit For
            End If
        Next Y
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)
    End Function

    Public Function UpdaterUpdate() As Boolean
        Dim sFilePath As String = AppDomain.CurrentDomain.BaseDirectory
        Dim sFileSrc As String
        Dim sFileDest As String
        Dim bResult As Boolean = False

        Dim lAttempts As Int32 = 0

        If sFilePath.EndsWith("\") = False Then sFilePath &= "\"

        sFileSrc = sFilePath & "UpdaterClient.ex_"
        sFileDest = sFilePath & "UpdaterClient.exe"

        If Debugger.IsAttached = False Then
            If Exists(sFilePath & "Meshes\Models.dat") = True Then
                Kill(sFilePath & "Meshes\Models.dat")
            End If
        End If

        If Exists(sFileSrc) = True Then
            'MSC 1/29/07 - check for file lock... (not pretty but it works)
            Dim bGood As Boolean = False
            While bGood = False AndAlso lAttempts < 10
                lAttempts += 1

                Dim oStream As IO.FileStream = Nothing
                Try
                    oStream = New IO.FileStream(sFileSrc, IO.FileMode.Append)
                    If oStream Is Nothing = False Then oStream.Dispose()
                    oStream = New IO.FileStream(sFileDest, IO.FileMode.Append)
                    If oStream Is Nothing = False Then oStream.Dispose()
                    bGood = True
                Catch ex As Exception
                    bGood = False
                Finally
                    If oStream Is Nothing = False Then oStream.Dispose()
                    oStream = Nothing
                End Try

                'If we're not valid yet then we wait...
                If bGood = False Then Threading.Thread.Sleep(500)
            End While

            If bGood = False Then
                MsgBox("An error occurred while attempting to update the updater client." & vbCrLf & "Please restart the updater and ensure that it is the only updater client running.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Unable to Update")
                End
            End If

            'Pulled this from the Updater Client...
            Dim oFSInfo As IO.FileInfo = New IO.FileInfo(sFileSrc)
            Dim dtSrc As Date = oFSInfo.CreationTime

            If oFSInfo.IsReadOnly = True Then

            End If
            oFSInfo = Nothing

            oFSInfo = New IO.FileInfo(sFileDest)
            If DateDiff("s", oFSInfo.CreationTime, dtSrc) <> 0 Then
                bResult = True
            End If
            oFSInfo = Nothing

            If bResult = True Then
                My.Computer.FileSystem.CopyFile(sFileSrc, sFileDest, True)

                Dim tMyDate As Date
                'now update the datetime stamp
                '2/17/2003 11:56:41 AM
                tMyDate = dtSrc

                oFSInfo = New IO.FileInfo(sFileDest)
                oFSInfo.CreationTime = tMyDate
                oFSInfo.LastAccessTime = tMyDate
                oFSInfo.LastWriteTime = tMyDate
                oFSInfo = Nothing
            End If
        End If

        'Always copy "epica.bm_" to "epica.bmp"
        sFileSrc = sFilePath & "BP.bm_"
        sFileDest = sFilePath & "BP.bmp"
        If Exists(sFileSrc) = True Then
            Try
                My.Computer.FileSystem.CopyFile(sFileSrc, sFileDest, True)
            Catch
            End Try
        End If

        Return bResult

    End Function

    Public Function GetDateTimeFromNumeric(ByVal lVal As Int32) As String
        Dim sValue As String = lVal.ToString

        If lVal = 0 Then Return ""

        Dim sMin As String = Right(sValue, 2)
        sValue = sValue.Substring(0, sValue.Length - 2)
        Dim sHr As String = Right(sValue, 2)
        sValue = sValue.Substring(0, sValue.Length - 2)
        Dim sDay As String = Right(sValue, 2)
        sValue = sValue.Substring(0, sValue.Length - 2)
        Dim sMonth As String = Right(sValue, 2)
        sValue = sValue.Substring(0, sValue.Length - 2)


        Dim lYear As Int32 = CInt(Val(sValue))
        lYear += 2000
        Dim sYear As String = lYear.ToString

        Return sMonth & "/" & sDay & "/" & sYear & " at " & sHr & ":" & sMin

    End Function

    Public Function GetDateFromNumber(ByVal lVal As Int32) As Date
        If lVal = 0 Then Return Date.MinValue

        Dim sVal As String = lVal.ToString

        'Work from right to left
        Dim lUB As Int32 = sVal.Length ' - 1

        'Ok, the bare minimum for this to work is 8
        If lUB < 8 Then Return Date.MinValue

        'Minute, last two values
        Dim sMin As String = sVal.Substring(lUB - 2)
        'Hour, two less from minute
        Dim sHr As String = sVal.Substring(lUB - 4, 2)
        'etc...
        Dim sDay As String = sVal.Substring(lUB - 6, 2)
        Dim sMon As String = sVal.Substring(lUB - 8, 2)

        Dim sYr As String = ""
        If lUB = 9 Then
            sYr = "0" & sVal.Substring(lUB - 9, 1)
        Else : sYr = sVal.Substring(lUB - 10, 2)
        End If


        Dim dtTemp As Date
        Try
            dtTemp = GetLocaleSpecificDT(sMon & "/" & sDay & "/20" & sYr & " " & sHr & ":" & sMin)
        Catch ex As Exception
            dtTemp = Now
        End Try
        Return dtTemp
    End Function

    Public Function GetLocaleSpecificDT(ByVal sUSBased As String) As Date
        Dim lIdx As Int32 = sUSBased.IndexOf("/"c)
        Dim lIdx2 As Int32 = sUSBased.IndexOf("/"c, lIdx + 1)
        Dim sMM As String = sUSBased.Substring(0, lIdx)
        Dim sDD As String = sUSBased.Substring(lIdx + 1, lIdx2 - lIdx - 1)
        Dim sYYYY As String = sUSBased.Substring(lIdx2 + 1, 4)
        Dim sTime As String = sUSBased.Substring(lIdx2 + 5).Trim

        Try
            Dim dtDate As Date = CDate("January 5, 2012")
            Dim sDate As String = dtDate.ToShortDateString.Replace(".", "/").Replace("-", "/")
            Dim lVal As Int32 = CInt(Val(sDate))
            If lVal = 1 Then
                'MM
                'now what?
                lIdx = sDate.IndexOf("/"c)
                If lIdx <> -1 Then sDate = sDate.Substring(lIdx + 1) Else sDate = sDate.Substring(2)
                lVal = CInt(Val(sDate))
                If lVal = 5 Then
                    'DD
                    'so, MM/DD/YYYY
                    Return CDate(sMM & "/" & sDD & "/" & sYYYY & " " & sTime)
                Else
                    'YYYY
                    'so, MM/YYYY/DD
                    Return CDate(sMM & "/" & sYYYY & "/" & sDD & " " & sTime)
                End If
            ElseIf lVal = 5 Then
                'DD
                lIdx = sDate.IndexOf("/"c)
                If lIdx <> -1 Then sDate = sDate.Substring(lIdx + 1) Else sDate = sDate.Substring(3)
                lVal = CInt(Val(sDate))
                If lVal = 1 Then
                    'MM
                    'so, DD/MM/YYYY
                    Return CDate(sDD & "/" & sMM & "/" & sYYYY & " " & sTime)
                Else
                    'YYYY
                    Return CDate(sDD & "/" & sYYYY & "/" & sMM & " " & sTime)
                End If
            Else
                'year
                lIdx = sDate.IndexOf("/"c)
                If lIdx <> -1 Then sDate = sDate.Substring(lIdx + 1) Else sDate = sDate.Substring(5)
                lVal = CInt(Val(sDate))
                If lVal = 1 Then
                    'MM
                    'so, YYYY/MM/DD
                    Return CDate(sYYYY & "/" & sMM & "/" & sDD & " " & sTime)
                Else
                    'DD
                    'so, YYYY/DD/MM
                    Return CDate(sYYYY & "/" & sDD & "/" & sMM & " " & sTime)
                End If
            End If
        Catch ex As Exception
            Try
                Return CDate(sUSBased)
            Catch
            End Try
        End Try
    End Function

    Public Function GetDateDisplayString(ByVal dtValue As Date, ByVal bDate As Boolean, ByVal bTime As Boolean) As String
        If bDate = True Then
            If bTime = True Then
                Return dtValue.ToShortDateString & " " & dtValue.ToShortTimeString
            Else
                Return dtValue.ToShortDateString
            End If
        Else
            If bTime = True Then
                Return dtValue.ToShortTimeString
            Else
                Return dtValue.ToString
            End If
        End If
    End Function

    Public Function GetDurationFromSeconds(ByVal lSeconds As Int32, ByVal bDHMS As Boolean) As String
        Dim lMinutes As Int32 = lSeconds \ 60 : lSeconds -= (lMinutes * 60)
        Dim lHours As Int32 = lMinutes \ 60 : lMinutes -= (lHours * 60)
        Dim lDays As Int32 = lHours \ 24 : lHours -= (lDays * 24)
        If bDHMS = True Then
            Dim sResult As String = ""
            If lDays <> 0 Then sResult &= lDays.ToString & "d "
            If lHours <> 0 Then sResult &= lHours.ToString & "h "
            If lMinutes <> 0 Then sResult &= lMinutes.ToString & "m "
            If lSeconds <> 0 OrElse sResult = "" Then sResult &= lSeconds.ToString & "s "
            Return sResult.Trim()
        Else
            Return lDays.ToString("00") & ":" & lHours.ToString("00") & ":" & lMinutes.ToString("00") & ":" & lSeconds.ToString("00")
        End If
    End Function

    Public Function GetDateAsNumber(ByVal dtDate As Date) As Int32
        Return CInt((dtDate.ToString("yyMMddHHmm")))
    End Function

    Public Function LaunchedFromUpdater() As Boolean
        Dim sCmd As String = Command()

        Dim oINI As New InitFile()
        Dim sOperator As String = oINI.GetString("CONNECTION", "OperatorIP", "")
        oINI = Nothing

        If sOperator = "12.228.7.156" Then Return True
        Return sCmd.Contains("EPICA_NORMAL_LAUNCH")
    End Function

    Public Function GetAlphaNumericSortable(ByVal lName As String, Optional ByVal lPadLen As Int32 = 3) As String
        Dim sChar As String = ""
        Dim oName As String = ""
        Dim bInNum As String = ""
        Dim lPadding As String = Space(lPadLen).Replace(" ", "0")
        If lPadding.Length = 0 Then lPadding = "0"
        For X As Int32 = 1 To lName.Length
            sChar = Mid(lName, X, 1)
            If Asc(sChar) >= 48 And Asc(sChar) <= 57 Then
                bInNum = bInNum & sChar
            Else
                If bInNum.Length > 0 Then
                    oName = oName & CInt(bInNum).ToString(lPadding) & sChar
                    bInNum = ""
                Else
                    oName = oName & sChar
                End If

            End If
        Next
        If bInNum.Length > 0 Then
            oName = oName & CInt(bInNum).ToString(lPadding)
        End If
        GetAlphaNumericSortable = oName

    End Function

    Public Function GetSortedMineralIdxArray(ByVal bIncludeUnknown As Boolean, Optional ByVal bIncludeAlloys As Boolean = True) As Int32()
        Static iLastArrayRefresh As Int32 = -1
        Static lMineralIdxArray_True_True() As Int32 = Nothing
        Static lMineralIdxArray_false_false() As Int32 = Nothing
        Static lMineralIdxArray_false_true() As Int32 = Nothing
        Static lMineralIdxArray_true_false() As Int32 = Nothing

        If glCurrentCycle - iLastArrayRefresh > 900 Then
            Debug.Print("GetSortedMineralIdxArray_new: " & ((glCurrentCycle - iLastArrayRefresh) / 30).ToString)
            iLastArrayRefresh = glCurrentCycle
            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To glMineralUB
                Dim lIdx As Int32 = -1
                If glMineralIdx(X) <> -1 AndAlso (bIncludeUnknown = True OrElse goMinerals(X).bDiscovered = True) AndAlso (bIncludeAlloys = True OrElse goMinerals(X).ObjectID <= 157) Then
                    Dim sMineral As String = GetAlphaNumericSortable(goMinerals(X).MineralName.ToUpper)
                    For Y As Int32 = 0 To lSortedUB
                        If GetAlphaNumericSortable(goMinerals(lSorted(Y)).MineralName.ToUpper) > sMineral Then
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

                End If
            Next X

            If bIncludeUnknown = True AndAlso bIncludeAlloys = True Then
                lMineralIdxArray_True_True = lSorted
            ElseIf bIncludeUnknown = False AndAlso bIncludeAlloys = False Then
                lMineralIdxArray_false_false = lSorted
            ElseIf bIncludeUnknown = False AndAlso bIncludeAlloys = True Then
                lMineralIdxArray_false_true = lSorted
            ElseIf bIncludeUnknown = True AndAlso bIncludeAlloys = False Then
                lMineralIdxArray_true_false = lSorted
            End If
            Return lSorted
        Else
            Debug.Print("GetSortedMineralIdxArray_old: " & ((glCurrentCycle - iLastArrayRefresh) / 30).ToString)
            If bIncludeUnknown = True AndAlso bIncludeAlloys = True Then
                Return lMineralIdxArray_True_True
            ElseIf bIncludeUnknown = False AndAlso bIncludeAlloys = False Then
                Return lMineralIdxArray_false_false
            ElseIf bIncludeUnknown = False AndAlso bIncludeAlloys = True Then
                Return lMineralIdxArray_false_true
            Else 'If bIncludeUnknown = True AndAlso bIncludeAlloys = False Then
                Return lMineralIdxArray_true_false
            End If
        End If
    End Function
    Public Enum GetSortedIndexArrayType As Int32
        eAgentName = 1
        ePlayerMissionGetListboxText = 2
        eMissionName = 3
        eFormationName = 4
        eCapturedAgentName = 5
        eGuildMemberListName = 6
        eGuildRelName = 7
        eGuildMemberRankName = 8
        eAgentStatus = 9
        eFleetName = 10
    End Enum

    Public Function GetSortedIndexArray(ByVal oItemArray() As Object, ByVal lItemIdxArray() As Int32, ByVal lVariableType As GetSortedIndexArrayType, Optional ByVal bDescending As Boolean = False) As Int32()

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        If oItemArray Is Nothing OrElse lItemIdxArray Is Nothing Then Return Nothing

        Dim lItemUB As Int32 = Math.Min(oItemArray.GetUpperBound(0), lItemIdxArray.GetUpperBound(0))

        For X As Int32 = 0 To lItemUB
            Dim lIdx As Int32 = -1

            If lItemIdxArray(X) <> -1 AndAlso oItemArray(X) Is Nothing = False Then

                Dim sNewValue As String = ""
                Select Case lVariableType
                    Case GetSortedIndexArrayType.eAgentName
                        sNewValue = CType(oItemArray(X), Agent).sAgentName
                    Case GetSortedIndexArrayType.eAgentStatus
                        sNewValue = Agent.GetStatusText(CType(oItemArray(X), Agent).lAgentStatus, CType(oItemArray(X), Agent).lTargetID, CType(oItemArray(X), Agent).iTargetTypeID, CType(oItemArray(X), Agent).InfiltrationType)
                    Case GetSortedIndexArrayType.eMissionName
                        sNewValue = CType(oItemArray(X), Mission).sMissionName
                    Case GetSortedIndexArrayType.ePlayerMissionGetListboxText
                        sNewValue = CType(oItemArray(X), PlayerMission).GetListBoxText
                    Case GetSortedIndexArrayType.eFormationName
                        sNewValue = CType(oItemArray(X), FormationDef).sName
                    Case GetSortedIndexArrayType.eFleetName
                        sNewValue = CType(oItemArray(X), UnitGroup).sName
                End Select

                For Y As Int32 = 0 To lSortedUB
                    Dim sLeftValue As String = "" 'CStr(CallByName(oItemArray(lSorted(Y)), sSortedProperty, CallType.Get))
                    Select Case lVariableType
                        Case GetSortedIndexArrayType.eAgentName
                            sLeftValue = CType(oItemArray(lSorted(Y)), Agent).sAgentName
                        Case GetSortedIndexArrayType.eAgentStatus
                            sLeftValue = Agent.GetStatusText(CType(oItemArray(lSorted(Y)), Agent).lAgentStatus, CType(oItemArray(lSorted(Y)), Agent).lTargetID, CType(oItemArray(lSorted(Y)), Agent).iTargetTypeID, CType(oItemArray(lSorted(Y)), Agent).InfiltrationType)
                        Case GetSortedIndexArrayType.eMissionName
                            sLeftValue = CType(oItemArray(lSorted(Y)), Mission).sMissionName
                        Case GetSortedIndexArrayType.ePlayerMissionGetListboxText
                            sLeftValue = CType(oItemArray(lSorted(Y)), PlayerMission).GetListBoxText
                        Case GetSortedIndexArrayType.eFormationName
                            sLeftValue = CType(oItemArray(lSorted(Y)), FormationDef).sName
                        Case GetSortedIndexArrayType.eFleetName
                            sLeftValue = CType(oItemArray(lSorted(Y)), UnitGroup).sName
                    End Select

                    If bDescending = False AndAlso sLeftValue.ToUpper > sNewValue.ToUpper Then
                        lIdx = Y
                        Exit For
                    ElseIf bDescending = True AndAlso sLeftValue.ToUpper < sNewValue.ToUpper Then
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
            End If
        Next X
        Return lSorted
    End Function
    Public Function GetSortedPlayerRelIdxArray() As Int32()
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        Dim sSortedVal() As String = Nothing

        If goCurrentPlayer Is Nothing Then Return Nothing

        For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
            Dim lIdx As Int32 = -1
            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(X)
            If oRel Is Nothing = False Then

                Dim lPID As Int32 = oRel.lThisPlayer
                If lPID = goCurrentPlayer.ObjectID Then lPID = oRel.lPlayerRegards
                Dim sName As String = GetCacheObjectValue(lPID, ObjectType.ePlayer).ToUpper

                For Y As Int32 = 0 To lSortedUB
                    If sSortedVal(Y).ToUpper > sName.ToUpper Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                ReDim Preserve sSortedVal(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                    sSortedVal(lSortedUB) = sName
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                        sSortedVal(Y) = sSortedVal(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                    sSortedVal(lIdx) = sName
                End If
            End If
        Next X
        Return lSorted
    End Function
    Public Function GetSortedIndexArrayNoIdxArray(ByVal oItemArray() As Object, ByVal lVariableType As GetSortedIndexArrayType, ByVal bDesc As Boolean) As Int32()

        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1
        If oItemArray Is Nothing Then Return Nothing

        Dim lItemUB As Int32 = oItemArray.GetUpperBound(0)

        For X As Int32 = 0 To lItemUB
            Dim lIdx As Int32 = -1

            If oItemArray(X) Is Nothing = False Then

                Dim sNewValue As String = ""
                Select Case lVariableType
                    Case GetSortedIndexArrayType.eAgentName
                        sNewValue = CType(oItemArray(X), Agent).sAgentName
                    Case GetSortedIndexArrayType.eMissionName
                        sNewValue = CType(oItemArray(X), Mission).sMissionName
                    Case GetSortedIndexArrayType.ePlayerMissionGetListboxText
                        sNewValue = CType(oItemArray(X), PlayerMission).GetListBoxText
                    Case GetSortedIndexArrayType.eFormationName
                        sNewValue = CType(oItemArray(X), FormationDef).sName
                    Case GetSortedIndexArrayType.eGuildMemberListName
                        With CType(oItemArray(X), GuildMember)
                            .sPlayerName = GetCacheObjectValue(.lMemberID, ObjectType.ePlayer)
                            sNewValue = .sPlayerName
                        End With
                    Case GetSortedIndexArrayType.eGuildMemberRankName
                        With CType(oItemArray(X), GuildMember)
                            sNewValue = goCurrentPlayer.oGuild.GetRankName(.lRankID)
                        End With
                    Case GetSortedIndexArrayType.eGuildRelName
                        'With CType(oItemArray(X), GuildRel)
                        '	.sName = GetCacheObjectValue(.lEntityID, .iEntityTypeID)
                        '	sNewValue = .sName
                        'End With
                End Select

                For Y As Int32 = 0 To lSortedUB
                    Dim sLeftValue As String = "" 'CStr(CallByName(oItemArray(lSorted(Y)), sSortedProperty, CallType.Get))
                    Select Case lVariableType
                        Case GetSortedIndexArrayType.eAgentName
                            sLeftValue = CType(oItemArray(lSorted(Y)), Agent).sAgentName
                        Case GetSortedIndexArrayType.eMissionName
                            sLeftValue = CType(oItemArray(lSorted(Y)), Mission).sMissionName
                        Case GetSortedIndexArrayType.ePlayerMissionGetListboxText
                            sLeftValue = CType(oItemArray(lSorted(Y)), PlayerMission).GetListBoxText
                        Case GetSortedIndexArrayType.eFormationName
                            sLeftValue = CType(oItemArray(lSorted(Y)), FormationDef).sName
                        Case GetSortedIndexArrayType.eGuildMemberListName
                            sLeftValue = CType(oItemArray(lSorted(Y)), GuildMember).sPlayerName
                        Case GetSortedIndexArrayType.eGuildRelName
                            'sLeftValue = CType(oItemArray(lSorted(Y)), GuildRel).sName
                        Case GetSortedIndexArrayType.eGuildMemberRankName
                            sLeftValue = goCurrentPlayer.oGuild.GetRankName(CType(oItemArray(lSorted(Y)), GuildMember).lRankID)
                    End Select

                    If bDesc = True Then
                        If sLeftValue.ToUpper < sNewValue.ToUpper Then
                            lIdx = Y
                            Exit For
                        End If
                    Else
                        If sLeftValue.ToUpper > sNewValue.ToUpper Then
                            lIdx = Y
                            Exit For
                        End If
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
            End If
        Next X
        Return lSorted
    End Function


    Public Function GetMonthName(ByVal lMonth As Int32) As String
        Select Case lMonth
            Case 1
                Return "January"
            Case 2
                Return "February"
            Case 3
                Return "March"
            Case 4
                Return "April"
            Case 5
                Return "May"
            Case 6
                Return "June"
            Case 7
                Return "July"
            Case 8
                Return "August"
            Case 9
                Return "September"
            Case 10
                Return "October"
            Case 11
                Return "November"
            Case 12
                Return "December"
            Case Else
                Return ""
        End Select
    End Function
    Public Function GetDayOfWeekName(ByVal lDayOfWeek As Int32) As String
        Select Case CType(lDayOfWeek, DayOfWeek)
            Case DayOfWeek.Friday
                Return "Friday"
            Case DayOfWeek.Monday
                Return "Monday"
            Case DayOfWeek.Saturday
                Return "Saturday"
            Case DayOfWeek.Sunday
                Return "Sunday"
            Case DayOfWeek.Thursday
                Return "Thursday"
            Case DayOfWeek.Tuesday
                Return "Tuesday"
            Case DayOfWeek.Wednesday
                Return "Wednesday"
        End Select
        Return ""
    End Function

#Region "  Bad Word Filter  "
    Private msWordChecks() As String = Nothing
    Private mlReplacementLen() As Int32 = Nothing
    Private mlWordCheckUB As Int32 = -1
    Private Sub InitializeWordChecks()
        AddWordCheck("fuck")
        AddWordCheck("bitch")
        'AddWordCheck("damnit")
        'AddWordCheck("damn")
        AddWordCheck("shit")
        AddWordCheck("cunt")
        AddWordCheck("dick")
        'AddWordCheck("ass")
        AddWordCheck("fagit")
        AddWordCheck("faggot")
        AddWordCheck("nigger")
        AddWordCheck("pussy")
        AddWordCheck("puss")
        AddWordCheck("cock")
        AddWordCheck("fag")
        AddWordCheck("whore")
        AddWordCheck("dike")
        AddWordCheck("kike")
        'AddWordCheck("jew")
    End Sub

    Private Sub AddWordCheck(ByVal sWord As String)
        ReDim Preserve msWordChecks(mlWordCheckUB + 2)
        ReDim Preserve mlReplacementLen(mlWordCheckUB + 2)

        Dim lNewIdx As Int32 = mlWordCheckUB + 1
        Dim lOtherIdx As Int32 = mlWordCheckUB + 2
        mlReplacementLen(lNewIdx) = sWord.Length
        mlReplacementLen(lOtherIdx) = sWord.Length
        msWordChecks(lNewIdx) = "\b"
        msWordChecks(lOtherIdx) = ""

        For X As Int32 = 0 To sWord.Length - 1
            Dim sChr As String = sWord(X).ToString
            Dim sCheckVal As String = "[" & sChr.ToLower & "|" & sChr.ToUpper
            If sChr.ToUpper = "S" Then sCheckVal &= "|$"
            If sChr.ToUpper = "I" Then sCheckVal &= "|!|1||"
            If sChr.ToUpper = "A" Then sCheckVal &= "|@"
            If sChr.ToUpper = "O" Then sCheckVal &= "|0"
            If sChr.ToUpper = "H" Then sCheckVal &= "|4"
            sCheckVal &= "]"
            msWordChecks(lNewIdx) &= sCheckVal & "[\W]?"
            msWordChecks(lOtherIdx) &= sCheckVal
        Next X

        msWordChecks(lOtherIdx) &= "\b"

        mlWordCheckUB += 2
    End Sub

    Public Function FilterBadWords(ByVal sInitial As String) As String
        If sInitial Is Nothing Then Return sInitial

        Try
            If muSettings.FilterBadWords = True Then
                If mlWordCheckUB = -1 Then InitializeWordChecks()

                For X As Int32 = 0 To msWordChecks.GetUpperBound(0)
                    sInitial = System.Text.RegularExpressions.Regex.Replace(sInitial, msWordChecks(X), StrDup(mlReplacementLen(X), "*"), System.Text.RegularExpressions.RegexOptions.IgnorePatternWhitespace)
                Next X
            End If
        Catch
        End Try

        Return sInitial
    End Function
#End Region

    Private mlCrashLogID As Int32 = 0
    Public Sub LogCrashEvent(ByVal ex As Exception, ByVal bDisplayMsgBox As Boolean, ByVal bDisplayInNotificationWindow As Boolean, ByVal sCallFrom As String, ByVal bForcefulShutdown As Boolean, ByRef oMsgSys As MsgSystem)
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"
        Dim sFile As String = "CrashLog_" & Now.ToString("yyyyMMddhhmmss") '& ".txt"
        If mlCrashLogID <> 0 Then sFile &= mlCrashLogID.ToString
        sFile &= ".txt"
        mlCrashLogID += 1

        Dim moSB As New System.Text.StringBuilder

        Try
            Dim oFS As New IO.FileStream(sPath & sFile, IO.FileMode.Create)
            Dim oWrite As New IO.StreamWriter(oFS)
            oWrite.AutoFlush = True

            moSB.AppendLine(Now.ToString("MM ddhhmmss") & (Now.Millisecond \ 100).ToString)
            moSB.AppendLine("Exception encountered in " & sCallFrom)
            moSB.AppendLine("Reported source: " & ex.Source)
            moSB.AppendLine(ex.Message)
            moSB.AppendLine("Stack Trace: ")
            moSB.AppendLine(ex.StackTrace)

            oWrite.Write(moSB.ToString)

            oWrite.WriteLine()
            oWrite.WriteLine()

            Dim oInner As Exception = ex.InnerException
            While oInner Is Nothing = False
                oWrite.WriteLine("Entering Inner")
                oWrite.WriteLine("Reported source: " & oInner.Source)
                oWrite.WriteLine(oInner.Message)
                oWrite.WriteLine("Stack Trace: ")
                oWrite.WriteLine(oInner.StackTrace)
                oWrite.WriteLine()
                oWrite.WriteLine()

                oInner = oInner.InnerException
            End While

            oWrite.Close()
            oWrite.Dispose()
            oFS.Close()
            oFS.Dispose()
        Catch
        End Try

        If oMsgSys Is Nothing = False Then
            Dim sVal As String = moSB.ToString
            If sVal.Length > 25000 Then sVal = sVal.Substring(0, 25000)
            Dim yMsg(5 + sVal.Length) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eExceptionReport).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(sVal.Length).CopyTo(yMsg, 2)
            System.Text.ASCIIEncoding.ASCII.GetBytes(sVal).CopyTo(yMsg, 6)
            Try
                oMsgSys.SendToPrimary(yMsg)
            Catch
            End Try
        End If

        If bDisplayMsgBox = True Then
            Dim sForceful As String = ""
            If bForcefulShutdown = True Then
                sForceful = vbCrLf & "Beyond Protocol has encountered an error and must shut down." & vbCrLf & "We apologize for the inconvenience this may have caused."
            End If
            MsgBox("An error log of a crash has been generated in the directory:" & vbCrLf & _
                   sPath & vbCrLf & vbCrLf & "The file is called:" & sFile & vbCrLf & vbCrLf & _
                   "Please send that file to support@darkskyentertainment.com." & sForceful, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Exception Detected")
        End If
        If bDisplayInNotificationWindow = True Then
            If goUILib Is Nothing = False Then
                goUILib.AddNotification("An error log was generated in the Beyond Protocol folder.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                goUILib.AddNotification("Please send that log in an email to support@darkskyentertainment.com.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If
    End Sub

    Public Sub HACK_SetCacheEntries()
        'And since we're at it...
        glEntityDefUB = 1
        ReDim glEntityDefIdx(glEntityDefUB)
        ReDim goEntityDefs(glEntityDefUB)
        goEntityDefs(0) = New EntityDef()
        With goEntityDefs(0)
            .DefName = "Enlisted x10"
            .ObjectID = 0
            .ObjTypeID = ObjectType.eEnlisted
            .RequiredProductionTypeID = ProductionType.eEnlisted
            .ProductionCost = New ProductionCost()
            .yChassisType = ChassisType.eGroundBased Or ChassisType.eSpaceBased
        End With
        With goEntityDefs(0).ProductionCost
            .ColonistCost = 20
            .CreditCost = 200
            .EnlistedCost = 0
            .ItemCostUB = -1
            ReDim .ItemCosts(-1)
            .PointsRequired = 100000
        End With
        glEntityDefIdx(0) = 0

        goEntityDefs(1) = New EntityDef()
        With goEntityDefs(1)
            .DefName = "Officers x10"
            .ObjectID = 0
            .ObjTypeID = ObjectType.eOfficers
            .RequiredProductionTypeID = ProductionType.eOfficers
            .ProductionCost = New ProductionCost()
            .yChassisType = ChassisType.eGroundBased Or ChassisType.eSpaceBased
        End With
        With goEntityDefs(1).ProductionCost
            .ColonistCost = 0
            .CreditCost = 200
            .EnlistedCost = 20
            .ItemCostUB = -1
            ReDim .ItemCosts(-1)
            .PointsRequired = 300000
        End With
        glEntityDefIdx(1) = 0
    End Sub

    Public Sub UpdateCP()
        frmEnvirDisplay.lMyFacilities = 0
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) <> -1 Then
                    Dim oTmpEntity As BaseEntity = oEnvir.oEntity(X)
                    If oTmpEntity Is Nothing OrElse oTmpEntity.oMesh Is Nothing Then Continue For
                    With oTmpEntity
                        If .ObjTypeID = ObjectType.eFacility AndAlso .OwnerID = glPlayerID Then
                            If .yProductionType = ProductionType.eSpaceStationSpecial Then frmEnvirDisplay.lMyFacilities += 50 Else frmEnvirDisplay.lMyFacilities += 1
                        End If
                    End With
                End If
            Next
        End If
    End Sub

    Public Function isAdmin() As Boolean
        Select Case glActualPlayerID
            Case 1              ' enochdagor
                Return True
            Case 2              ' Csaj
                Return True
            Case 6              ' Aurelius
                Return True
            Case 7              ' normacenva
                Return True
            Case 25             ' Aurelium Pirates
                Return False
            Case 131            ' Rakura
                Return False
            Case 2067           ' WillPotter
                Return False
            Case 2076           ' matwurks
                Return False
            Case 3253           ' StrawberryBunny
                Return False
            Case 3510, 21296    ' EpicFail
                Return True
            Case 20611, 28005   ' Nero
                Return True
            Case Else           ' you
                Return False
        End Select
    End Function

    Public Function GetRomanNumeralSortStr(ByVal sVal As String) As String
        Select Case sVal.ToUpper
            Case "I", "PRIME", "581B"
                Return "001"
            Case "II", "SECONDUS", "SECUNDUS", "581C"
                Return "002"
            Case "III", "RTEJADA", "581D"
                Return "003"
            Case "IV", "VEXUS", "581E"
                Return "004"
            Case "V", "NARN", "581F"
                Return "005"
            Case "VI", "BEAN", "581G"
                Return "006"
            Case "VII"
                Return "007"
            Case "VIII"
                Return "008"
            Case "IX"
                Return "009"
            Case "X"
                Return "010"
            Case "XI"
                Return "011"
            Case "XII"
                Return "012"
            Case "XIII"
                Return "013"
            Case "XIV"
                Return "014"
            Case "XV"
                Return "015"
            Case "XVI"
                Return "016"
            Case "XVII"
                Return "017"
            Case "XVIII"
                Return "018"
            Case "XIX"
                Return "019"
            Case "XX"
                Return "020"
            Case "XXI"
                Return "021"
            Case "XXII"
                Return "022"
            Case "XXIII"
                Return "023"
            Case "XXIV"
                Return "024"
            Case "XXV"
                Return "025"
            Case "XXVI"
                Return "026"
            Case "XXVII"
                Return "027"
            Case "XXVIII"
                Return "028"
            Case "XXIX"
                Return "029"
            Case "XXX"
                Return "030"
            Case "XXXI"
                Return "031"
            Case Else
                Return sVal
        End Select
    End Function

    Public Function GetPlanetNameValue(ByVal sName As String) As Int32
        Try
            Dim lTmpVal As Int32 = sName.LastIndexOf(" "c)
            Dim sTemp As String
            If lTmpVal > -1 Then
                sTemp = sName.Substring(lTmpVal).Trim
            Else
                sTemp = sName
            End If
            Return CInt(Val(GetRomanNumeralSortStr(sTemp)))
        Catch
        End Try
    End Function

    Public Function GotoEnvironmentWrapper(ByVal ObjectID As Int32, ByVal ObjTypeID As Short) As Boolean
        Try
            If ObjTypeID = ObjectType.eSolarSystem Then
                For x As Int32 = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(x).ObjectID = ObjectID Then
                        frmMain.ForceChangeEnvironment(ObjectID, ObjTypeID)
                        Return True
                    End If
                Next
            Else
                For x As Int32 = 0 To goGalaxy.mlSystemUB
                    For y As Int32 = 0 To goGalaxy.moSystems(x).PlanetUB
                        If goGalaxy.moSystems(x).moPlanets(y).ObjectID = ObjectID Then
                            frmMain.ForceChangeEnvironment(ObjectID, ObjTypeID)
                            Return True
                        End If
                    Next
                Next
            End If
        Catch
        End Try
    End Function
End Module