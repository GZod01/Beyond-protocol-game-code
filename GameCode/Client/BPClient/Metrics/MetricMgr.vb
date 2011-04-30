Option Strict On

Namespace BPMetrics
    Public Enum eInterface As Int32
        eViewInterface = 1              'viewed an interface
        eCloseInterface = 2             'closed an interface
        eFullScreenModeSwitch = 3       'switched from fs to window or window to fs
        eGotoOrbitClick = 4             'the goto orbit button was clicked on the multi-select or advanced display windows
        eUIElementClick = 5             'a ui element such as button was clicked
        eUIElementHover = 6             'a ui element was hovered over
        eScrollCamera = 7               'the edge of the screen or arrow key causing the camera to scroll
        eRotateCamera = 8               'normal rotation
        eZoomCameraOut = 9              'zoom the camera out
        eZoomCameraIn = 10              'zoom the camera in
        eChangeCurrentEnvirView = 11    'changed the view (galaxy view, planet view, etc...)
        eChatMessageSent = 12           'chat message was sent - per type
        eHelpWindowTopicExpand = 13     'a topic in the help window is expanded
        eHelpWindowItemClick = 14       'an item in the help window is clicked to view
        eDraftEmail = 15                'player saved a draft of an email
        eEmailDeleted = 16              'email delete click
        eBuyOrderCategoryClick = 17     'player clicked on a buy order category
        eSellOrderCategoryClick = 18    'player clicked on a sell order category
        'eColonyWindowExpandClick = 19   'the button to expand/collapse the colony window was clicked
        eBudgetWindowSortClick = 20     'an option to sort the budget window was clicked
        'eOrbitalBombardmentSetting = 21 'a setting for the orbital bombardment was used
        eEntityChangeName = 22          'changing the name of an entity or colony
        'eCargoContentsShown = 23        'either via switch view or opening the contents window
        'eHangarContentsShown = 24       'either via switch view or opening the contents window
        eMouseMoveDelta = 25            'mouse move (start pt to finish pt)
        eMouseMoveTotalDelta = 26       'total distance a mouse moved (in pixels) - this number would be different from eMouseMoveDelta if the mouse move was "swirly" compared to straight line
    End Enum
    Public Enum eOverallGameplay As Int32
        eFinishedTutorial1 = 1000
        eFinishedTutorial2 = 1001
        eIdleTime = 1002                    'a duration of time where the player is not doing anything
    End Enum
    Public Enum eKeyPress As Int32
        eAltTab = 2000
        eCtrlEscape = 2001
        eArrowKey = 2002
        eTilda = 2003
        eTabKey = 2004
        eHomeKey = 2005
        eBKeyToOpenBuild = 2006
        eCKeyToOpenContents = 2007
        eAltKey = 2008
        eTKeyToTrack = 2009
        eZKeyToZoom = 2010
        eF1Key = 2011
        eF2Key = 2012
        eF3Key = 2013
        eF4Key = 2014
        eF5Key = 2015
        eF6Key = 2016
        eF7Key = 2017
        eF8Key = 2018
        eF9Key = 2019
        eF10Key = 2020
        eF11Key = 2021
        eF12Key = 2022
        eEscapeKey = 2023
        eSelectNextEngineer = 2024
        eSelectNextFacility = 2025
        eSelectNextIdleEngineer = 2026
        eToggleMineralCaches = 2027
        eSelectNextEntity = 2028
        eSelectNextUnpoweredFacility = 2029
        eCtrlQ = 2030
        eToggleRallyPoint = 2031
        eSelectSimilar = 2032
        eSelectNextUnit = 2033
        eCtrlGroupAssign = 2034
        eRecallCtrlGroup = 2035
        eCtrlShiftCtrlGroup = 2036
        eCaptureScreenshot = 2037
        eSelectNextUnpoweredResidence = 2038
        eOKeyToOpenOrders = 2039
    End Enum
    Public Enum eOrders As Int32
        eGroupMovement = 3001
        eMoveFormation = 3002
        eDockRequest = 3003
        eChangeEnvironmentMovement = 3004
        eSetPrimaryTarget = 3005
    End Enum
    Public Enum eWindow As Int32
        '--- agent folder ---
        eAgentView = 5001
        eAgentMain = 5002
        eCapturedAgent = 5003
        eIntelReport = 5004
        eCreateMission = 5005
        eSkillDetailWindow = 5006
        eViewMission = 5007
        '--- build windows ---
        eAvailableResources = 5010
        eBuildWindow = 5011
        eCancelBuild = 5012
        eRepairWindow = 5013
        eResearchMain = 5014
        eSelectFac = 5015
        '--- chat ---
        eChannelConfig = 5020
        eChannels = 5021
        eChatWindow = 5022
        eChatTabProps = 5023
        eColorSelect = 5024
        '--- colony ---
        eBudgetWindow = 5030
        eColonyResearch = 5031
        eColonyStats = 5032
        eColonyStatsSmall = 5033
        '--- diplomacy ---
        eAliasingWindow = 5040
        eConfirmWarDec = 5041
        eCustomTitle = 5042
        eDiplomacy = 5043
        eFaction = 5044
        eMyAliasesWindow = 5045
        '--- Email ---
        eAddressBook = 5050
        eEmailMain = 5051
        eEmailSetup = 5052
        eNewEmail = 5053
        '--- Fleet Management ---
        eBombard = 5060
        eFleet = 5061
        eFleetOrders = 5062
        eFormations = 5063
        eRouteConfig = 5064
        '--- Guild ---
        eAttachments = 5070
        eGuildMain_Assets = 5071
        eGuildMain_Events = 5072
        eGuildMain_Internal = 5073
        eGuildMain_Main = 5074
        eGuildMain_Membership = 5075
        eGuildMain_Politics = 5076
        eGuildMain_Structure = 5077
        eGuildSearch = 5078
        eGuildSetup = 5079
        eInviteFromGuild = 5080
        eRaceConfig = 5081
        eSketchPad = 5082
        eTournament = 5083
        eGuildMain_Finances = 5084
        '--- HUD ---
        eAdvanceDisplay = 5090
        eBehavior = 5091
        eCommand = 5092
        eContents = 5093
        eEnvirDisplay = 5094
        eMinimapZoom = 5095
        eMultiDisplay = 5096
        eNotification = 5097
        eNotificationHistory = 5098
        eProdStatus = 5099
        eQuickBar = 5100
        eSingleSelect = 5101
        '--- Mining ---
        eMiningBid = 5110
        eMining = 5111
        '--- Senate ---
        eSenate = 5120
        '--- Tech Builders ---
        eAlloyBuilder = 5130
        eArmorBuilder = 5131
        eEngineBuilder = 5132
        eHullBuilder = 5133
        eMaterialResearch = 5134
        eMineralDetail = 5135
        ePrototypeBuilder = 5136
        eRadarBuilder = 5137
        eShieldBuilder = 5138
        eSpecialTech = 5139
        eWeaponAppearance = 5140
        eBombBuilder = 5141
        eMissileBuilder = 5142
        eProjectileBuilder = 5143
        ePulseBuilder = 5144
        eSolidBeamBuilder = 5145
        '--- Trade ---
        eCreateBuyOrder = 5150
        eCreateBuyOrderReqs = 5151
        eDeliverBuyOrder = 5152
        eDirectTrade = 5153
        eTradeContents = 5154
        eTradeHistory = 5155
        eTradeInProgress = 5156
        eNonOwnerItemDetails = 5157
        eTradeMain = 5158
        '--- Tutorial ---
        eTutorialStep = 5170
        eTutorialTOC = 5171
        eTutorialOld = 5172
        '--- UIScreens folder ---
        eCacheSelect = 5180
        eDeath = 5181
        eInputBox = 5182
        eLoginScreen = 5183
        eMsgBox = 5184
        eNotePad = 5185
        eObserve = 5186
        eOptions = 5187
        ePlanetDetails = 5188
        ePlayerSetup = 5189
        eProdCost = 5190
        eTransfer = 5191
        '--- ARENA ---
        eArenaMain = 5200
        eArenaConfig = 5201
    End Enum

    Public Class MetricMgr
        Private Class Activity
            'ddhhmmssX where X is a fraction of a second
            Public ActivityID As Int32
            Public Value1 As Int32
            Public Value2 As Int32
            Public MiniTimeStamp As Int32

            Public Sub New(ByVal lActivityID As Int32, ByVal lVal1 As Int32, ByVal lVal2 As Int32)
                MiniTimeStamp = CInt(Now.ToString("ddhhmmss") & (Now.Millisecond \ 100).ToString)
                ActivityID = lActivityID
                Value1 = lVal1
                Value2 = lVal2
            End Sub
        End Class

        Private Shared mcolActivities As Collection

        Public Shared lFramesRendered As Int32
        Public Shared dtSessionStarted As DateTime

        Public Shared Sub AddActivity(ByVal lActivityCode As Int32, ByVal lVal1 As Int32, ByVal lVal2 As Int32)
            If mcolActivities Is Nothing Then mcolActivities = New Collection

            Dim oActivity As New Activity(lActivityCode, lVal1, lVal2)
            mcolActivities.Add(oActivity)
        End Sub

        Public Shared Function GetActivityReport(ByVal lFPS As Int32) As Byte()
            Dim colTemp As Collection = mcolActivities
            mcolActivities = New Collection()

            If colTemp.Count = 0 Then Return Nothing

            Dim yMsg(21 + (colTemp.Count * 16)) As Byte
            Dim lPos As Int32 = 0

            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerActivityReport).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(colTemp.Count).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(lFPS).CopyTo(yMsg, lPos) : lPos += 4
            Dim oProc As Process = Process.GetCurrentProcess()
            System.BitConverter.GetBytes(CInt(oProc.TotalProcessorTime.TotalSeconds)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CInt(oProc.VirtualMemorySize64 \ 1000000)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CInt(oProc.WorkingSet64 \ 1000000)).CopyTo(yMsg, lPos) : lPos += 4

            For Each oAct As Activity In colTemp
                With oAct
                    System.BitConverter.GetBytes(.ActivityID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.Value1).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.Value2).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.MiniTimeStamp).CopyTo(yMsg, lPos) : lPos += 4
                End With
            Next oAct

            Return yMsg
        End Function

    End Class
End Namespace