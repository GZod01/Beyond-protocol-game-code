Option Strict On

'This is for the exclamation point system
Public Class TutorialAlert
    Public Enum eTriggerType As Int32
        EndsTutorialInAurelium = 1
        LowPowerEvent = 2
        LowCreditsEvent = 3
        LowResourcesEvent = 4
        FirstMiningFacilityBuilt = 5
        FirstProducerBuilt = 6
        TradepostBuilt = 7
        LowEnlisted = 8
        OpenBudgetWindow = 9
        OpenEmailWindow = 10
        OpenBattlegroupWindow = 11
        OpenDiplomacyWindow = 12
        MineralCacheHover = 13
        'OpenMiningWindow = 14
        OpenColonyStatsWindow = 15
        OpenResearchWindow = 16
        DiscoverNewMineral = 17
        OpenAgentMainWindow = 18
        GetFirstAgent = 19
        OpenCreateMissionWindow = 20
        'OpenAgentWindow = 21
        OpenFormationsWindow = 22
        'OpenColonyProdResWindow = 23
        OpenAvailableResourcesWindow = 24
        OpenCommandManagementWindow = 25
        'OpenSenateWindow = 26
        OpenChatConfigWindow = 27
        ClickOnOrbitalBombardment = 28
        BuildSpaceCapableUnit = 29
        'OpenContentsWindow = 30
        OpenOrdersWindow = 31
        OpenRoutesWindow = 32
        SelectDmgdUnit = 33
        'SelectDmgdFacility = 34
        GuildSearchOpen = 35
        'LowCapacity = 36
        LowOfficers = 37

        eLastTriggerType        'for looping
    End Enum

    Public Enum eTutorialTriggerItem As Int32
        WhatIsAurelium = 1
        WhatToDoNow = 2
        MeetingPlayers = 3
        GoingLive = 4
        GettingMoreHelp = 5

        AllAboutPower = 10
        MakingCredits = 11
        AvailableResources = 12
        LowEnlisted = 13
        LowOfficers = 14

        MineralBidding = 20
        BuildAUnit = 21
        TheTradepost = 22
        MineralCacheAndMining = 23

        BudgetWindow = 25
        BudgetWindowSpecifics = 26
        DeathBudget = 27
        EmailWindow = 28
        BattlegroupWindow = 29
        DiplomacyWindow = 30
        ColonyDetailsWindow = 31
        ColonyIntelligence = 32
        ColonyMorale = 33
        ColonyEfficiency = 34
        SystemToSystemMove = 35
        ResearchWindow = 36
        UndiscoveredMineral = 37
        AgentMainWindow = 38
        AgentBasics = 39
        AgentMissionSetup = 40
        FormationWindow = 41
        CommandWindow = 42
        ChatChannelMgmt = 43
        OrbitalBombardment = 44
        UnitToSpace = 45
        OrdersWindow = 46
        RoutesWindow = 47
        RepairAUnit = 48
        FindAGuild = 49

        GettingTechSupport = 50
    End Enum



    Public Shared Sub InitializeAlertSystem()
        ReDim gyTriggerFired(eTriggerType.eLastTriggerType - 1)
        For X As Int32 = 0 To gyTriggerFired.GetUpperBound(0)
            gyTriggerFired(X) = 0
        Next X


        If goCurrentPlayer Is Nothing = False Then
            Dim sDATFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sDATFile.EndsWith("\") = False Then sDATFile &= "\"
            sDATFile &= goCurrentPlayer.PlayerName & ".tut"

            Dim oINI As New InitFile(sDATFile)

            Dim sTriggers As String = oINI.GetString("Tutorial", "ExcPntTriggers", StrDup(eTriggerType.eLastTriggerType, "0"c))
            For X As Int32 = 0 To Math.Min(gyTriggerFired.GetUpperBound(0), sTriggers.Length - 1)
                gyTriggerFired(X) = CByte(Val(sTriggers.Substring(X, 1)))
            Next X

            oINI = Nothing

            'Now, check our triggers, any 1's, show them
            For X As Int32 = 0 To gyTriggerFired.GetUpperBound(0)
                If gyTriggerFired(X) = 1 Then
                    SpawnExclamationPoint(X)
                End If
            Next X
        End If
    End Sub

    Public Shared Sub CheckTrigger(ByVal lTrigger As eTriggerType)
        If goCurrentPlayer Is Nothing Then Return
        If goCurrentPlayer.yPlayerPhase = 0 Then Return
        If gyTriggerFired Is Nothing Then InitializeAlertSystem()
        'If muSettings.gbDisableExclamations = True Then Exit Sub
        If lTrigger > -1 AndAlso lTrigger < eTriggerType.eLastTriggerType Then
            If gyTriggerFired(lTrigger) = 0 Then
                'Ok, spawn it
                SpawnExclamationPoint(lTrigger)

                'and set the trigger to 1
                gyTriggerFired(lTrigger) = 1    'indicating that it has been shown but not clicked

                SaveData()
            End If
        End If
    End Sub

    Public Shared Sub ExclamationPointClicked(ByVal lTrigger As Int32)
        If gyTriggerFired Is Nothing Then InitializeAlertSystem()
        If lTrigger > -1 AndAlso lTrigger < eTriggerType.eLastTriggerType Then
            gyTriggerFired(lTrigger) = 2    'set to 2 for being shown and clicked
        End If
        SaveData()
    End Sub

    Private Shared Sub SaveData()
        If goCurrentPlayer Is Nothing = False Then
            Dim sDATFile As String = AppDomain.CurrentDomain.BaseDirectory
            If sDATFile.EndsWith("\") = False Then sDATFile &= "\"
            sDATFile &= goCurrentPlayer.PlayerName & ".tut"

            Dim oINI As New InitFile(sDATFile)

            Dim sTriggers As String = ""
            '
            For X As Int32 = 0 To gyTriggerFired.GetUpperBound(0)
                'gyTriggerFired(X) = CByte(Val(sTriggers.Substring(X, 1)))
                sTriggers &= gyTriggerFired(X).ToString
            Next
            'oINI.WriteString("Tutorial", "ExcPntTriggers", StrDup(eTriggerType.eLastTriggerType, "0"c))
            oINI.WriteString("Tutorial", "ExcPntTriggers", sTriggers)


            oINI = Nothing
        End If
    End Sub

    Private Shared Sub SpawnExclamationPoint(ByVal lAlertTrigger As Int32)

        'Ok, have to translate our trigger
        Dim lTrigger As Int32 = 0
        Select Case CType(lAlertTrigger, eTriggerType)
            Case eTriggerType.BuildSpaceCapableUnit
                lTrigger = eTutorialTriggerItem.UnitToSpace
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.ClickOnOrbitalBombardment
                lTrigger = eTutorialTriggerItem.OrbitalBombardment
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.DiscoverNewMineral
                lTrigger = eTutorialTriggerItem.UndiscoveredMineral
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.EndsTutorialInAurelium
                lTrigger = eTutorialTriggerItem.WhatIsAurelium
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.WhatToDoNow
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.MeetingPlayers
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.GoingLive
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.GettingMoreHelp
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.FirstMiningFacilityBuilt
                lTrigger = eTutorialTriggerItem.MineralBidding
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.FirstProducerBuilt
                lTrigger = eTutorialTriggerItem.BuildAUnit
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.GetFirstAgent
                lTrigger = eTutorialTriggerItem.AgentBasics
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.GuildSearchOpen
                lTrigger = eTutorialTriggerItem.FindAGuild
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.LowCreditsEvent
                lTrigger = eTutorialTriggerItem.MakingCredits
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.LowEnlisted
                lTrigger = eTutorialTriggerItem.LowEnlisted
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.LowOfficers
                lTrigger = eTutorialTriggerItem.LowOfficers
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.LowPowerEvent
                lTrigger = eTutorialTriggerItem.AllAboutPower
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.LowResourcesEvent
                lTrigger = eTutorialTriggerItem.AvailableResources
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.MineralCacheHover
                lTrigger = eTutorialTriggerItem.MineralCacheAndMining
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenAgentMainWindow
                lTrigger = eTutorialTriggerItem.AgentMainWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenAvailableResourcesWindow
                lTrigger = eTutorialTriggerItem.AvailableResources
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenBattlegroupWindow
                lTrigger = eTutorialTriggerItem.BattlegroupWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenBudgetWindow
                lTrigger = eTutorialTriggerItem.BudgetWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.BudgetWindowSpecifics
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.DeathBudget
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenChatConfigWindow
                lTrigger = eTutorialTriggerItem.ChatChannelMgmt
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenColonyStatsWindow
                lTrigger = eTutorialTriggerItem.ColonyDetailsWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.ColonyIntelligence
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.ColonyMorale
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing

                lTrigger = eTutorialTriggerItem.ColonyEfficiency
                oFrm = New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenCommandManagementWindow
                lTrigger = eTutorialTriggerItem.CommandWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenCreateMissionWindow
                lTrigger = eTutorialTriggerItem.AgentMissionSetup
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenDiplomacyWindow
                lTrigger = eTutorialTriggerItem.DiplomacyWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenEmailWindow
                lTrigger = eTutorialTriggerItem.EmailWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenFormationsWindow
                lTrigger = eTutorialTriggerItem.FormationWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenOrdersWindow
                lTrigger = eTutorialTriggerItem.OrdersWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenResearchWindow
                lTrigger = eTutorialTriggerItem.ResearchWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.OpenRoutesWindow
                lTrigger = eTutorialTriggerItem.RoutesWindow
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.SelectDmgdUnit
                lTrigger = eTutorialTriggerItem.RepairAUnit
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
            Case eTriggerType.TradepostBuilt
                lTrigger = eTutorialTriggerItem.TheTradepost
                Dim oFrm As New frmExclPoint(goUILib, lTrigger, lAlertTrigger)
                oFrm.Visible = True
                oFrm = Nothing
        End Select



    End Sub
End Class
