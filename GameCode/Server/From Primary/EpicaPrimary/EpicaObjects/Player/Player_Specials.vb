Option Strict On

Public Enum PlayerSpecialAttributeID As Short
    eSpecialSpecialAttribute = 0            'hehehehe
    'eAerialUpkeepReduction = 0
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

Public Class Player_Specials
    Public Enum SuperSpecialID As Int32
        eNoSuperSpecials = 0
        'eAdjustableColonyGrowth = 1        'no longer used here
        eShieldsAllowedOnPlanets = 2
        eWormholesTraversable = 4
        eSpaceStationsCloserToPlanet = 8
        eOrbitalMiningPlatform = 16
        eSpaceElevator = 32
        eStargates = 64
        eLightningShields = 128
        'eControlledMorale = 256            'no longer used here
        eFakePlanetGravityWell = 512
        eMassDriver = 1024
        eOverloadingCapacitor = 2048
        ePerfectCrystals = 4096
        eFoldSpace = 8192
        eHyperSpace = 16384
        eWarpDrive = 32768
        eEnhancedSpaceRadar = 65536
        eDoctrinesOfLeadership = 131072
    End Enum
    Public Enum elNo2XValues As Int32
        eNoBurn2XImpact = 1
        eNoChem2XBurn = 2
        eNoImpact2XBurn = 4
        eNoMag2XChem = 8
        eNoPierce2XBeam = 16
    End Enum
    'NOTE: Anywhere I say "stored as a percentage" when using a BYTE, that indicates that the value
    '  is multiplied by 100. For example, a value of 2 would be .02 because the value is multiplied by 100: .02 * 100 = 2.
    '  To get the actual result, simply divide by 100.

    Public oPlayer As Player        'pointer    

#Region "  Special Tech Link Management  "
    'Public mcolClosed As New Collection
    'Public mcolOpen As New Collection
    'Public mcolUnavailable As New Collection
    Public moSpecialTech(-1) As PlayerSpecialTech
    Public mlSpecialTechIdx(-1) As Int32
    Public mlSpecialTechUB As Int32 = -1


    'Public Sub PerformLinkTest()
    '    'Ok, run thru our unavailable list to see if we can move any of them to the Open list

    '    If mcolUnavailable.Count = 0 Then
    '        mcolUnavailable.Clear()
    '        For X As Int32 = 0 To glSpecialTechUB
    '            If glSpecialTechIdx(X) > -1 Then
    '                Dim sKey As String = "KEY" & goSpecialTechs(X).ObjectID
    '                If mcolClosed.Contains(sKey) = False AndAlso mcolOpen.Contains(sKey) = False Then
    '                    mcolUnavailable.Add(goSpecialTechs(X), sKey)
    '                End If
    '            End If
    '        Next X
    '        For X As Int32 = 0 To mlSpecialTechUB
    '            If mlSpecialTechIdx(X) > -1 Then
    '                If moSpecialTech(X) Is Nothing = False Then
    '                    Dim sKey As String = "KEY" & moSpecialTech(X).oTech.ObjectID
    '                    If mcolUnavailable.Contains(sKey) Then
    '                        mcolUnavailable.Remove(sKey)
    '                    End If
    '                End If
    '            End If
    '        Next X
    '    End If

    '    Dim bDone As Boolean = False
    '    While bDone = False
    '        bDone = True
    '        For Each oTech As SpecialTech In mcolUnavailable
    '            If oTech.bCanBeLinked = True AndAlso oTech.MeetsPreRequisites(oPlayer) > -1 Then
    '                bDone = False
    '                Dim oNew As PlayerSpecialTech = AddTechToSP_List(oTech)
    '                'oNew.yFlags = 0
    '                'oNew.ObjectID = oTech.ObjectID
    '                'oNew.ObjTypeID = oTech.ObjTypeID
    '                'oNew.Owner = oPlayer
    '                'oNew.LinkAttempts = 0
    '                'oNew.ComponentDevelopmentPhase = 0
    '                If mcolOpen.Contains("KEY" & oTech.ObjectID) = False Then mcolOpen.Add(oNew, "KEY" & oTech.ObjectID) 'oTech.sName)

    '                'mlSpecialTechUB += 1
    '                'ReDim Preserve moSpecialTech(mlSpecialTechUB)
    '                'ReDim Preserve mlSpecialTechIdx(mlSpecialTechUB)
    '                'moSpecialTech(mlSpecialTechUB) = oNew
    '                'mlSpecialTechIdx(mlSpecialTechUB) = oNew.oTech.ObjectID

    '                If mcolUnavailable.Contains("KEY" & oTech.ObjectID) = True Then mcolUnavailable.Remove("KEY" & oTech.ObjectID) 'oTech.sName)
    '                Exit For
    '            End If
    '        Next
    '    End While

    '    'Ok, unavailables have been tested, go through the specialtechs we have and cnt how many are linked
    '    Dim lLinks As Int32 = 0
    '    Dim lBypassed As Int32 = 0
    '    For X As Int32 = 0 To mlSpecialTechUB
    '        If mlSpecialTechIdx(X) > -1 AndAlso moSpecialTech(X).ComponentDevelopmentPhase <> 2 AndAlso moSpecialTech(X).bLinked = True Then
    '            If moSpecialTech(X).bInTheTank = False Then lLinks += 1 Else lBypassed += 1
    '        End If
    '    Next X
    '    If lBypassed > 25 Then Return 'if the player has bypassed more than 25 techs, do not link anything else

    '    'Ok, get our max links
    '    Dim lMaxLinks As Int32 = 3

    '    'now, if we can, test our links vs. max links
    '    If lLinks < lMaxLinks Then
    '        'First, we need to sort our open list...
    '        Dim colTemp As New Collection
    '        For Each oTech As PlayerSpecialTech In mcolOpen
    '            Dim bFound As Boolean = False

    '            oTech.lThisPassLinkChance = oTech.oTech.FallOffSuccess
    '            Dim lTmp As Int32 = oTech.oTech.MeetsPreRequisites(Me.oPlayer)
    '            If lTmp > 0 Then oTech.lThisPassLinkChance += lTmp

    '            For Each oTmp As PlayerSpecialTech In colTemp
    '                If oTech.lThisPassLinkChance > oTmp.lThisPassLinkChance Then
    '                    colTemp.Add(oTech, "KEY" & oTech.oTech.ObjectID, "KEY" & oTmp.oTech.ObjectID)
    '                    bFound = True
    '                    Exit For
    '                End If
    '            Next
    '            If bFound = False Then colTemp.Add(oTech, "KEY" & oTech.oTech.ObjectID) ' oTech.oTech.sName)
    '        Next
    '        mcolOpen = colTemp

    '        'go thru the OPEN list and run our link test
    '        bDone = False
    '        Dim bGotOne As Boolean = False
    '        While bDone = False
    '            If mcolOpen.Count = 0 AndAlso bGotOne = False Then
    '                'time to reset!!!
    '                'does the player have techs that have been bypassed?
    '                Dim lReturnedCnt As Int32 = 0
    '                For X As Int32 = 0 To mlSpecialTechUB
    '                    If mlSpecialTechIdx(X) > -1 Then
    '                        If moSpecialTech(X).ComponentDevelopmentPhase <> 2 AndAlso moSpecialTech(X).bInTheTank = True Then
    '                            'yes, so, return this one to the link
    '                            mcolClosed.Remove("KEY" & moSpecialTech(X).oTech.ObjectID)
    '                            mcolOpen.Add(moSpecialTech(X), "KEY" & moSpecialTech(X).oTech.ObjectID)
    '                            moSpecialTech(X).yFlags = moSpecialTech(X).yFlags Xor PlayerSpecialTech.eySpecialTechFlags.eInTheTank
    '                            lReturnedCnt += 1
    '                            If lReturnedCnt = 3 Then Continue While
    '                        End If
    '                    End If
    '                Next X
    '                If lReturnedCnt <> 0 Then Continue While

    '                'First, does the player have the item they selected?
    '                Dim bGuaranteedResearched As Boolean = False
    '                Dim bGTechLinkable As Boolean = True
    '                For X As Int32 = 0 To mlSpecialTechUB
    '                    If mlSpecialTechIdx(X) = oPlayer.GuaranteedSpecialTechID Then
    '                        bGTechLinkable = moSpecialTech(X).oTech.bCanBeLinked
    '                        If moSpecialTech(X).ComponentDevelopmentPhase = 2 Then
    '                            'yes
    '                            bGuaranteedResearched = True
    '                        End If
    '                        Exit For
    '                    End If
    '                Next X
    '                If bGuaranteedResearched = False AndAlso bGTechLinkable = True Then
    '                    If LinkNextPreqForTech(GetEpicaSpecialTech(oPlayer.GuaranteedSpecialTechID)) = True Then Return
    '                End If

    '                Me.oPlayer.SpecTechCostMult += 1
    '                For Each oTech As PlayerSpecialTech In mcolClosed
    '                    oTech.LinkAttempts = 0
    '                    oTech.yFlags = 0
    '                    oTech.SaveObject()
    '                    mcolOpen.Add(oTech, "KEY" & oTech.oTech.ObjectID) ' oTech.oTech.sName)
    '                Next
    '                mcolClosed.Clear()
    '                Exit While
    '            End If
    '            If mcolOpen.Count = 0 Then bDone = True

    '            For Each oTech As PlayerSpecialTech In mcolOpen
    '                oTech.LinkAttempts += 1
    '                If Rnd() * 100 < oTech.lThisPassLinkChance Then

    '                    oTech.yFlags = PlayerSpecialTech.eySpecialTechFlags.eLinked
    '                    mcolOpen.Remove("KEY" & oTech.oTech.ObjectID) ' oTech.oTech.sName)
    '                    oTech.ComponentDesigned()
    '                    'Ok, send player message, even if offline
    '                    Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
    '                    If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns) = True Then
    '                        Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech.GetResearchCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
    '                        Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech.GetProductionCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
    '                    End If
    '                    oTech.SaveObject()

    '                    bGotOne = True
    '                    bDone = Rnd() * 100 > 20
    '                    lLinks += 1
    '                    If lLinks >= lMaxLinks Then bDone = True
    '                    Continue While
    '                ElseIf oTech.LinkAttempts >= oTech.oTech.MaxLinkChanceAttempts Then
    '                    mcolOpen.Remove("KEY" & oTech.oTech.ObjectID) ' oTech.oTech.sName)
    '                    mcolClosed.Add(oTech, "KEY" & oTech.oTech.ObjectID) 'oTech.oTech.sName)

    '                    oTech.SaveObject()

    '                    Continue While
    '                End If
    '            Next
    '        End While
    '    End If

    'End Sub

    Public Sub PerformLinkTest()
        'First... check if we have a tech
        If mlSpecialTechUB = -1 Then
            'Ok, we don't so initialize the player
            For X As Int32 = 0 To glSpecialTechUB
                If glSpecialTechIdx(X) = 1 Then
                    HandleLinkTest(goSpecialTechs(X), 100)
                    'Return
                ElseIf glSpecialTechIdx(X) = 351 Then
                    HandleLinkTest(goSpecialTechs(X), 100)
                    '				ElseIf glSpecialTechIdx(X) = 277 Then
                    'HandleLinkTest(goSpecialTechs(X), 100)
                End If
            Next X
            Return
        End If

        If Me.oPlayer.AccountStatus <> AccountStatusType.eActiveAccount OrElse oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return

        Dim lChances() As Int32
        ReDim lChances(glSpecialTechUB)

        'Ok, now, go through the prerequisites
        For X As Int32 = 0 To glSpecTechPreqUB
            'Ok, special tech preqs...
            'TODO: As more PreqType are added, let's update this list, like MineralProperty, Mineral, etc...
            With goSpecTechPreq(X)
                Select Case .iPreqTypeID
                    Case ObjectType.eSpecialTech

                        Dim lIdx As Int32 = -1
                        For Y As Int32 = 0 To glSpecialTechUB
                            If .TechID = glSpecialTechIdx(Y) Then
                                lIdx = Y
                                Exit For
                            End If
                        Next Y

                        If lIdx <> -1 AndAlso lChances(lIdx) <> -1 Then
                            lChances(lIdx) = goSpecialTechs(lIdx).FallOffSuccess
                            Dim bFound As Boolean = False
                            For Y As Int32 = 0 To mlSpecialTechUB
                                If mlSpecialTechIdx(Y) = .lPreqID Then
                                    'we found it in our linked list...
                                    bFound = True

                                    'check our required value
                                    If .RequiredValue = 0 Then
                                        'Ok, add our chance to succeed ONLY if development phase <> researched
                                        If moSpecialTech(Y).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                            lChances(lIdx) += .ChanceToOpenLink
                                        End If
                                    ElseIf moSpecialTech(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                        'ok, add our chance only if we have researched it
                                        lChances(lIdx) += .ChanceToOpenLink
                                    Else : lChances(lIdx) = -1
                                    End If
                                End If
                            Next Y

                            'If we didn't find it, and it is a required prerequisite...
                            If bFound = False AndAlso .RequiredPrerequisite = True Then
                                'set our lchances to -1 so that it will never work
                                lChances(lIdx) = -1
                            End If
                            If goSpecialTechs(lIdx).bCanBeLinked = False Then lChances(lIdx) = -1
                        End If
                End Select
            End With
        Next X

        'Ok, our temp array
        Dim lIndices() As Int32
        Dim lSortedChances() As Int32
        Dim bAttemptedChance() As Boolean
        Dim lUB As Int32 = -1

        For X As Int32 = 0 To glSpecialTechUB
            If lChances(X) > 0 Then lUB += 1
        Next X
        ReDim lIndices(lUB)
        ReDim lSortedChances(lUB)
        ReDim bAttemptedChance(lUB)
        For X As Int32 = 0 To lUB
            lIndices(X) = -1 : lSortedChances(X) = -1 : bAttemptedChance(X) = False
        Next X

        'now fill our array
        Dim lEndIdx As Int32 = -1
        For X As Int32 = 0 To glSpecialTechUB
            If lChances(X) > 0 Then
                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To lUB
                    If lIndices(Y) = -1 Then
                        Exit For
                    ElseIf lSortedChances(Y) < lChances(X) Then
                        'we have found our spot so insert here
                        lIdx = Y
                        Exit For
                    End If
                Next Y

                If lIdx = -1 Then
                    lEndIdx += 1
                    lIdx = lEndIdx
                Else
                    'Here we are shifting our array
                    For Y As Int32 = lEndIdx To lIdx Step -1
                        lIndices(Y + 1) = lIndices(Y)
                        lSortedChances(Y + 1) = lSortedChances(Y)
                    Next Y
                    'increment lEndIdx AFTER assignment
                    lEndIdx += 1
                End If
                'Now, lIdx is where this value belongs
                lIndices(lIdx) = X
                lSortedChances(lIdx) = lChances(X)
            End If
        Next X

        Randomize()

        'Now, go through our chances...
        Dim lCurrentIdx As Int32 = 0
        Dim bDone As Boolean = False
        Dim bGotOne As Boolean = False

        Dim lActiveLinks As Int32 = 0
        Dim lMaxLinks As Int32 = 3 'oSpecials.yMaxLinks
        Dim lAddLinkChance As Int32 = 100 \ yMultipleLinkChance
        Dim lBypassed As Int32 = 0
        Dim lLinkable As Int32 = 0
        For X As Int32 = 0 To mlSpecialTechUB
            If moSpecialTech(X).bLinked = True AndAlso moSpecialTech(X).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                If moSpecialTech(X).bInTheTank = True Then lBypassed += 1 Else lActiveLinks += 1
            ElseIf moSpecialTech(X).bLinked = False Then
                If moSpecialTech(X).LinkAttempts < moSpecialTech(X).oTech.MaxLinkChanceAttempts AndAlso moSpecialTech(X).oTech.bCanBeLinked = True Then

                    'Safty check on past rolled specials, not yet linked, and the pre-reqs failed. This is then not concidered lLinkable.
                    Dim lIndex As Int32 = 0
                    For y As Int32 = 0 To glSpecTechPreqUB
                        If goSpecialTechs(y).ObjectID = moSpecialTech(X).ObjectID Then
                            lIndex = y
                            Exit For
                        End If
                    Next y
                    If lIndex > 0 Then
                        If lChances(lIndex) <> -1 Then
                            lLinkable += 1
                        Else
                            Console.WriteLine("PerformLinkTest: " & BytesToString(Me.oPlayer.PlayerName) & " - " & BytesToString(moSpecialTech(X).oTech.TechName) & " Rolled but pre-req is not met")
                        End If
                    Else
                        lLinkable += 1
                    End If
                End If
            End If
        Next X
        If lActiveLinks >= lMaxLinks OrElse lBypassed > 25 Then Return

        'Now, check for wormhole special tech
        If Me.oPlayer.lWormholeUB <> -1 AndAlso (lSuperSpecials And Player_Specials.SuperSpecialID.eWormholesTraversable) = 0 Then
            'Ok, forcefully place the wormhole traversable tech next if it is not already linked
            For X As Int32 = 0 To glSpecialTechUB
                If glSpecialTechIdx(X) <> -1 Then
                    If goSpecialTechs(X).ProgramControl = PlayerSpecialAttributeID.eSuperSpecials AndAlso goSpecialTechs(X).lNewValue = Player_Specials.SuperSpecialID.eWormholesTraversable Then
                        If HandleLinkTest(goSpecialTechs(X), 100) <> 0 Then Return
                        Exit For
                    End If
                End If
            Next X
        End If

        Dim lTestCnt As Int32 = 0
        Dim lCurrentChance As Int32
        If lCurrentIdx <= lUB Then lCurrentChance = lSortedChances(lCurrentIdx)

        While bDone = False

            lTestCnt += 1
            If lTestCnt > 1300 Then
                'gfrmDisplayForm.AddEventLine("TestCnt > 300 in PerformLinkTest, Breaking out now.")
                Exit While
            End If

            'Ok, determine our index from the list of items that match our current chance
            Dim lItemCnt As Int32 = 0
            Dim lNextChance As Int32 = 0
            For X As Int32 = 0 To lSortedChances.GetUpperBound(0)
                If lSortedChances(X) = lCurrentChance Then
                    If bAttemptedChance(X) = False Then lItemCnt += 1
                ElseIf lSortedChances(X) < lCurrentChance AndAlso lSortedChances(X) > lNextChance Then
                    lNextChance = lSortedChances(X)
                End If
            Next X
            If lItemCnt = 0 Then
                lCurrentChance = lNextChance
                bDone = (lCurrentChance = 0)

                If bDone = True Then
                    'Ok, verify that there are techs in my list that are still linkable
                    If lLinkable = 0 AndAlso lBypassed = 0 Then
                        'Ok, recycle... check for the guaranteed tech first...
                        'First, does the player have the item they selected?
                        Dim bGuaranteedResearched As Boolean = False
                        Dim bGTechLinkable As Boolean = True
                        For X As Int32 = 0 To mlSpecialTechUB
                            If mlSpecialTechIdx(X) = oPlayer.GuaranteedSpecialTechID Then
                                bGTechLinkable = moSpecialTech(X).oTech.bCanBeLinked
                                If moSpecialTech(X).ComponentDevelopmentPhase = 2 Then
                                    'yes
                                    bGuaranteedResearched = True
                                End If
                                Exit For
                            End If
                        Next X
                        If bGuaranteedResearched = False AndAlso bGTechLinkable = True Then
                            If LinkNextPreqForTech(GetEpicaSpecialTech(oPlayer.GuaranteedSpecialTechID)) = True Then Return
                        End If

                        'recycle all techs not linked
                        Me.oPlayer.SpecTechCostMult += 1
                        For X As Int32 = 0 To mlSpecialTechUB
                            Dim oTech As PlayerSpecialTech = moSpecialTech(X)
                            If oTech Is Nothing = False Then
                                If oTech.bLinked = False Then
                                    oTech.LinkAttempts = 0
                                    oTech.yFlags = 0
                                    oTech.SaveObject()
                                End If
                            End If
                        Next X
                        Exit While
                    End If
                End If

                Continue While
            Else
                Dim lTemp As Int32 = CInt(Rnd() * (lItemCnt - 1))
                For X As Int32 = 0 To lSortedChances.GetUpperBound(0)
                    If lSortedChances(X) = lCurrentChance AndAlso bAttemptedChance(X) = False Then
                        If lTemp < 1 Then
                            lCurrentIdx = X
                            Exit For
                        End If
                        lTemp -= 1
                    End If
                Next X
            End If

            If lCurrentIdx > lIndices.GetUpperBound(0) Then
                If bGotOne = False Then
                    lCurrentIdx = 0
                    lTestCnt = 0

                    bDone = True
                    For Y As Int32 = 0 To mlSpecialTechUB
                        'Was original bInQueue = True
                        If moSpecialTech(Y).LinkAttempts < moSpecialTech(Y).oTech.MaxLinkChanceAttempts AndAlso moSpecialTech(Y).bLinked = False Then
                            bDone = False
                            Exit For
                        End If
                    Next Y
                    'If bDone = True Then Stop
                Else : bDone = True
                End If
            Else
                bAttemptedChance(lCurrentIdx) = True
                'Go through list...
                Dim yValue As Byte = HandleLinkTest(goSpecialTechs(lIndices(lCurrentIdx)), lSortedChances(lCurrentIdx))
                If yValue = 0 Then
                    'lCurrentIdx += 1
                    bDone = False
                ElseIf yValue = 1 Then
                    bGotOne = True
                    bDone = Rnd() * 100 > lAddLinkChance

                    lActiveLinks += 1
                    If lActiveLinks >= lMaxLinks Then bDone = True

                    'lCurrentIdx += 1
                End If
            End If
        End While

        'MSC - 11/20/07 - Now, forcefully save the links to remove chance for lost data
        For X As Int32 = 0 To mlSpecialTechUB
            Try
                If mlSpecialTechIdx(X) <> -1 Then moSpecialTech(X).SaveObject()
            Catch ex As Exception
                LogEvent(LogEventType.Warning, "PerformLinkTest.ForceSave: " & ex.Message)
            End Try
        Next X


    End Sub

    Private Function LinkNextPreqForTech(ByRef oTech As SpecialTech) As Boolean
        If oTech Is Nothing Then Return True

        For X As Int32 = 0 To oTech.lPreqUB
            If oTech.oPreqs(X).iPreqTypeID = 51 Then
                If oTech.oPreqs(X).oTech.bCanBeLinked = False Then
                    If oTech.oPreqs(X).RequiredPrerequisite = False Then Continue For
                    Return False
                End If
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To mlSpecialTechUB
                    If mlSpecialTechIdx(Y) = oTech.oPreqs(X).lPreqID Then
                        bFound = True
                        If moSpecialTech(Y).bLinked = False Then
                            Return LinkNextPreqForTech(moSpecialTech(Y).oTech)
                        ElseIf moSpecialTech(Y).ComponentDevelopmentPhase <> 2 Then
                            Return True
                        End If
                        Exit For
                    End If
                Next Y
                If bFound = False Then
                    Return LinkNextPreqForTech(GetEpicaSpecialTech(oTech.oPreqs(X).lPreqID))
                End If
            End If
        Next X

        'if we are here, then link the special tech for htis
        For X As Int32 = 0 To mlSpecialTechUB
            If mlSpecialTechIdx(X) = oTech.ObjectID Then
                moSpecialTech(X).yFlags = PlayerSpecialTech.eySpecialTechFlags.eLinked

                moSpecialTech(X).ComponentDesigned()
                'Ok, send player message, even if offline
                Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(X), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns) = True Then
                    Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(X).GetResearchCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                    Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(X).GetProductionCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                End If

                moSpecialTech(X).SaveObject()

                'If mcolOpen.Contains("KEY" & oTech.ObjectID) Then mcolOpen.Remove("KEY" & oTech.ObjectID) 'oTech.sName)
                Return True
            End If
        Next X

        Dim oNew As PlayerSpecialTech = AddTechToSP_List(oTech)
        oNew.yFlags = PlayerSpecialTech.eySpecialTechFlags.eLinked
        oNew.LinkAttempts = 1
        oNew.ComponentDesigned()

        Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oNew, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
        If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns) = True Then
            Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oNew.GetResearchCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
            Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oNew.GetProductionCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
        End If
        oNew.SaveObject()

        Return True
    End Function

    Private Function AddTechToSP_List(ByRef oTech As SpecialTech) As PlayerSpecialTech
        Dim oNew As New PlayerSpecialTech()
        oNew.yFlags = 0
        oNew.ObjectID = oTech.ObjectID
        oNew.ObjTypeID = oTech.ObjTypeID
        oNew.Owner = oPlayer
        oNew.LinkAttempts = 0
        oNew.ComponentDevelopmentPhase = 0
        oNew.ResearchAttempts = 0
        oNew.SuccessChanceIncrement = oTech.IncrementalSuccess
        oNew.RandomSeed = 1.0F
        oNew.ErrorReasonCode = 0
        oNew.CurrentSuccessChance = oTech.InitialSuccessChance
        oNew.SaveObject()

        mlSpecialTechUB += 1
        ReDim Preserve moSpecialTech(mlSpecialTechUB)
        ReDim Preserve mlSpecialTechIdx(mlSpecialTechUB)
        moSpecialTech(mlSpecialTechUB) = oNew
        mlSpecialTechIdx(mlSpecialTechUB) = oNew.oTech.ObjectID

        Return oNew
    End Function

    'Public Sub PerformLinkTest()
    '    'First... check if we have a tech
    '    If mlSpecialTechUB = -1 Then
    '        'Ok, we don't so initialize the player
    '        For X As Int32 = 0 To glSpecialTechUB
    '            If glSpecialTechIdx(X) = 1 Then
    '                HandleLinkTest(goSpecialTechs(X), 100)
    '                'Return
    '            ElseIf glSpecialTechIdx(X) = 351 Then
    '                HandleLinkTest(goSpecialTechs(X), 100)
    '                '				ElseIf glSpecialTechIdx(X) = 277 Then
    '                'HandleLinkTest(goSpecialTechs(X), 100)
    '            End If
    '        Next X
    '        Return
    '    End If

    '    If Me.oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return

    '    Dim lChances() As Int32
    '    ReDim lChances(glSpecialTechUB)

    '    'Ok, now, go through the prerequisites
    '    For X As Int32 = 0 To glSpecTechPreqUB
    '        'Ok, special tech preqs...
    '        'TODO: As more PreqType are added, let's update this list, like MineralProperty, Mineral, etc...
    '        With goSpecTechPreq(X)
    '            Select Case .iPreqTypeID
    '                Case ObjectType.eSpecialTech

    '                    Dim lIdx As Int32 = -1
    '                    For Y As Int32 = 0 To glSpecialTechUB
    '                        If .TechID = glSpecialTechIdx(Y) Then
    '                            lIdx = Y
    '                            Exit For
    '                        End If
    '                    Next Y

    '                    If lIdx <> -1 AndAlso lChances(lIdx) <> -1 Then
    '                        lChances(lIdx) = goSpecialTechs(lIdx).FallOffSuccess
    '                        Dim bFound As Boolean = False
    '                        For Y As Int32 = 0 To mlSpecialTechUB
    '                            If mlSpecialTechIdx(Y) = .lPreqID Then
    '                                'we found it in our linked list...
    '                                bFound = True

    '                                'check our required value
    '                                If .RequiredValue = 0 Then
    '                                    'Ok, add our chance to succeed ONLY if development phase <> researched
    '                                    If moSpecialTech(Y).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
    '                                        lChances(lIdx) += .ChanceToOpenLink
    '                                    End If
    '                                ElseIf moSpecialTech(Y).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
    '                                    'ok, add our chance only if we have researched it
    '                                    lChances(lIdx) += .ChanceToOpenLink
    '                                Else : lChances(lIdx) = -1
    '                                End If
    '                            End If
    '                        Next Y

    '                        'If we didn't find it, and it is a required prerequisite...
    '                        If bFound = False AndAlso .RequiredPrerequisite = True Then
    '                            'set our lchances to -1 so that it will never work
    '                            lChances(lIdx) = -1
    '                        End If
    '                        If goSpecialTechs(lIdx).bInGuaranteeList = True Then lChances(lIdx) = -1
    '                    End If
    '            End Select
    '        End With
    '    Next X

    '    'Ok, our temp array
    '    Dim lIndices() As Int32
    '    Dim lSortedChances() As Int32
    '    Dim bAttemptedChance() As Boolean
    '    Dim lUB As Int32 = -1

    '    For X As Int32 = 0 To glSpecialTechUB
    '        If lChances(X) > 0 Then lUB += 1
    '    Next X
    '    ReDim lIndices(lUB)
    '    ReDim lSortedChances(lUB)
    '    ReDim bAttemptedChance(lUB)
    '    For X As Int32 = 0 To lUB
    '        lIndices(X) = -1 : lSortedChances(X) = -1 : bAttemptedChance(X) = False
    '    Next X

    '    'now fill our array
    '    Dim lEndIdx As Int32 = -1
    '    For X As Int32 = 0 To glSpecialTechUB
    '        If lChances(X) > 0 Then
    '            Dim lIdx As Int32 = -1
    '            For Y As Int32 = 0 To lUB
    '                If lIndices(Y) = -1 Then
    '                    Exit For
    '                ElseIf lSortedChances(Y) < lChances(X) Then
    '                    'we have found our spot so insert here
    '                    lIdx = Y
    '                    Exit For
    '                End If
    '            Next Y

    '            If lIdx = -1 Then
    '                lEndIdx += 1
    '                lIdx = lEndIdx
    '            Else
    '                'Here we are shifting our array
    '                For Y As Int32 = lEndIdx To lIdx Step -1
    '                    lIndices(Y + 1) = lIndices(Y)
    '                    lSortedChances(Y + 1) = lSortedChances(Y)
    '                Next Y
    '                'increment lEndIdx AFTER assignment
    '                lEndIdx += 1
    '            End If
    '            'Now, lIdx is where this value belongs
    '            lIndices(lIdx) = X
    '            lSortedChances(lIdx) = lChances(X)
    '        End If
    '    Next X

    '    Randomize()

    '    'Now, go through our chances...
    '    Dim lCurrentIdx As Int32 = 0
    '    Dim bDone As Boolean = False
    '    Dim bGotOne As Boolean = False

    '    Dim lActiveLinks As Int32 = 0
    '    Dim lMaxLinks As Int32 = 3 'oSpecials.yMaxLinks
    '    Dim lAddLinkChance As Int32 = 100 \ yMultipleLinkChance
    '    Dim lBypassed As Int32 = 0
    '    For X As Int32 = 0 To mlSpecialTechUB
    '        If moSpecialTech(X).bLinked = True AndAlso moSpecialTech(X).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
    '            If moSpecialTech(X).bInTheTank = True Then lBypassed += 1 Else lActiveLinks += 1
    '        End If
    '    Next X
    '    If lActiveLinks >= lMaxLinks OrElse lBypassed > 25 Then Return

    '    'Now, check for wormhole special tech
    '    If Me.oPlayer.lWormholeUB <> -1 AndAlso (lSuperSpecials And Player_Specials.SuperSpecialID.eWormholesTraversable) = 0 Then
    '        'Ok, forcefully place the wormhole traversable tech next if it is not already linked
    '        For X As Int32 = 0 To glSpecialTechUB
    '            If glSpecialTechIdx(X) <> -1 Then
    '                If goSpecialTechs(X).ProgramControl = PlayerSpecialAttributeID.eSuperSpecials AndAlso goSpecialTechs(X).lNewValue = Player_Specials.SuperSpecialID.eWormholesTraversable Then
    '                    If HandleLinkTest(goSpecialTechs(X), 100) <> 0 Then Return
    '                    Exit For
    '                End If
    '            End If
    '        Next X
    '    End If

    '    Dim lTestCnt As Int32 = 0
    '    Dim lCurrentChance As Int32
    '    If lCurrentIdx <= lUB Then lCurrentChance = lSortedChances(lCurrentIdx)

    '    While bDone = False

    '        lTestCnt += 1
    '        If lTestCnt > 1300 Then
    '            'gfrmDisplayForm.AddEventLine("TestCnt > 300 in PerformLinkTest, Breaking out now.")
    '            Exit While
    '        End If

    '        'Ok, determine our index from the list of items that match our current chance
    '        Dim lItemCnt As Int32 = 0
    '        Dim lNextChance As Int32 = 0
    '        For X As Int32 = 0 To lSortedChances.GetUpperBound(0)
    '            If lSortedChances(X) = lCurrentChance Then
    '                If bAttemptedChance(X) = False Then lItemCnt += 1
    '            ElseIf lSortedChances(X) < lCurrentChance AndAlso lSortedChances(X) > lNextChance Then
    '                lNextChance = lSortedChances(X)
    '            End If
    '        Next X
    '        If lItemCnt = 0 Then
    '            lCurrentChance = lNextChance
    '            bDone = (lCurrentChance = 0)
    '            Continue While
    '        Else
    '            Dim lTemp As Int32 = CInt(Rnd() * (lItemCnt - 1))
    '            For X As Int32 = 0 To lSortedChances.GetUpperBound(0)
    '                If lSortedChances(X) = lCurrentChance AndAlso bAttemptedChance(X) = False Then
    '                    If lTemp < 1 Then
    '                        lCurrentIdx = X
    '                        Exit For
    '                    End If
    '                    lTemp -= 1
    '                End If
    '            Next X
    '        End If

    '        If lCurrentIdx > lIndices.GetUpperBound(0) Then
    '            If bGotOne = False Then
    '                lCurrentIdx = 0
    '                lTestCnt = 0

    '                bDone = True
    '                For Y As Int32 = 0 To mlSpecialTechUB
    '                    'Was original bInQueue = True
    '                    If moSpecialTech(Y).LinkAttempts < moSpecialTech(Y).oTech.MaxLinkChanceAttempts AndAlso moSpecialTech(Y).bLinked = False Then
    '                        bDone = False
    '                        Exit For
    '                    End If
    '                Next Y
    '                'If bDone = True Then Stop
    '            Else : bDone = True
    '            End If
    '        Else
    '            bAttemptedChance(lCurrentIdx) = True
    '            'Go through list...
    '            Dim yValue As Byte = HandleLinkTest(goSpecialTechs(lIndices(lCurrentIdx)), lSortedChances(lCurrentIdx))
    '            If yValue = 0 Then
    '                'lCurrentIdx += 1
    '                bDone = False
    '            ElseIf yValue = 1 Then
    '                bGotOne = True
    '                bDone = Rnd() * 100 > lAddLinkChance

    '                lActiveLinks += 1
    '                If lActiveLinks >= lMaxLinks Then bDone = True

    '                'lCurrentIdx += 1
    '            End If
    '        End If
    '    End While

    '    'MSC - 11/20/07 - Now, forcefully save the links to remove chance for lost data
    '    For X As Int32 = 0 To mlSpecialTechUB
    '        Try
    '            If mlSpecialTechIdx(X) <> -1 Then moSpecialTech(X).SaveObject()
    '        Catch ex As Exception
    '            LogEvent(LogEventType.Warning, "PerformLinkTest.ForceSave: " & ex.Message)
    '        End Try
    '    Next X

    'End Sub

    Private Function HandleLinkTest(ByRef oSpecTech As SpecialTech, ByVal lChance As Int32) As Byte
        'First, see if the item is already in the link list
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlSpecialTechUB
            If mlSpecialTechIdx(X) = oSpecTech.ObjectID Then

                'Is the tech link failed already?
                If moSpecialTech(X).LinkAttempts > oSpecTech.MaxLinkChanceAttempts Then Return 0
                If moSpecialTech(X).bLinked = True Then Return 0

                lIdx = X
                Exit For
            End If
        Next X

        'Safty check on pre-reqs
        For x As Int32 = 0 To oSpecTech.lPreqUB
            If oSpecTech.oPreqs(x).iPreqTypeID = 51 Then
                If oSpecTech.oPreqs(x).oTech.bCanBeLinked = False Then Return 0
                Dim bFound As Boolean = False
                'Console.WriteLine(oSpecTech.ObjectID.ToString & ":" & BytesToString(oSpecTech.TechName) & " - Requires - " & oSpecTech.oPreqs(x).lPreqID)
                For y As Int32 = 0 To mlSpecialTechUB
                    If mlSpecialTechIdx(y) = oSpecTech.oPreqs(x).lPreqID Then
                        bFound = True
                        'Console.WriteLine(oSpecTech.ObjectID.ToString & ":" & BytesToString(oSpecTech.TechName) & " - Requires - " & moSpecialTech(y).oTech.ObjectID.ToString & ":" & BytesToString(moSpecialTech(y).oTech.TechName) & " which is at phase - " & CType(moSpecialTech(y).ComponentDevelopmentPhase, Epica_Tech.eComponentDevelopmentPhase).ToString)
                        If moSpecialTech(y).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                            Console.WriteLine("HandleLinkTest: " & BytesToString(Me.oPlayer.PlayerName) & " - Attempted Link " & oSpecTech.ObjectID.ToString & ":" & BytesToString(oSpecTech.TechName) & " but pre-req of " & moSpecialTech(y).oTech.ObjectID.ToString & ":" & BytesToString(moSpecialTech(y).oTech.TechName) & " is not linked")
                            Return 0
                        End If
                        Exit For
                    End If
                Next y

                'They have not rolled, let alone linked, the needed pre-req
                If bFound = False Then
                    Console.WriteLine("HandleLinkTest: " & BytesToString(Me.oPlayer.PlayerName) & " - Attempted Link " & oSpecTech.ObjectID.ToString & ":" & BytesToString(oSpecTech.TechName) & " but pre-req is not rolled.")
                    Return 0
                End If
            End If
        Next

        'Ok, is it in the link list?
        If lIdx = -1 Then
            'no, so add it
            lIdx = mlSpecialTechUB + 1
            ReDim Preserve moSpecialTech(lIdx)
            ReDim Preserve mlSpecialTechIdx(lIdx)
            moSpecialTech(lIdx) = New PlayerSpecialTech()

            With moSpecialTech(lIdx)
                '.bLinked = False
                .yFlags = 0
                .ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eComponentDesign
                .CurrentSuccessChance = oSpecTech.InitialSuccessChance
                .ErrorReasonCode = 0
                .LinkAttempts = 0
                .ObjectID = oSpecTech.ObjectID
                .ObjTypeID = ObjectType.eSpecialTech
                .Owner = Me.oPlayer
                .RandomSeed = 1.0F
                .ResearchAttempts = 0
                .SuccessChanceIncrement = oSpecTech.IncrementalSuccess
            End With
            mlSpecialTechIdx(lIdx) = oSpecTech.ObjectID
            mlSpecialTechUB += 1
        End If

        With moSpecialTech(lIdx)
            'ok, if we are here, increment our link attempts
            '.bNeedsSaved = True
            .LinkAttempts += 1
            If Rnd() * 100.0F < lChance Then
                'ok, it linked!!! so call ComponentDesigned
                .ComponentDesigned()

                If (Me.oPlayer.iInternalEmailSettings And eEmailSettings.eResearchComplete) <> 0 Then
                    Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder
                    oSB.AppendLine("Your scientists want to inform you that you have discovered " & BytesToString(.oTech.TechName) & ".")
                    Dim oPC As PlayerComm = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Special Projects Discovery", Me.oPlayer.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.oPlayer.PlayerName), Nothing)
                    If oPC Is Nothing = False Then Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If

                'Ok, send player message, even if offline
                Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(lIdx), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns) = True Then
                    Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(lIdx).GetResearchCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                    Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(lIdx).GetProductionCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                End If
                .SaveObject()
                Return 1
            End If
            .SaveObject()
        End With

        Return 0
    End Function
#End Region

    Public lSuperSpecials As SuperSpecialID = SuperSpecialID.eNoSuperSpecials

    'Indicative of rate at which minerals can automatically be moved. 1 per X where X is seconds.
    Public yAutomaticMinMove As Byte = 0

    'Indicates the Command Point limit for the player
    Public iCPLimit As Int16 = 300

    'Miscellaneous...
    Public yDiscoverTime As Byte = 1                    'divided into the Points Required for Discover costs
    Public yMineralConcentrationBonus As Byte = 0       'added to the concentration of a mineral cache when mining
    Public yEnlistedColonistCost As Byte = 20           'Cost in Number of colonists per enlistedx10 produced
    Public yOfficersEnlistedCost As Byte = 20           'cost in number of enlisted per officerx10 produced
    Public yStudyMineralEffect As Byte = 1              'multiplier applied to the effects of studying minerals
    Public yRepairSpeedBonus As Byte = 0                'stored as percentage. Indicates efficiency for repairing.
    Public yPersonnelCargoUsage As Byte = 1             'amount of cargo space used per personnel cargo entry

    'Special Tech Links and research-specific
    Public yMultipleLinkChance As Byte = 5              'rated as 1 in X... so 5 would be 20% of time, 4 would be 25%, 3 would be 33%, etc...
    Public yMaxLinks As Byte = 3                        'maximum links at a given time

    Public myResearchFacAssistBonus As Byte = 50       'stored as percentage. Indicates production granted from an assisting research facility

    'Battlegroups
    Public yMaxBattlegroups As Byte = 2                 'maximum number of battlegroups permitted at once
    Public yMaxBattleGroupUnits As Byte = 10            'maximum number of units per battlegroup

    Public yNavalAbility As Byte = 0                    '1 = Naval Units, 2 = Underwater Facilities
    Public yPayloadTypeAvailable As Byte = 1            '0 = none, 1 = explosive, 2 = Chemical, 3 = Magnetic
    Public yAlloyBuilderImprovements As Byte = 3        'number of improvements allowed
    Public yMaxAlloyImprovement As Byte = 1             'max level of improvement available

    Public yControlledGrowthLimit As Byte = 0
    Public yControlledMoraleLimit As Byte = 0

#Region "  Espionage Missions  "
    Public yCommandEspionage As Byte = 0               '1 = Battlegroup Composition, 2 = Battlegroup Commands, 3 = Waypoints
    Public yConstructionEspionage As Byte = 0          '1 = Intercept Command Centers Built, 2 = Intercept All Buildings
    Public yDiplomaticEspionage As Byte = 0            '1 = View Allies List, 2 = View All Contacts
    Public yResearchEspionage As Byte = 0              '1 = Special techs
    Public yEconomicEspionage As Byte = 0              '1 = Direct Trades
#End Region

#Region "  Super Specials  "
    Public yPrecisionMissiles As Byte = 0               '1 = Power, 2 = Radar, 3 = Shields, 4 = Thrust
    Public yTransporters As Byte = 0                    '1 = Planetside Only, 2 = Planet to Station
    Public yCloaking As Byte = 0                        'BITWISE: 1 = Ground Units, 2 = Facilities, 4 = Aerial
#End Region

#Region "  Upkeep Cost Reductions  "
    'Upkeep Reductions... values are stored as a percentage. Applied during the Budget section.
    Private myAerialUpkeepReduct As Byte = 0
    Public fAerialUpkeepMult As Single = 1.0F
    Private myFactoryUpkeepReduct As Byte = 0
    Public fFactoryUpkeepMult As Single = 1.0F
    Private myNonAerialUpkeepReduct As Byte = 0
    Public fNonAerialUpkeepMult As Single = 1.0F
    Private myOtherFacUpkeepReduct As Byte = 0
    Public fOtherFacUpkeepMult As Single = 1.0F
    Private myPopulationUpkeepReduct As Byte = 0
    Public fPopulationUpkeepMult As Single = 1.0F
    Private myResearchFacUpkeepReduct As Byte = 0
    Public fResearchFacUpkeepMult As Single = 1.0F
    Private mySpaceportUpkeepReduct As Byte = 0
    Public fSpaceportUpkeepMult As Single = 1.0F
    Private myUnderwaterUpkeepReduct As Byte = 0
    Public fUnderwaterUpkeepMult As Single = 1.0F
    Private myUnemploymentUpkeepReduct As Byte = 0
    Public fUnemploymentUpkeepMult As Single = 1.0F

#End Region

#Region "  Colony Stats Specific  "
    Public yHomelessAllowedBeforePenalty As Byte = 0        'stored as a percentage
    Public yTaxRateBeforePenalty As Byte = 0                'stored as a percentage
    Public yUnemploymentBeforePenalty As Byte = 0           'stored as a percentage
    Public yColonyMoraleBoost As Byte = 0                   'Added directly to all colony's morales
    Public yGrowthRateBonus As Byte = 0                     'Added to the growth rate of the colony
#End Region

#Region "  Tradepost Specifics  "
    'Max Slots
    Public yBuyOrderSlots As Byte = 1 '0
    Public ySellOrderSlots As Byte = 10 '5
    Public yBuySellTechLvl As Byte = 0

    Public yTradeBoardRange As Byte = 3                     '0 = none, 1 = planet, 2 = system, 3 = sector, 4 = galaxy
    Public yTradeCosts As Byte = 25                         'stored as percentage. Used to determine GTC taxation.
    Public yTradeGTCSpeed As Byte = 0                       '0 = 100, 1 = 300, 2 = 500, 3 = 800

#End Region

#Region "  Component Builder Specifics  "
    Public HullPerResidence As Byte = 11
    Public EnlistedTrainingFactor As Byte = 5
    Public BaseEnlistedPerTrainer As Byte = 10
    Public BaseOfficerPerTrainer As Byte = 5
    Public OfficerTrainingFactor As Byte = 10
    Public MiningWorkers As Byte = 72
    Public ResearchCrewQtrs As Byte = 30
    Public ResearchProdFactor As Byte = 100
    Public PowerGenProdToWorker As Byte = 4

#Region " Armor "
    'Bonuses applied to armor resistances after final research of the armor plate. Stored as a Percentage. 
    Public yBeamResistImprove As Byte = 0
    Public yBurnResistImprove As Byte = 0
    Public yChemResistImprove As Byte = 0
    Public yImpactResistImprove As Byte = 0
    Public yMagneticResistImprove As Byte = 0
    Public yPierceResistImprove As Byte = 0
    Public yRadarResistImprove As Byte = 0

    'Cheaper Hull To Hp. Relates to cell K4 in MattsAlloyBuilder.Matts Armor spreadsheet
    Public yArmorHullToHP As Byte = 10

    Public lNoResistAt2X As Int32 = 0
#End Region
#Region " Bombs "
    Public MinBombROF As Int32 = 900    '30 seconds
    Public MaxBombAOE As Byte = 10
    Public MaxBombGuidance As Byte = 10
    Public MaxBombRange As Byte = 30
    Public BombPayloadType As Byte = 0
#End Region
#Region " Engine "
    'Limits
    Public yMaxSpeed As Byte = 30
    Public lPowerThrustLimit As Int32 = 10000
    Public yMaxManeuver As Byte = 30

    'Bonuses applied to results. Stored as a percentage.
    Public yPowerBonus As Byte = 0
    Public yManeuverBonus As Byte = 0
    Public yMaxSpeedBonus As Byte = 0
#End Region
#Region " Radar "
    'Limits
    Public yJamEffectAvailable As Byte = 1         'bit-wise
    Public yJamImmunityAvailable As Byte = 1       'bit-wise
    Public yJamStrength As Byte = 50
    Public yJamTargets As Byte = 1                 'ranged value with 128 being bitwise indicating Area Effect (0)
    Public yRadarDisRes As Byte = 50
    Public yRadarMaxRange As Byte = 50
    Public yRadarOptRange As Byte = 50
    Public yRadarScanRes As Byte = 50
    Public yRadarWpnAcc As Byte = 50

    'Bonuses applied to final product. Stored as percentages.
    Public yRadarMaxRangeBonus As Byte = 0
    Public yRadarOptRangeBonus As Byte = 0
    Public yRadarScanResBonus As Byte = 0
    Public yRadarWpnAccBonus As Byte = 0

#End Region
#Region " Missiles "
    'Limits

    'Bonuses applied to results. stored as a percentage
    Public yMissileExpRadBonus As Byte = 0
    Public yMissileFlightTimeBonus As Byte = 0
    Public yMissileHullSizeBonus As Byte = 0
    Public yMissileManeuverBonus As Byte = 0
    Public yMissileMaxDmgBonus As Byte = 0
    Public yMissileMaxSpeedBonus As Byte = 0
#End Region
#Region " Projectiles "
    'Limits
    Public yProjExpRadius As Byte = 10
    Public iProjMaxRange As Int16 = 50

    'Bonuses
#End Region
#Region " Pulse Beam "
    'Limits
    Public iPulseOptRng As Int16 = 30
    Public lPulseMaxCompressFactor As Int32 = 2
    Public iPulseWpnROF As Int16 = 240            'cycles (8 seconds)
    Public lPulseMaxInputEnergy As Int32 = 30

    'Bonuses... applied to the final product. Stored as a percentage.
    Public yPulseMaxDmgBonus As Byte = 0
    Public yPulseMinDmgBonus As Byte = 0
    Public yPulseOptRngBonus As Byte = 0
    Public yPulseROFBonus As Byte = 0
    Public yPulsePowerReduced As Byte = 0
#End Region
#Region " Shields "
    'Limits
    Public lShieldMaxHP As Int32 = 50
    Public lShieldProjHullSize As Int32 = 500
    Public iShieldRechargeFreq As Int16 = 150             'minimum value... in cycles (5 seconds)
    Public lShieldRechargeRate As Int32 = 50

    'Bonuses... applied to the final product. Stored as a percentage.
    Public yShieldMaxHPBonus As Byte = 0
    Public yShieldProjHullSizeBonus As Byte = 0
    Public yShieldRechargeFreqBonus As Byte = 0           'decreases
    Public yShieldRechargeRateBonus As Byte = 0
#End Region
#Region " Solid Beam "
    'Limits
    Public iSolidBeamOptimumRange As Int16 = 40
    Public ySolidBeamPierceRatio As Byte = 10
    Public ySolidBeamMaxAccuracy As Byte = 25
    Public lSolidBeamMaxDmg As Int32 = 40          'maximum
    Public iSolidBeamROF As Int16 = 360           'cycles (12 seconds)

    'Bonuses... applied to the final product. Stored as a percentage.
    Public ySolidBeamMaxDmgBonus As Byte = 0
    Public ySolidBeamMinDmgBonus As Byte = 0
    Public ySolidBeamOptRngBonus As Byte = 0
    Public ySolidBeamPowerBonus As Byte = 0           'reduces pwoer
    Public ySolidBeamROFBonus As Byte = 0             'reduces ROF
#End Region
#End Region

    Public Shared Function GetPropertyValueByID(ByRef oObj As Player_Specials, ByVal lID As Int32) As Int32
        Dim lValue As Int32 = 0
        With oObj
            Select Case CType(lID, PlayerSpecialAttributeID)
                Case PlayerSpecialAttributeID.eAerialUpkeepReduction
                    lValue = .myAerialUpkeepReduct
                Case PlayerSpecialAttributeID.eAlloyBuilderImprovements
                    lValue = .yAlloyBuilderImprovements
                Case PlayerSpecialAttributeID.eAreaEffectJamming
                    If (.yJamTargets And 128) <> 0 Then
                        lValue = 1
                    Else : lValue = 0
                    End If
                Case PlayerSpecialAttributeID.eArmorCheaperHullToHP
                    lValue = .yArmorHullToHP
                Case PlayerSpecialAttributeID.eAutomaticMineralMovement
                    lValue = .yAutomaticMinMove
                Case PlayerSpecialAttributeID.eBaseEnlistedPerTrainer
                    lValue = .BaseEnlistedPerTrainer
                Case PlayerSpecialAttributeID.eBaseOfficerPerTrainer
                    lValue = .BaseOfficerPerTrainer
                Case PlayerSpecialAttributeID.eBeamResistImprove
                    lValue = .yBeamResistImprove
                Case PlayerSpecialAttributeID.eBetterHullToResidence
                    lValue = .HullPerResidence
                Case PlayerSpecialAttributeID.eBombMaxAOE
                    lValue = .MaxBombAOE
                Case PlayerSpecialAttributeID.eBombMaxGuidance
                    lValue = .MaxBombGuidance
                Case PlayerSpecialAttributeID.eBombMaxRange
                    lValue = .MaxBombRange
                Case PlayerSpecialAttributeID.eBombMinROF
                    lValue = .MinBombROF
                Case PlayerSpecialAttributeID.eBombPayloadType
                    lValue = .BombPayloadType
                Case PlayerSpecialAttributeID.eBurnResistImprove
                    lValue = .yBurnResistImprove
                Case PlayerSpecialAttributeID.eBuyAndSellOrderSlots
                    lValue = .yBuySellTechLvl
                Case PlayerSpecialAttributeID.eBuyOrderSlots
                    lValue = .yBuyOrderSlots
                Case PlayerSpecialAttributeID.eChemResistImprove
                    lValue = .yChemResistImprove
                Case PlayerSpecialAttributeID.eColonyMoraleBoost
                    lValue = .yColonyMoraleBoost
                Case PlayerSpecialAttributeID.eCommandEspionageAbility
                    lValue = .yCommandEspionage
                Case PlayerSpecialAttributeID.eConstructionEspionageAbility
                    lValue = .yConstructionEspionage
                Case PlayerSpecialAttributeID.eControlledGrowth
                    lValue = .yControlledGrowthLimit
                Case PlayerSpecialAttributeID.eControlledMorale
                    lValue = .yControlledMoraleLimit
                Case PlayerSpecialAttributeID.eCPLimit
                    lValue = .iCPLimit
                Case PlayerSpecialAttributeID.eDiscoverTime
                    lValue = .yDiscoverTime
                Case PlayerSpecialAttributeID.eDiplomaticEspionage
                    lValue = .yDiplomaticEspionage
                Case PlayerSpecialAttributeID.eEngineMaxSpeed
                    lValue = .yMaxSpeed
                Case PlayerSpecialAttributeID.eEnginePowerBonus
                    lValue = .yPowerBonus
                Case PlayerSpecialAttributeID.eEnginePowerThrustLimit
                    lValue = .lPowerThrustLimit
                Case PlayerSpecialAttributeID.eEnlistedsColonistsCost
                    lValue = .yEnlistedColonistCost
                Case PlayerSpecialAttributeID.eEnlistedTrainingFactor
                    lValue = .EnlistedTrainingFactor
                Case PlayerSpecialAttributeID.eFactoryUpkeepReduction
                    lValue = .myFactoryUpkeepReduct
                Case PlayerSpecialAttributeID.eGrowthRateBonus
                    lValue = .yGrowthRateBonus
                Case PlayerSpecialAttributeID.eHomelessAllowedBeforePenalty
                    lValue = .yHomelessAllowedBeforePenalty
                Case PlayerSpecialAttributeID.eImpactResistImprove
                    lValue = .yImpactResistImprove
                Case PlayerSpecialAttributeID.eJammingEffectAvailable
                    lValue = .yJamEffectAvailable
                Case PlayerSpecialAttributeID.eJammingImmunityAvailable
                    lValue = .yJamImmunityAvailable
                Case PlayerSpecialAttributeID.eJammingStrength
                    lValue = .yJamStrength
                Case PlayerSpecialAttributeID.eJammingTargets
                    lValue = .yJamTargets
                Case PlayerSpecialAttributeID.eLinkChanceBonus
                    lValue = .yMultipleLinkChance
                Case PlayerSpecialAttributeID.eMagResistImprove
                    lValue = .yMagneticResistImprove
                Case PlayerSpecialAttributeID.eManeuverBonus
                    lValue = .yManeuverBonus
                Case PlayerSpecialAttributeID.eManeuverLimit
                    lValue = .yMaxManeuver
                Case PlayerSpecialAttributeID.eMaxAlloyImprovement
                    lValue = .yMaxAlloyImprovement
                Case PlayerSpecialAttributeID.eMaxBattlegroups
                    lValue = .yMaxBattlegroups
                Case PlayerSpecialAttributeID.eMaxBattlegroupUnits
                    lValue = .yMaxBattleGroupUnits
                Case PlayerSpecialAttributeID.eMaxnumberofconcurrentLinks
                    lValue = .yMaxLinks
                Case PlayerSpecialAttributeID.eMaxSpeedBonus
                    lValue = .yMaxSpeedBonus
                Case PlayerSpecialAttributeID.eMineralConcentrationBonus
                    lValue = .yMineralConcentrationBonus
                Case PlayerSpecialAttributeID.eMiningWorkers
                    lValue = .MiningWorkers
                Case PlayerSpecialAttributeID.eMissileExplosionRadiusBonus
                    lValue = .yMissileExpRadBonus
                Case PlayerSpecialAttributeID.eMissileFlightTimeBonus
                    lValue = .yMissileFlightTimeBonus
                Case PlayerSpecialAttributeID.eMissileHullSizeImprove
                    lValue = .yMissileHullSizeBonus
                Case PlayerSpecialAttributeID.eMissileManeuverBonus
                    lValue = .yMissileManeuverBonus
                Case PlayerSpecialAttributeID.eMissileMaxDmgBonus
                    lValue = .yMissileMaxDmgBonus
                Case PlayerSpecialAttributeID.eMissileMaxSpeedImprove
                    lValue = .yMissileMaxSpeedBonus
                Case PlayerSpecialAttributeID.eNavalAbility
                    lValue = .yNavalAbility
                Case PlayerSpecialAttributeID.eNonAerialUpkeepReduction
                    lValue = .myNonAerialUpkeepReduct
                Case PlayerSpecialAttributeID.eNoResistAt2X
                    lValue = .lNoResistAt2X
                Case PlayerSpecialAttributeID.eOfficersEnlistedCost
                    lValue = .yOfficersEnlistedCost
                Case PlayerSpecialAttributeID.eOfficerTrainingFactor
                    lValue = .OfficerTrainingFactor
                Case PlayerSpecialAttributeID.eOtherFacUpkeepReduction
                    lValue = .myOtherFacUpkeepReduct
                Case PlayerSpecialAttributeID.ePayloadTypeAvailable
                    lValue = .yPayloadTypeAvailable
                Case PlayerSpecialAttributeID.ePierceResistImprove
                    lValue = .yPierceResistImprove
                Case PlayerSpecialAttributeID.ePopulationUpkeepReduction
                    lValue = .myPopulationUpkeepReduct
                Case PlayerSpecialAttributeID.ePowerGenProdToWorker
                    lValue = .PowerGenProdToWorker
                Case PlayerSpecialAttributeID.ePrecisionMissiles
                    lValue = .yPrecisionMissiles
                Case PlayerSpecialAttributeID.eProjectileExplosionRadius
                    lValue = .yProjExpRadius
                Case PlayerSpecialAttributeID.eProjectileMaxRng
                    lValue = .iProjMaxRange
                Case PlayerSpecialAttributeID.ePulseBeamMaxDmgImprove
                    lValue = .yPulseMaxDmgBonus
                Case PlayerSpecialAttributeID.ePulseBeamMinDmgImprove
                    lValue = .yPulseMinDmgBonus
                Case PlayerSpecialAttributeID.ePulseBeamOptimumRange
                    lValue = .iPulseOptRng
                Case PlayerSpecialAttributeID.ePulseBeamOptRngImprove
                    lValue = .yPulseOptRngBonus
                Case PlayerSpecialAttributeID.ePulseBeamPowerReduced
                    lValue = .yPulsePowerReduced
                Case PlayerSpecialAttributeID.ePulseBeamROFImprove
                    lValue = .yPulseROFBonus
                Case PlayerSpecialAttributeID.ePulseBeamMaxCompressFactor
                    lValue = .lPulseMaxCompressFactor
                Case PlayerSpecialAttributeID.ePulseBeamInputEnergyMax
                    lValue = .lPulseMaxInputEnergy
                Case PlayerSpecialAttributeID.ePulseBeamWpnROF
                    lValue = .iPulseWpnROF
                Case PlayerSpecialAttributeID.eRadarDisRes
                    lValue = .yRadarDisRes
                Case PlayerSpecialAttributeID.eRadarMaxRange
                    lValue = .yRadarMaxRange
                Case PlayerSpecialAttributeID.eRadarMaxRangeBonus
                    lValue = .yRadarMaxRangeBonus
                Case PlayerSpecialAttributeID.eRadarOptRange
                    lValue = .yRadarOptRange
                Case PlayerSpecialAttributeID.eRadarOptRangeBonus
                    lValue = .yRadarOptRangeBonus
                Case PlayerSpecialAttributeID.eRadarResistImprove
                    lValue = .yRadarResistImprove
                Case PlayerSpecialAttributeID.eRadarScanRes
                    lValue = .yRadarScanRes
                Case PlayerSpecialAttributeID.eRadarScanResBonus
                    lValue = .yRadarScanResBonus
                Case PlayerSpecialAttributeID.eRadarWpnAcc
                    lValue = .yRadarWpnAcc
                Case PlayerSpecialAttributeID.eRadarWpnAccBonus
                    lValue = .yRadarWpnAccBonus
                Case PlayerSpecialAttributeID.eRepairSpeedBonus
                    lValue = .yRepairSpeedBonus
                Case PlayerSpecialAttributeID.eResearchCrewQtrs
                    lValue = .ResearchCrewQtrs
                Case PlayerSpecialAttributeID.eResearchEspionage
                    lValue = .yResearchEspionage
                Case PlayerSpecialAttributeID.eResearchFacAssistBonus
                    lValue = .myResearchFacAssistBonus
                Case PlayerSpecialAttributeID.eResearchFacUpkeepReduction
                    lValue = .myResearchFacUpkeepReduct
                Case PlayerSpecialAttributeID.eResearchProdFactor
                    lValue = .ResearchProdFactor
                Case PlayerSpecialAttributeID.eSellOrderSlots
                    lValue = .ySellOrderSlots
                Case PlayerSpecialAttributeID.eShieldMaxHP
                    lValue = .lShieldMaxHP
                Case PlayerSpecialAttributeID.eShieldMaxHPBonus
                    lValue = .yShieldMaxHPBonus
                Case PlayerSpecialAttributeID.eShieldProjSize
                    lValue = .lShieldProjHullSize
                Case PlayerSpecialAttributeID.eShieldProjSizeBonus
                    lValue = .yShieldProjHullSizeBonus
                Case PlayerSpecialAttributeID.eShieldRechargeFreqDecrease
                    lValue = .yShieldRechargeFreqBonus
                Case PlayerSpecialAttributeID.eShieldRechargeFreqLow
                    lValue = .iShieldRechargeFreq
                Case PlayerSpecialAttributeID.eShieldRechargeRate
                    lValue = .lShieldRechargeRate
                Case PlayerSpecialAttributeID.eShieldRechargeRateBonus
                    lValue = .yShieldRechargeRateBonus
                Case PlayerSpecialAttributeID.eSolidBeamMaxDmgBonus
                    lValue = .ySolidBeamMaxDmgBonus
                Case PlayerSpecialAttributeID.eSolidBeamMinDmgBonus
                    lValue = .ySolidBeamMinDmgBonus
                Case PlayerSpecialAttributeID.eSolidBeamOptimumRange
                    lValue = .iSolidBeamOptimumRange
                Case PlayerSpecialAttributeID.eSolidBeamOptRangeBonus
                    lValue = .ySolidBeamOptRngBonus
                Case PlayerSpecialAttributeID.eSolidBeamPierceRatio
                    lValue = .ySolidBeamPierceRatio
                Case PlayerSpecialAttributeID.eSolidBeamPowerReduced
                    lValue = .ySolidBeamPowerBonus
                Case PlayerSpecialAttributeID.eSolidBeamROFDecrease
                    lValue = .ySolidBeamROFBonus
                Case PlayerSpecialAttributeID.eSolidBeamWeaponMaxAccuracy
                    lValue = .ySolidBeamMaxAccuracy
                Case PlayerSpecialAttributeID.eSolidBeamWpnMaxDmg
                    lValue = .lSolidBeamMaxDmg
                Case PlayerSpecialAttributeID.eSolidBeamWpnROF
                    lValue = .iSolidBeamROF
                Case PlayerSpecialAttributeID.eSpaceportUpkeepReduction
                    lValue = .mySpaceportUpkeepReduct
                Case PlayerSpecialAttributeID.eStudyMineralEffect
                    lValue = .yStudyMineralEffect
                Case PlayerSpecialAttributeID.eTaxRateWithoutPenalty
                    lValue = .yTaxRateBeforePenalty
                Case PlayerSpecialAttributeID.eTradeBoardRange
                    lValue = .yTradeBoardRange
                Case PlayerSpecialAttributeID.eTradeCosts
                    lValue = .yTradeCosts
                Case PlayerSpecialAttributeID.eTradeGTCSpeed
                    lValue = .yTradeGTCSpeed
                Case PlayerSpecialAttributeID.eTransporters
                    lValue = .yTransporters
                Case PlayerSpecialAttributeID.eUnderwaterFacUpkeepReduct
                    lValue = .myUnderwaterUpkeepReduct
                Case PlayerSpecialAttributeID.eUnemploymentUpkeepReduct
                    lValue = .myUnemploymentUpkeepReduct
                Case PlayerSpecialAttributeID.eUnemploymentWithoutPenalty
                    lValue = .yUnemploymentBeforePenalty
            End Select
        End With
        Return lValue
    End Function

    Public Sub SetCPLimit(ByVal iNewLimit As Int16)
        If iCPLimit <= iNewLimit Then
            iCPLimit = iNewLimit
            goMsgSys.SendPlayerCPLimitUpdate(oPlayer.ObjectID, iCPLimit)
        End If
    End Sub

    Public Sub ProcessTechResearched(ByRef oTech As SpecialTech)
        If oTech.ObjectID = 0 Then Return

        Dim lValue As Int32 = oTech.lNewValue

        Select Case CType(oTech.ProgramControl, PlayerSpecialAttributeID)
            Case PlayerSpecialAttributeID.eAerialUpkeepReduction
                myAerialUpkeepReduct = Math.Max(myAerialUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.eAlloyBuilderImprovements
                yAlloyBuilderImprovements = Math.Max(yAlloyBuilderImprovements, CByte(lValue))
            Case PlayerSpecialAttributeID.eAreaEffectJamming
                yJamTargets = CByte(yJamTargets Or 128)
            Case PlayerSpecialAttributeID.eArmorCheaperHullToHP
                yArmorHullToHP = Math.Max(yArmorHullToHP, CByte(lValue))
            Case PlayerSpecialAttributeID.eAutomaticMineralMovement
                If yAutomaticMinMove <> 0 Then
                    yAutomaticMinMove = Math.Min(yAutomaticMinMove, CByte(lValue))
                Else : yAutomaticMinMove = CByte(lValue)
                End If
            Case PlayerSpecialAttributeID.eBaseEnlistedPerTrainer
                BaseEnlistedPerTrainer = Math.Max(BaseEnlistedPerTrainer, CByte(lValue))
            Case PlayerSpecialAttributeID.eBaseOfficerPerTrainer
                BaseOfficerPerTrainer = Math.Max(BaseOfficerPerTrainer, CByte(lValue))
            Case PlayerSpecialAttributeID.eBeamResistImprove
                yBeamResistImprove = Math.Max(yBeamResistImprove, CByte(lValue))
            Case PlayerSpecialAttributeID.eBetterHullToResidence
                HullPerResidence = Math.Min(HullPerResidence, CByte(lValue))
            Case PlayerSpecialAttributeID.eBombMaxAOE
                MaxBombAOE = Math.Max(CByte(lValue), MaxBombAOE)
            Case PlayerSpecialAttributeID.eBombMaxGuidance
                MaxBombGuidance = Math.Max(CByte(lValue), MaxBombGuidance)
            Case PlayerSpecialAttributeID.eBombMaxRange
                MaxBombRange = Math.Max(CByte(lValue), MaxBombRange)
            Case PlayerSpecialAttributeID.eBombMinROF
                MinBombROF = Math.Min(CByte(lValue), MinBombROF)
            Case PlayerSpecialAttributeID.eBombPayloadType
                BombPayloadType = BombPayloadType Or CByte(lValue)
            Case PlayerSpecialAttributeID.eBurnResistImprove
                yBurnResistImprove = Math.Max(yBurnResistImprove, CByte(lValue))
            Case PlayerSpecialAttributeID.eBuyAndSellOrderSlots
                yBuySellTechLvl = Math.Max(yBuySellTechLvl, CByte(lValue))

                yBuyOrderSlots = 1 '0
                ySellOrderSlots = 10 '5
                Select Case yBuySellTechLvl
                    Case 1
                        ySellOrderSlots = 15 '10
                        yBuyOrderSlots = 1 '0
                    Case 2
                        ySellOrderSlots = 20 '15
                        yBuyOrderSlots = 2 '1
                    Case 3
                        ySellOrderSlots = 25 '20
                        yBuyOrderSlots = 3
                    Case 4
                        ySellOrderSlots = 30 '25
                        yBuyOrderSlots = 5
                End Select
            Case PlayerSpecialAttributeID.eBuyOrderSlots
                yBuyOrderSlots = Math.Max(yBuyOrderSlots, CByte(lValue))
            Case PlayerSpecialAttributeID.eChemResistImprove
                yChemResistImprove = Math.Max(yChemResistImprove, CByte(lValue))
            Case PlayerSpecialAttributeID.eColonyMoraleBoost
                yColonyMoraleBoost = Math.Max(yColonyMoraleBoost, CByte(lValue))
            Case PlayerSpecialAttributeID.eCommandEspionageAbility
                yCommandEspionage = Math.Max(yCommandEspionage, CByte(lValue))
            Case PlayerSpecialAttributeID.eConstructionEspionageAbility
                yConstructionEspionage = Math.Max(yConstructionEspionage, CByte(lValue))
            Case PlayerSpecialAttributeID.eControlledGrowth
                yControlledGrowthLimit = Math.Max(yControlledGrowthLimit, CByte(lValue))
            Case PlayerSpecialAttributeID.eControlledMorale
                yControlledMoraleLimit = Math.Max(yControlledMoraleLimit, CByte(lValue))
            Case PlayerSpecialAttributeID.eCPLimit
                SetCPLimit(CShort(lValue))
            Case PlayerSpecialAttributeID.eDiscoverTime
                yDiscoverTime = Math.Max(yDiscoverTime, CByte(lValue))
            Case PlayerSpecialAttributeID.eDiplomaticEspionage
                yDiplomaticEspionage = Math.Max(yDiplomaticEspionage, CByte(lValue))
            Case PlayerSpecialAttributeID.eEngineMaxSpeed
                yMaxSpeed = Math.Max(yMaxSpeed, CByte(lValue))
            Case PlayerSpecialAttributeID.eEnginePowerBonus
                yPowerBonus = Math.Max(yPowerBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eEnginePowerThrustLimit
                lPowerThrustLimit = Math.Max(lPowerThrustLimit, lValue)
            Case PlayerSpecialAttributeID.eEnlistedsColonistsCost
                yEnlistedColonistCost = Math.Min(yEnlistedColonistCost, CByte(lValue))
            Case PlayerSpecialAttributeID.eEnlistedTrainingFactor
                EnlistedTrainingFactor = Math.Min(EnlistedTrainingFactor, CByte(lValue))
            Case PlayerSpecialAttributeID.eFactoryUpkeepReduction
                myFactoryUpkeepReduct = Math.Max(myFactoryUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.eGrowthRateBonus
                yGrowthRateBonus = Math.Max(yGrowthRateBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eHomelessAllowedBeforePenalty
                yHomelessAllowedBeforePenalty = Math.Max(yHomelessAllowedBeforePenalty, CByte(lValue))
            Case PlayerSpecialAttributeID.eImpactResistImprove
                yImpactResistImprove = Math.Max(yImpactResistImprove, CByte(lValue))
            Case PlayerSpecialAttributeID.eJammingEffectAvailable
                yJamEffectAvailable = CByte(yJamEffectAvailable Or lValue)
            Case PlayerSpecialAttributeID.eJammingImmunityAvailable
                yJamImmunityAvailable = CByte(yJamImmunityAvailable Or lValue)
            Case PlayerSpecialAttributeID.eJammingStrength
                yJamStrength = Math.Max(yJamStrength, CByte(lValue))
            Case PlayerSpecialAttributeID.eJammingTargets
                If (yJamTargets And 128) <> 0 Then
                    Dim lTmp As Int32 = yJamTargets - 128
                    lTmp = Math.Max(lTmp, lValue)
                    yJamTargets = CByte(lTmp Or 128)
                Else : yJamTargets = Math.Max(yJamTargets, CByte(lValue))
                End If
            Case PlayerSpecialAttributeID.eLinkChanceBonus
                yMultipleLinkChance = Math.Min(yMultipleLinkChance, CByte(lValue))
            Case PlayerSpecialAttributeID.eMagResistImprove
                yMagneticResistImprove = Math.Max(yMagneticResistImprove, CByte(lValue))
            Case PlayerSpecialAttributeID.eManeuverBonus
                yManeuverBonus = Math.Max(yManeuverBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eManeuverLimit
                yMaxManeuver = Math.Max(yMaxManeuver, CByte(lValue))
            Case PlayerSpecialAttributeID.eMaxAlloyImprovement
                yMaxAlloyImprovement = Math.Max(yMaxAlloyImprovement, CByte(lValue))
            Case PlayerSpecialAttributeID.eMaxBattlegroups
                yMaxBattlegroups = Math.Max(yMaxBattlegroups, CByte(lValue))
            Case PlayerSpecialAttributeID.eMaxBattlegroupUnits
                yMaxBattleGroupUnits = Math.Max(yMaxBattleGroupUnits, CByte(lValue))
            Case PlayerSpecialAttributeID.eMaxnumberofconcurrentLinks
                yMaxLinks = Math.Max(yMaxLinks, CByte(lValue))
            Case PlayerSpecialAttributeID.eMaxSpeedBonus
                yMaxSpeedBonus = Math.Max(yMaxSpeedBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eMineralConcentrationBonus
                yMineralConcentrationBonus = Math.Max(yMineralConcentrationBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eMiningWorkers
                MiningWorkers = Math.Max(MiningWorkers, CByte(lValue))
            Case PlayerSpecialAttributeID.eMissileExplosionRadiusBonus
                yMissileExpRadBonus = Math.Max(yMissileExpRadBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eMissileFlightTimeBonus
                yMissileFlightTimeBonus = Math.Max(yMissileFlightTimeBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eMissileHullSizeImprove
                yMissileHullSizeBonus = Math.Max(yMissileHullSizeBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eMissileManeuverBonus
                yMissileManeuverBonus = Math.Max(yMissileManeuverBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eMissileMaxDmgBonus
                yMissileMaxDmgBonus = Math.Max(yMissileMaxDmgBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eMissileMaxSpeedImprove
                yMissileMaxSpeedBonus = Math.Max(yMissileMaxSpeedBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eNavalAbility
                yNavalAbility = Math.Max(yNavalAbility, CByte(lValue))
            Case PlayerSpecialAttributeID.eNonAerialUpkeepReduction
                myNonAerialUpkeepReduct = Math.Max(myNonAerialUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.eNoResistAt2X
                lNoResistAt2X = lNoResistAt2X Or lValue
            Case PlayerSpecialAttributeID.eOfficersEnlistedCost
                yOfficersEnlistedCost = Math.Max(yOfficersEnlistedCost, CByte(lValue))
            Case PlayerSpecialAttributeID.eOfficerTrainingFactor
                OfficerTrainingFactor = Math.Max(OfficerTrainingFactor, CByte(lValue))
            Case PlayerSpecialAttributeID.eOtherFacUpkeepReduction
                myOtherFacUpkeepReduct = Math.Max(myOtherFacUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.ePayloadTypeAvailable
                yPayloadTypeAvailable = Math.Max(yPayloadTypeAvailable, CByte(lValue))
            Case PlayerSpecialAttributeID.ePierceResistImprove
                yPierceResistImprove = Math.Max(yPierceResistImprove, CByte(lValue))
            Case PlayerSpecialAttributeID.ePopulationUpkeepReduction
                myPopulationUpkeepReduct = Math.Max(myPopulationUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.ePowerGenProdToWorker
                PowerGenProdToWorker = Math.Max(PowerGenProdToWorker, CByte(lValue))
            Case PlayerSpecialAttributeID.ePrecisionMissiles
                yPrecisionMissiles = Math.Max(yPrecisionMissiles, CByte(lValue))
            Case PlayerSpecialAttributeID.eProjectileExplosionRadius
                yProjExpRadius = Math.Max(yProjExpRadius, CByte(lValue))
            Case PlayerSpecialAttributeID.eProjectileMaxRng
                iProjMaxRange = Math.Max(iProjMaxRange, CShort(lValue))
            Case PlayerSpecialAttributeID.ePulseBeamMaxDmgImprove
                yPulseMaxDmgBonus = Math.Max(yPulseMaxDmgBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.ePulseBeamMinDmgImprove
                yPulseMinDmgBonus = Math.Max(yPulseMinDmgBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.ePulseBeamOptimumRange
                iPulseOptRng = Math.Max(iPulseOptRng, CShort(lValue))
            Case PlayerSpecialAttributeID.ePulseBeamOptRngImprove
                yPulseOptRngBonus = Math.Max(yPulseOptRngBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.ePulseBeamPowerReduced
                yPulsePowerReduced = Math.Max(yPulsePowerReduced, CByte(lValue))
            Case PlayerSpecialAttributeID.ePulseBeamROFImprove
                yPulseROFBonus = Math.Max(yPulseROFBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.ePulseBeamMaxCompressFactor
                lPulseMaxCompressFactor = Math.Max(lPulseMaxCompressFactor, lValue)
            Case PlayerSpecialAttributeID.ePulseBeamInputEnergyMax
                lPulseMaxInputEnergy = Math.Max(lPulseMaxInputEnergy, lValue)
            Case PlayerSpecialAttributeID.ePulseBeamWpnROF
                iPulseWpnROF = Math.Min(iPulseWpnROF, CShort(lValue))
            Case PlayerSpecialAttributeID.eRadarDisRes
                yRadarDisRes = Math.Max(yRadarDisRes, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarMaxRange
                yRadarMaxRange = Math.Max(yRadarMaxRange, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarMaxRangeBonus
                yRadarMaxRangeBonus = Math.Max(yRadarMaxRangeBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarOptRange
                yRadarOptRange = Math.Max(yRadarOptRange, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarOptRangeBonus
                yRadarOptRangeBonus = Math.Max(yRadarOptRangeBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarResistImprove
                yRadarResistImprove = Math.Max(yRadarResistImprove, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarScanRes
                yRadarScanRes = Math.Max(yRadarScanRes, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarScanResBonus
                yRadarScanResBonus = Math.Max(yRadarScanResBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarWpnAcc
                yRadarWpnAcc = Math.Max(yRadarWpnAcc, CByte(lValue))
            Case PlayerSpecialAttributeID.eRadarWpnAccBonus
                yRadarWpnAccBonus = Math.Max(yRadarWpnAccBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eRepairSpeedBonus
                yRepairSpeedBonus = Math.Max(yRepairSpeedBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eResearchCrewQtrs
                ResearchCrewQtrs = Math.Min(ResearchCrewQtrs, CByte(lValue))
            Case PlayerSpecialAttributeID.eResearchEspionage
                yResearchEspionage = Math.Max(yResearchEspionage, CByte(lValue))
            Case PlayerSpecialAttributeID.eResearchFacAssistBonus
                myResearchFacAssistBonus = Math.Max(myResearchFacAssistBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eResearchFacUpkeepReduction
                myResearchFacUpkeepReduct = Math.Max(myResearchFacUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.eResearchProdFactor
                ResearchProdFactor = Math.Min(ResearchProdFactor, CByte(lValue))
            Case PlayerSpecialAttributeID.eSellOrderSlots
                ySellOrderSlots = Math.Max(ySellOrderSlots, CByte(lValue))
            Case PlayerSpecialAttributeID.eShieldMaxHP
                lShieldMaxHP = Math.Max(lShieldMaxHP, lValue)
            Case PlayerSpecialAttributeID.eShieldMaxHPBonus
                yShieldMaxHPBonus = Math.Max(yShieldMaxHPBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eShieldProjSize
                lShieldProjHullSize = Math.Max(lShieldProjHullSize, lValue)
            Case PlayerSpecialAttributeID.eShieldProjSizeBonus
                yShieldProjHullSizeBonus = Math.Max(yShieldProjHullSizeBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eShieldRechargeFreqDecrease
                yShieldRechargeFreqBonus = Math.Max(yShieldRechargeFreqBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eShieldRechargeFreqLow
                iShieldRechargeFreq = Math.Min(iShieldRechargeFreq, CShort(lValue))
            Case PlayerSpecialAttributeID.eShieldRechargeRate
                lShieldRechargeRate = Math.Max(lShieldRechargeRate, lValue)
            Case PlayerSpecialAttributeID.eShieldRechargeRateBonus
                yShieldRechargeRateBonus = Math.Max(yShieldRechargeRateBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamMaxDmgBonus
                ySolidBeamMaxDmgBonus = Math.Max(ySolidBeamMaxDmgBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamMinDmgBonus
                ySolidBeamMinDmgBonus = Math.Max(ySolidBeamMinDmgBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamOptimumRange
                iSolidBeamOptimumRange = Math.Max(iSolidBeamOptimumRange, CShort(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamOptRangeBonus
                ySolidBeamOptRngBonus = Math.Max(ySolidBeamOptRngBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamPierceRatio
                ySolidBeamPierceRatio = Math.Max(ySolidBeamPierceRatio, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamPowerReduced
                ySolidBeamPowerBonus = Math.Max(ySolidBeamPowerBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamROFDecrease
                ySolidBeamROFBonus = Math.Max(ySolidBeamROFBonus, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamWeaponMaxAccuracy
                ySolidBeamMaxAccuracy = Math.Max(ySolidBeamMaxAccuracy, CByte(lValue))
            Case PlayerSpecialAttributeID.eSolidBeamWpnMaxDmg
                lSolidBeamMaxDmg = Math.Max(lSolidBeamMaxDmg, lValue)
            Case PlayerSpecialAttributeID.eSolidBeamWpnROF
                iSolidBeamROF = Math.Min(iSolidBeamROF, CShort(lValue))
            Case PlayerSpecialAttributeID.eSpaceportUpkeepReduction
                mySpaceportUpkeepReduct = Math.Max(mySpaceportUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.eStudyMineralEffect
                yStudyMineralEffect = Math.Max(yStudyMineralEffect, CByte(lValue))
            Case PlayerSpecialAttributeID.eSuperSpecials
                lSuperSpecials = CType(lSuperSpecials Or lValue, SuperSpecialID)
            Case PlayerSpecialAttributeID.eTaxRateWithoutPenalty
                yTaxRateBeforePenalty = Math.Max(yTaxRateBeforePenalty, CByte(lValue))
            Case PlayerSpecialAttributeID.eTradeBoardRange
                yTradeBoardRange = Math.Max(yTradeBoardRange, CByte(lValue))
            Case PlayerSpecialAttributeID.eTradeCosts
                yTradeCosts = Math.Min(yTradeCosts, CByte(lValue))
            Case PlayerSpecialAttributeID.eTradeGTCSpeed
                yTradeGTCSpeed = Math.Max(yTradeGTCSpeed, CByte(lValue))
            Case PlayerSpecialAttributeID.eTransporters
                yTransporters = Math.Max(yTransporters, CByte(lValue))
            Case PlayerSpecialAttributeID.eUnderwaterFacUpkeepReduct
                myUnderwaterUpkeepReduct = Math.Max(myUnderwaterUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.eUnemploymentUpkeepReduct
                myUnemploymentUpkeepReduct = Math.Max(myUnemploymentUpkeepReduct, CByte(lValue))
            Case PlayerSpecialAttributeID.eUnemploymentWithoutPenalty
                yUnemploymentBeforePenalty = Math.Max(yUnemploymentBeforePenalty, CByte(lValue))
        End Select

        'TODO: Not sure this is the best idea for this...
        If oPlayer.lConnectedPrimaryID > -1 Then
            Try
                oPlayer.CrossPrimarySafeSendMsg(GetSpecialAttributesMsg)
            Catch ex As Exception
                'do nothing
            End Try
        End If
    End Sub

    Public Function GetSpecialAttributesMsg() As Byte()
        Dim yMsg(11 + (PlayerSpecialAttributeID.eAttributeFinalValue * 4)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerTechValue).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSuperSpecials).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(CShort(PlayerSpecialAttributeID.eAttributeFinalValue)).CopyTo(yMsg, lPos) : lPos += 2
        For X As Int32 = 0 To PlayerSpecialAttributeID.eAttributeFinalValue - 1
            Dim lValue As Int32 = GetPropertyValueByID(Me, X)
            System.BitConverter.GetBytes(lValue).CopyTo(yMsg, lPos) : lPos += 4
        Next X
        Return yMsg
    End Function
End Class