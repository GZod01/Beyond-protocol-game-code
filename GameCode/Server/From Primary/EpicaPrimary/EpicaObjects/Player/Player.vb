#Region "  Enum and Structures  "

Public Enum eEmailSettings As Short
    ePlayerRels = 1         'player rel changes
    eLowResources = 2       'low resource msgs
    eEngagedEnemy = 4       'engaged the enemy
    eUnderAttack = 8        'under attack alerts
    'eBuyOrderAccepted = 16  'Buy Order Accepted
    eAllInternalEmails = 16
    eResearchComplete = 32  'Research Complete
    eTradeRequested = 64
    eRebuildCancel = 128
    eAgentUpdates = 256
    eGuildMembershipNotices = 512
    eMineralOutbid = 1024
    eFacilityLost = 2048
    eColonyLost = 4096
    eTitleChange = 8192
    eFactionUpdates = 16384
End Enum

Public Enum ePlayerSpecialAttributeSetting As Short
    eCelebrationPeriod = 0
    eInPirateSpawn = 1
    eNegativeCashflow = 2
    eDoctrineOfLeadership = 3
    eBadWarDecMoralePenalty = 4
    eBadWarDecCPPenalty = 5
    'TODO: add more here
    ePlayerTitle = 6
End Enum

Public Enum eDoctrineOfLeadershipSetting As Byte
    eNoSetting = 0      'only allowed if the player has never set one
    eDefense = 1        '+2 to Armor Resists
    eFlexibility = 2    '-2% Military Costs
    eMobility = 3       '+1 to Speed And Maneuver
    eSurprise = 4       '+5 Detect Resist
    eUrgency = 5        '+5% Structural HP
End Enum

Public Enum AliasingRights As Int32
    eNoRights = 0
    eMoveUnits = 1
    eDockUndockUnits = 2
    eChangeBehavior = 4
    eAddProduction = 8
    eCancelProduction = 16
    eViewTechDesigns = 32
    eCreateBattleGroups = 64
    eViewBattleGroups = 128
    eModifyBattleGroups = 256
    eViewDiplomacy = 512
    eAlterDiplomacy = 1024
    eViewAgents = 2048
    eAlterAgents = 4096
    eViewColonyStats = 8192
    eAlterColonyStats = 16384
    eViewBudget = 32768
    eViewMining = 65536
    eViewEmail = 131072
    eAlterEmail = 262144
    eViewTreasury = 524288
    eViewTrades = 1048576
    eAlterTrades = 2097152
    eAddResearch = 4194304
    eCancelResearch = 8388608
    eChangeEnvironment = 16777216
    eViewResearch = 33554432
    eViewUnitsAndFacilities = 67108864
    eAlterAutoLaunchPower = 134217728
    eTransferCargo = 268435456
    eCreateDesigns = 536870912
    eDismantle = 1073741824
End Enum

Public Structure AliasLogin
    Public yUserName() As Byte
    Public yPassword() As Byte
    Public lRights As Int32
End Structure

Public Structure ImposedAgentEffect
    Public oTargetPlayer As Player
    Public sSpecificName As String

    Public lStartCycle As Int32     'when this effect began
    Public lDuration As Int32       'in cycles, how long this effect will last
    Public yType As AgentEffectType 'the type of effect this applies
    Public lAmount As Int32         'amount of effect applied
    Public bAmountAsPerc As Boolean 'if true, indicates that lAmount is percentage based (needs to be divided by 100)
End Structure

Public Enum eySlotState As Byte
    ''' <summary>
    ''' No party receives benefits. Appears in list if the list has room. Default setting.
    ''' </summary>
    ''' <remarks></remarks>
    Unaccepted = 0
    ''' <summary>
    ''' Means that the slot configuration has been accepted.
    ''' </summary>
    ''' <remarks></remarks>
    Accepted = 1
    ''' <summary>
    ''' Lesser is at war with someone that the greater is not. The Greater loses the benefits.
    ''' </summary>
    ''' <remarks></remarks>
    WarNotJoined = 2
    ''' <summary>
    ''' Someone's rank is too high/too low for this slot config
    ''' </summary>
    ''' <remarks></remarks>
    RankTooHigh = 4
    ''' <summary>
    ''' Someone's rank is too high/too low for this slot config
    ''' </summary>
    ''' <remarks></remarks>
    InsufficientFactionSlots = 8
    ''' <summary>
    ''' Greater is at war with someone that the Lesser is slotted in (has as a faction). Lesser loses the benefits.
    ''' </summary>
    ''' <remarks></remarks>
    FactionAtWar = 16
    ''' <summary>
    ''' Indicates that one side of the faction is an ex-rank
    ''' </summary>
    ''' <remarks></remarks>
    ExRankMember = 32
    ''' <summary>
    ''' Used to indicate that one party is breaking the slot config forcefully
    ''' </summary>
    ''' <remarks></remarks>
    ForceRemove = 255
End Enum

Public Enum eyPlayerPhase As Byte
    eInitialPhase = 0               'first part of tutorial (newb system 1) just starting
    eSecondPhase = 1                'second part of tutorial (newb system 2)
    eThirdPhase = 2                 'third part of tutorial (newb system 3)
    eFullLivePhase = 255            'player is live
End Enum

Public Enum eIronCurtainState As Byte
    IronCurtainIsDown = 0
    IronCurtainIsUpOnSelectedPlanet = 1
    IronCurtainIsUpOnEverything = 2
    RaisingIronCurtainOnEverything = 3
    RaisingIronCurtainOnSelectedPlanet = 4
    IronCurtainIsFalling = 5
End Enum

Public Enum elCustomRankPermissions As Int32            'what ranks the player is able to assign to their custom value
    Explorer = 1
    Scientist = 2
    MasterScientist = 4
    Inquisitor = 8
    ChiefScientist = 16
    Preeminence = 32
    Transcendent = 64

    Diplomat = 128
    Arbiter = 256
    Counselor = 512
    Senator = 1024
    HighSenator = 2048
    Chancellor = 4096
    SupremeChancellor = 8192

    Trader = 16384
    Merchant = 32768
    Broker = 65536
    TradeLord = 131072
    ChiefBroker = 262144
    MasterMerchant = 524288
    CommerceCzar = 1048576
End Enum
Public Enum eyCustomRank As Byte
    Magistrate = 0
    Governor = 1
    Overseer = 2
    Duke = 3
    Baron = 4
    King = 5
    Emperor = 6
    TraderShift = 16
    DiplomacyShift = 32
    ResearcherShift = 64
End Enum

Public Enum elPlayerStatusFlag As int32
    FullInvulnerabilityRaised = 1
    AlwaysRaiseFullInvul = 2
End Enum

Public Enum elWarpointCheckFlag As int32
    eNoCheckRequired = 0
    eCheckAllNormalUnits = 1
    eCheckAllGuildShare = 2
End Enum

Public Enum eySpawnSettings As Byte
    eNoSetting = 0          'no setting flags have changed
    eOfferedSpawn = 1       'the player was offered to spawn in a spawn system
    eSpawnAccept = 2        'the player responded to the offer with an Accept spawn
    eSpawnRefuse = 4        'the player responded to the offer with a Refuse spawn
    eSpawnProcessOver = 8   'the entire spawn process is over, the actions of the spawn system have been taken based on the decisions defined above
End Enum
#End Region

Public Class Player
    Inherits Epica_GUID

    Public Enum PlayerRank As Byte
        Magistrate = 0
        Governor = 1
        Overseer = 2
        Duke = 3
        Baron = 4
        King = 5
        Emperor = 6
        ExRankShift = 128
    End Enum

    Public AccountStatus As Int32 = AccountStatusType.eInactiveAccount

#Region "  For Cross Primary Connection Management  "
    Public lConnectedPrimaryID As Int32 = -1        '-1 indicates the player is not connected to any primary at present
#End Region

    Public LastLogin As Date
    Public dtAccountWentInactive As Date = Date.MinValue
    Private mlTotalPlayTime As Int32 = 0
    Private mdtLastLoginPlayTimeStore As Date
     
    Public Property TotalPlayTime() As Int32
        Get
            Return mlTotalPlayTime
        End Get
        Set(ByVal value As Int32)
            If LastLogin = mdtLastLoginPlayTimeStore AndAlso gbServerInitializing = False Then
                Return
            End If
            If value - mlTotalPlayTime > 28800 AndAlso gbServerInitializing = False Then
                'Return
                LogEvent(LogEventType.Warning, "Abnormal Play Time Adjustment: " & Me.sPlayerName & " (" & Me.ObjectID & "). Difference: " & (value - mlTotalPlayTime) & ". LastLogin: " & LastLogin)
            End If
            mdtLastLoginPlayTimeStore = LastLogin
            mlTotalPlayTime = value
        End Set
    End Property

    Public ServerIndex As Int32 = -1    'index in the global array

    Public PlayerName(19) As Byte         'displayed to other players
    Public EmpireName(19) As Byte         'displayed to other players
    Public RaceName(19) As Byte           'displayed to other players

#Region "  Alias System Lookups  "
    'Aliases that I am allowing to other players
    Public oAliases() As Player
    Public lAliasIdx() As Int32
    Public uAliasLogin() As AliasLogin
    Public lAliasUB As Int32 = -1

    'Aliases that I am allowed to use
    Public oAllowances() As Player
    Public lAllowanceIdx() As Int32
    Public uAllowanceLogin() As AliasLogin
    Public lAllowanceUB As Int32 = -1

    'Indicates the ID that the player is aliasing
    Public lAliasingPlayerID As Int32 = -1

    Public ReadOnly Property HasOnlineAliases(ByVal lRequiredRight As AliasingRights) As Boolean
        Get
            For X As Int32 = 0 To lAliasUB
                If lAliasIdx(X) <> -1 AndAlso (lRequiredRight = 0 OrElse (uAliasLogin(X).lRights And lRequiredRight) <> 0) AndAlso oAliases(X) Is Nothing = False AndAlso oAliases(X).lAliasingPlayerID = Me.ObjectID AndAlso oAliases(X).lConnectedPrimaryID > -1 Then Return True
            Next X
            Return False
        End Get
    End Property

    Public Sub AddPlayerAlias(ByRef oPlayer As Player, ByVal yUserName() As Byte, ByVal yPassword() As Byte, ByVal lRights As Int32)
        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1

        'Ok, now, find a place in that for us
        For X As Int32 = 0 To lAliasUB
            If lAliasIdx(X) = oPlayer.ObjectID Then
                lIdx = X
                Exit For
            ElseIf lFirstIdx = -1 AndAlso lAliasIdx(X) = -1 Then
                lFirstIdx = X
            End If
        Next X
        If lIdx = -1 Then
            If lFirstIdx = -1 Then
                lIdx = lAliasUB + 1
                ReDim Preserve oAliases(lIdx)
                ReDim Preserve lAliasIdx(lIdx)
                ReDim Preserve uAliasLogin(lIdx)
                lAliasUB += 1
            Else : lIdx = lFirstIdx
            End If
        End If
        oAliases(lIdx) = oPlayer
        lAliasIdx(lIdx) = oPlayer.ObjectID
        With uAliasLogin(lIdx)
            ReDim .yUserName(19)
            ReDim .yPassword(19)
            Array.Copy(yUserName, 0, .yUserName, 0, 20)
            Array.Copy(yPassword, 0, .yPassword, 0, 20)
            .lRights = lRights
        End With

        SaveAliasConfig(lIdx)
    End Sub
    Public Sub AddPlayerAliasAllowance(ByRef oPlayer As Player, ByVal yUserName() As Byte, ByVal yPassword() As Byte, ByVal lRights As Int32)
        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1

        'Ok, now, find a place in that for us
        For X As Int32 = 0 To lAllowanceUB
            If lAllowanceIdx(X) = oPlayer.ObjectID Then
                lIdx = X
                Exit For
            ElseIf lFirstIdx = -1 AndAlso lAllowanceIdx(X) = -1 Then
                lFirstIdx = X
            End If
        Next X
        If lIdx = -1 Then
            If lFirstIdx = -1 Then
                lIdx = lAllowanceUB + 1
                ReDim Preserve oAllowances(lIdx)
                ReDim Preserve lAllowanceIdx(lIdx)
                ReDim Preserve uAllowanceLogin(lIdx)
                lAllowanceUB += 1
            Else : lIdx = lFirstIdx
            End If
        End If
        oAllowances(lIdx) = oPlayer
        lAllowanceIdx(lIdx) = oPlayer.ObjectID
        With uAllowanceLogin(lIdx)
            ReDim .yUserName(19)
            ReDim .yPassword(19)
            Array.Copy(yUserName, 0, .yUserName, 0, 20)
            Array.Copy(yPassword, 0, .yPassword, 0, 20)
            .lRights = lRights
        End With
    End Sub

    Public Function GetAliasMissionList() As String
        Dim oSB As New System.Text.StringBuilder
        Dim bAliasHeader As Boolean = False
        Dim bAllowHeader As Boolean = False

        For X As Int32 = 0 To lAliasUB
            If lAliasIdx(X) > -1 Then
                If oAliases(X) Is Nothing = False Then
                    If bAliasHeader = False Then
                        bAliasHeader = True
                        oSB.AppendLine("The following players have alias accounts available to them for " & Me.sPlayerNameProper)
                    End If
                    oSB.AppendLine("  " & oAliases(X).sPlayerNameProper)
                End If
            End If
        Next X
        For X As Int32 = 0 To lAllowanceUB
            If lAllowanceIdx(X) > -1 Then
                If oAllowances(X) Is Nothing = False Then
                    If bAllowHeader = False Then
                        bAllowHeader = True
                        If bAliasHeader = True Then oSB.AppendLine()
                        oSB.AppendLine("The following players have set up alias accounts for " & Me.sPlayerNameProper)
                    End If
                    oSB.AppendLine("  " & oAllowances(X).sPlayerNameProper)
                End If
            End If
        Next X

        If bAliasHeader = False AndAlso bAllowHeader = False Then
            Return Me.sPlayerNameProper & " has no alias configurations available and has not set up any alias accounts for other players."
        Else
            Return oSB.ToString
        End If
    End Function
#End Region

    Public ExternalEmailAddress(254) As Byte        'sent to email srvr
    Public iEmailSettings As Int16 = eEmailSettings.eEngagedEnemy Or eEmailSettings.eLowResources Or eEmailSettings.ePlayerRels Or eEmailSettings.eUnderAttack
    Public iInternalEmailSettings As Int16 = eEmailSettings.ePlayerRels Or eEmailSettings.eLowResources Or eEmailSettings.eEngagedEnemy Or _
     eEmailSettings.eUnderAttack Or eEmailSettings.eResearchComplete Or eEmailSettings.eTradeRequested Or _
     eEmailSettings.eRebuildCancel Or eEmailSettings.eAgentUpdates Or eEmailSettings.eGuildMembershipNotices Or eEmailSettings.eMineralOutbid Or _
     eEmailSettings.eFacilityLost Or eEmailSettings.eColonyLost Or eEmailSettings.eTitleChange Or eEmailSettings.eFactionUpdates

    Public sPlayerName As String
    Public sPlayerNameProper As String
    Public lPlayerIcon As Int32

    'NOTE: THE NEXT TWO ITEMS (USERNAME AND PASSWORD) ARE TO NEVER BE TRANSMITTED!!!!
    Public PlayerUserName(19) As Byte
    Public PlayerPassword(19) As Byte

    'Public oSenate As Senate			'the senate to which this player belongs to
    Public CommEncryptLevel As Short
    Public EmpireTaxRate As Byte

    Public oPlayerMinerals() As PlayerMineral
    Public lPlayerMineralUB As Int32 = -1

    Public moMinProperties() As PlayerMineralProperty
    Public mlMinPropertyUB As Int32 = -1

    Public BaseMorale As Byte           'the base morale for this player's people
    Public BaseGrowthRate As Byte = 30

    'For bad war declarations
    Public BadWarDecMoralePenalty As Int32 = 0
    Public BadWarDecMoralePenaltyEndCycle As Int32 = 0
    Public BadWarDecCPIncrease As Int32 = 0
    Public BadWarDecCPIncreaseEndCycle As Int32 = 0
    'End of bad war declarations

    Public blMaxPopulation As Int64 = 0
    Public blDBPopulation As Int64 = 0
    Public blCurrentPopulation As Int64 = 0

    Public bInPlayerRequestDetails As Boolean = False
    Public oSocket As NetSock           'the socket this user is connected on
    Public lAlreadyInUseCnt As Int32 = 0

    Public lLastViewedEnvir As Int32
    Public iLastViewedEnvirType As Int16
    Public lStartedEnvirID As Int32
    Public iStartedEnvirTypeID As Int16
    Public lStartLocX As Int32 = Int32.MinValue
    Public lStartLocZ As Int32 = Int32.MinValue

    'Public blCredits As Int64			 'credits the player currently has
    Private mblCredits As Int64 = 0
    Public blCrossPrimaryAdjustCredits As Int64 = 0
    Public Property blCredits() As Int64
        Get
            Return mblCredits
        End Get
        Set(ByVal value As Int64)
            If InMyDomain = False Then
                Dim blDiff As Int64 = value - mblCredits
                blCrossPrimaryAdjustCredits += blDiff
            End If
            mblCredits = value
        End Set
    End Property
    Public lWarSentiment As Int32 = 0

    Public lLastAvailableResourcesRequest As Int32 = Int32.MinValue     'Not saved to DB
    Public lLastUniversalAssetRequest As Int32 = Int32.MinValue

    Public lHangarManMult As Int32 = 1

    Public DeathBudgetBalance As Int32 = 0
    Public DeathBudgetEndTime As Int32 = Int32.MinValue
    Public DeathBudgetFundsRemaining As Int32 = 0

    'Always available
    Public oBudget As Budget = New Budget()
    Public oSecurity As PlayerSecurity

    Public PirateStartLocX As Int32 = Int32.MinValue
    Public PirateStartLocZ As Int32 = Int32.MinValue

    'MSC - 1/1/9 - added so that setting the iron curtain planet sets the iron curtain system
    'Public lIronCurtainPlanet As Int32 = -1
    Private mlIronCurtainPlanet As Int32 = -1
    Public Property lIronCurtainPlanet() As Int32
        Get
            Return mlIronCurtainPlanet
        End Get
        Set(ByVal value As Int32)
            mlIronCurtainPlanet = value
            mlIronCurtainSystem = -1
        End Set
    End Property
    Private mlIronCurtainSystem As Int32 = -1
    Public ReadOnly Property lIronCurtainSystem() As Int32
        Get
            If mlIronCurtainSystem = -1 Then
                If mlIronCurtainPlanet > -1 Then
                    Dim oPlanet As Planet = GetEpicaPlanet(mlIronCurtainPlanet)
                    If oPlanet Is Nothing = False Then
                        If oPlanet.ParentSystem Is Nothing = False Then mlIronCurtainSystem = oPlanet.ParentSystem.ObjectID
                    End If
                End If
            End If
            Return mlIronCurtainSystem
        End Get
    End Property



    Public lCurrent15MinutesRemaining As Int32 = 27000
    Public lNext15MinuteReset As Int32 = 0
    Public bInFullLockDown As Boolean = False
    Public yIronCurtainState As eIronCurtainState = eIronCurtainState.IronCurtainIsUpOnSelectedPlanet

    Public yPlayerPhase As eyPlayerPhase = 0
    Public lTutorialStep As Int32 = 1
    Public lStatusFlags As Int32 = 0

#Region "  DX Diag Recording  "
    Public sDXDiag As String = ""
    Public bDXDiagRequested As Boolean = False
    Public Sub SaveDXDiag()
        If sDXDiag = "" Then Return
        If sDXDiag = "SAVED" Then Return
        If sDXDiag.StartsWith("READY") = False Then Return

        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing
        Try
            'remove the "Ready" from the start of the dxdiag
            sDXDiag = sDXDiag.Substring(5)

            'First, try to update
            sSQL = "UPDATE tblPlayerDXDiag SET DXDiag = '" & MakeDBStr(sDXDiag) & "' WHERE PlayerID = " & Me.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                'Ok, do an insert
                sSQL = "INSERT INTO tblPlayerDXDiag (PlayerID, DXDiag) VALUES (" & Me.ObjectID & ", '" & MakeDBStr(sDXDiag) & "')"
                oComm = Nothing
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
            End If
        Catch ex As Exception
            Dim sFile As String = "DXD" & Me.ObjectID.ToString & ".txt"
            LogEvent(LogEventType.CriticalError, "Player.SaveDXDiag(): " & ex.Message & "... Saving file " & sFile)
            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            Try
                Dim oFS As New IO.FileStream(sPath & sFile, IO.FileMode.Create)
                Dim oWriter As New IO.StreamWriter(oFS)
                oWriter.Write(sDXDiag)
                oWriter.Close()
                oWriter.Dispose()
                oFS.Close()
                oFS.Dispose()
            Catch
            End Try
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

        sDXDiag = "SAVED"
        Return
    End Sub
#End Region

    Private mbPlayerIsDead As Boolean = False
    Public Property PlayerIsDead() As Boolean
        Get
            Return mbPlayerIsDead
        End Get
        Set(ByVal value As Boolean)
            If value <> mbPlayerIsDead Then
                mbPlayerIsDead = value
                If lConnectedPrimaryID > -1 Then
                    Dim yMsg(2) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yMsg, 0)
                    If mbPlayerIsDead = True Then yMsg(2) = 1 Else yMsg(2) = 0
                    CrossPrimarySafeSendMsg(yMsg)
                End If
            End If
        End Set
    End Property
    Public yRespawnWithGuild As Byte = 0
    'Public bForcedRestart As Boolean = False

    Private mySendString() As Byte

#Region "  Special Properties  "
    Public lCelebrationEnds As Int32 = 0
    Private mbInNegativeCashflow As Boolean = False
    Public Property bInNegativeCashflow() As Boolean
        Get
            Return mbInNegativeCashflow
        End Get
        Set(ByVal value As Boolean)
            If mbInNegativeCashflow <> value Then
                mbInNegativeCashflow = value
                Dim lVal As Int32 = 0
                If mbInNegativeCashflow = True Then
                    lVal = 1

                    LogEvent(LogEventType.ExtensiveLogging, "PlayerNegCash: " & Me.ObjectID & " @ " & oBudget.GetCashFlow().ToString)
                End If
                Me.CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eNegativeCashflow, lVal)
            End If
        End Set
    End Property
    Private mbInPirateSpawn As Boolean = False
    Public Property bInPirateSpawn() As Boolean
        Get
            Return mbInPirateSpawn
        End Get
        Set(ByVal value As Boolean)
            If mbInPirateSpawn <> value Then
                mbInPirateSpawn = value
                Dim lVal As Int32 = 0
                If mbInPirateSpawn = True Then lVal = 1
                Me.CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eInPirateSpawn, lVal)
            End If
        End Set
    End Property
    Public yCurrentDoctrine As Byte = eDoctrineOfLeadershipSetting.eNoSetting

    Public Claimables(-1) As Claimable
#End Region

    'Public PlayedTimeWhenTimerStarted As Int32 = Int32.MinValue         'for the tutorial phase 2, when the four hour timer started, what was the player's play time?
    Public PlayedTimeWhenFirstWave As Int32 = Int32.MinValue
    Public PlayedTimeAtEndOfWaves As Int32 = Int32.MinValue
    Public PlayedTimeInTutorialOne As Int32 = Int32.MinValue
    Public TutorialPhaseWaves As Int32 = 0
    Public PhaseOneMoneyClicks As Int32 = 0

    Public dtLastRespawn As Date = Date.MinValue
    Public bDeclaredWarOn As Boolean = False
    Public bSpawnAtSameLoc As Boolean = False

    Public dtLastGuildMembership As Date = Date.MinValue

    Public yGender As Byte = 0              '0 = Unsure, 1 = male, 2 = female

    Public bSentMicroExplosionInitial As Boolean = False
    Public ySpawnSystemSetting As Byte = 0      'see eySpawnSettings enum

#Region "  Guild Stuff  "
    Public lGuildID As Int32 = -1
    Public lGuildRankID As Int32 = -1
    Public lJoinedGuildOn As Int32 = 0
    Private moGuild As Guild = Nothing

    Public lInvitedToJoins() As Int32
    Public lInvitedToJoinUB As Int32 = -1
    Public Sub AddInvitedToJoin(ByVal lGuildID As Int32)
        If lInvitedToJoins Is Nothing Then ReDim lInvitedToJoins(-1)
        Dim lCurUB As Int32 = Math.Min(lInvitedToJoinUB, lInvitedToJoins.GetUpperBound(0))
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCurUB
            If lInvitedToJoins(X) = lGuildID Then
                Return
            ElseIf lInvitedToJoins(X) = -1 AndAlso lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            SyncLock lInvitedToJoins
                ReDim Preserve lInvitedToJoins(lInvitedToJoinUB + 1)
                lInvitedToJoins(lInvitedToJoinUB + 1) = lGuildID
                lInvitedToJoinUB += 1
            End SyncLock
        Else
            lInvitedToJoins(lIdx) = lGuildID
        End If
    End Sub
    Public Sub RemoveInvitation(ByVal lGuildID As Int32)
        If lInvitedToJoins Is Nothing Then ReDim lInvitedToJoins(-1)
        Dim lCurUB As Int32 = Math.Min(lInvitedToJoinUB, lInvitedToJoins.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If lInvitedToJoins(X) = lGuildID Then
                lInvitedToJoins(X) = -1
            End If
        Next X
    End Sub

    Public oFormingGuild As Guild = Nothing         'used only if the player is forming a guild

    Public ReadOnly Property oGuild() As Guild
        Get
            If moGuild Is Nothing OrElse moGuild.ObjectID <> lGuildID Then
                moGuild = Nothing
                If lGuildID <> -1 Then moGuild = GetEpicaGuild(lGuildID)
            End If
            Return moGuild
        End Get
    End Property
#End Region

#Region "  Fast Colony Lookup Management  "
    Public mlColonyUB As Int32 = -1
    Public mlColonyIdx() As Int32
    Public mlColonyID() As Int32        'for verification

    'Public mlSpaceFacs(-1) As Int32
    'Public mlSpaceFacUB As Int32 = -1
    'Public Sub AddSpaceFacIdx(ByVal lIdx As Int32)
    '    Dim lFirst As Int32 = -1
    '    For X As Int32 = 0 To mlSpaceFacUB
    '        If mlSpaceFacs(X) = lIdx Then Return
    '        If lFirst = -1 AndAlso mlSpaceFacs(X) = -1 Then
    '            lFirst = X
    '        End If
    '    Next X
    '    If lFirst = -1 Then
    '        mlSpaceFacUB += 1
    '        ReDim Preserve mlSpaceFacs(mlSpaceFacUB)
    '        lFirst = mlSpaceFacUB
    '    End If
    '    mlSpaceFacs(lFirst) = lIdx
    'End Sub
    'Public Sub RemoveSpaceFacIdx(ByVal lIdx As Int32)
    '    For X As Int32 = 0 To mlSpaceFacUB
    '        If mlSpaceFacs(X) = lIdx Then
    '            mlSpaceFacs(X) = -1
    '            Exit For
    '        End If
    '    Next X
    'End Sub

    Public Sub AddColonyIndex(ByVal lColonyGlobalArrayIdx As Int32)
        Dim lIdx As Int32 = -1

        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) = lColonyGlobalArrayIdx Then
                'Already here so return
                PlayerIsDead = False
                mlColonyID(X) = glColonyIdx(lColonyGlobalArrayIdx)
                Return
            ElseIf (mlColonyIdx(X) = -1) AndAlso lIdx = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            mlColonyUB += 1
            ReDim Preserve mlColonyIdx(mlColonyUB)
            ReDim Preserve mlColonyID(mlColonyUB)
            lIdx = mlColonyUB
        End If
        mlColonyIdx(lIdx) = lColonyGlobalArrayIdx
        mlColonyID(lIdx) = glColonyIdx(lColonyGlobalArrayIdx)

        PlayerIsDead = False
    End Sub

    Public Function GetColonyFromParent(ByVal lParentID As Int32, ByVal iParentTypeID As Int16) As Int32
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    With CType(oColony.ParentObject, Epica_GUID)
                        If .ObjectID = lParentID AndAlso .ObjTypeID = iParentTypeID Then Return mlColonyIdx(X)
                    End With
                End If
            End If
        Next X
        Return -1
    End Function

    Public Sub RemoveFastColonyLookup(ByVal lColonyID As Int32)
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = lColonyID Then
                mlColonyIdx(X) = -1
                mlColonyID(X) = -1
                Exit For
            End If
        Next X
    End Sub

    Public Sub HandleCheckForPlayerDeath()
        'Ok, let's check our colony list
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    For Y As Int32 = 0 To oColony.ChildrenUB
                        If oColony.lChildrenIdx(Y) <> -1 Then Return
                    Next Y
                End If
            End If
        Next X

        'Ok, if we are here, then release all captured agents
        If Me.oSecurity Is Nothing = False Then
            For X As Int32 = 0 To Me.oSecurity.lCapturedAgentUB
                If Me.oSecurity.lCapturedAgentIdx(X) > -1 Then
                    Dim oAgent As Agent = oSecurity.oCapturedAgents(X)
                    If oAgent Is Nothing = False Then
                        'Agent is no longer captured
                        If (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then oAgent.lAgentStatus = oAgent.lAgentStatus Xor AgentStatus.HasBeenCaptured
                        'Agent is a fugitive (suspicion = 200)
                        oAgent.Suspicion = 200
                        'Agent Infiltration = 0
                        oAgent.InfiltrationLevel = 0

                        oAgent.yHealth = 100
                        oAgent.lCapturedBy = -1
                        oAgent.lCapturedOn = -1
                        oAgent.lPrisonTestCycles = 0

                        oAgent.lInterrogatorID = -1
                        oAgent.yInterrogationState = 0
                        goAgentEngine.CancelAllAgentEvents(oAgent.ObjectID)
                        oAgent.ReturnHome()

                        Dim yMsg(6) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eCaptureKillAgent).CopyTo(yMsg, 0)
                        System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yMsg, 2)
                        yMsg(6) = 3
                        Me.SendPlayerMessage(yMsg, False, AliasingRights.eViewAgents)


                        'Capturer is notified
                        If (iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                            Dim sAgent As String = BytesToString(oAgent.AgentName)
                            Dim sIntDeath As String = ""
                            
                            Dim oPC As PlayerComm = AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                               sAgent & " has escaped captivity. Somehow the fugitive broke through security and moved outside of the confines of the compound, likely during the destruction of the compound." & _
                               vbCrLf & vbCrLf & "Authorities have been alerted and are in pursuit." & sIntDeath, "Prisoner Escape", Me.ObjectID, _
                               GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If

                        'Return Home call above already sends this... right?
                        'oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                    End If
 
                End If
            Next X
        End If

        'Ok, if we are here, the player is dead.
        PlayerIsDead = True
        'Dim yMsg(2) As Byte
        'System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yMsg, 0)
        'yMsg(2) = 1
        'If Me.oSocket Is Nothing = False Then oSocket.SendData(yMsg)
    End Sub

    Public Function GetColonyFromColonyID(ByVal lID As Int32) As Int32
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = lID Then
                Return mlColonyIdx(X)
            End If
        Next X
        Return -1
    End Function

    Public Function GetClosestSpaceColonyToLoc(ByVal lSystemID As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32) As Int32
        Dim lCurUB As Int32 = Math.Min(mlColonyUB, mlColonyIdx.GetUpperBound(0))
        Dim fMinDist As Single = 50000
        Dim lCurrentID As Int32 = -1

        For X As Int32 = 0 To lCurUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    If CType(oColony.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.eFacility Then Continue For
                    Dim oFac As Facility = CType(oColony.ParentObject, Facility)
                    Dim oParent As Epica_GUID = CType(oFac.ParentObject, Epica_GUID)
                    If oParent Is Nothing = False AndAlso oParent.ObjTypeID = ObjectType.eSolarSystem AndAlso oParent.ObjectID = lSystemID Then
                        Try
                            Dim fDX As Single = oFac.LocX - lLocX
                            Dim fDZ As Single = oFac.LocZ - lLocZ

                            fDX *= fDX
                            fDZ *= fDZ

                            Dim fDist As Single = CSng(Math.Sqrt(fDX + fDZ))
                            If fDist < fMinDist Then
                                fDist = fMinDist
                                lCurrentID = mlColonyIdx(X)
                            End If
                        Catch
                        End Try
                    End If
                End If
            End If
        Next X
        Return lCurrentID
    End Function

    Public Function GetRandomColonyID() As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                lCnt += 1
            End If
        Next X
        lCnt = CInt(Math.Floor(Rnd() * lCnt))

        Dim lLastValid As Int32 = -1
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                lCnt -= 1
                If lCnt < 1 Then Return mlColonyID(X)
                lLastValid = mlColonyID(X)
            End If
        Next X

        Return lLastValid
    End Function

    Public Function GetColonyCount() As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 Then lCnt += 1
        Next X
        Return lCnt
    End Function

    Public Function GetControlledPlanetList() As Planet()
        Dim oResult(-1) As Planet

        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    If CType(oColony.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                        Dim oPlanet As Planet = CType(oColony.ParentObject, Planet)
                        If oPlanet.OwnerID = Me.ObjectID Then
                            ReDim Preserve oResult(oResult.GetUpperBound(0) + 1)
                            oResult(oResult.GetUpperBound(0)) = oPlanet
                        End If
                    End If
                End If
            End If
        Next X
        Return oResult
    End Function

    Public Function GetAllColonyCargoForMineral(ByVal lMineralID As Int32) As Int64
        Dim blTotal As Int64 = 0
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) > -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = Me.ObjectID Then
                    For Y As Int32 = 0 To oColony.mlMineralCacheUB
                        If oColony.mlMineralCacheMineralID(Y) = lMineralID Then
                            If oColony.mlMineralCacheIdx(Y) > -1 AndAlso oColony.mlMineralCacheID(Y) = glMineralCacheIdx(oColony.mlMineralCacheIdx(Y)) Then
                                Dim oCache As MineralCache = goMineralCache(oColony.mlMineralCacheIdx(Y))
                                If oCache Is Nothing = False AndAlso oCache.oMineral Is Nothing = False AndAlso oCache.oMineral.ObjectID = lMineralID Then
                                    blTotal += CLng(oCache.Quantity)
                                    Exit For
                                End If
                            End If
                        End If
                    Next Y
                End If
            End If
        Next X
        Return blTotal
    End Function

    Public Function GetColonyWithRoom(ByVal lAmt As Int32) As Colony
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) > -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    If oColony.TotalCargoCapAvailable > lAmt Then Return oColony
                End If
            End If
        Next X
        Return Nothing
    End Function

    Public Sub UnlockRespawnSystems()
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) > -1 Then
                If glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                    Dim oColony As Colony = goColony(mlColonyIdx(X))
                    If oColony Is Nothing = False Then
                        If oColony.Owner Is Nothing = False Then
                            If oColony.Owner.ObjectID = Me.ObjectID Then
                                'Now, get the colony's parent
                                Dim oParent As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                                If oParent Is Nothing = False Then

                                    If oParent.ObjTypeID = ObjectType.eFacility Then
                                        oParent = CType(CType(oParent, Facility).ParentObject, Epica_GUID)
                                        If oParent Is Nothing Then Continue For
                                    End If

                                    Dim oSys As SolarSystem = Nothing
                                    If oParent.ObjTypeID = ObjectType.ePlanet Then
                                        oSys = CType(oParent, Planet).ParentSystem
                                    ElseIf oParent.ObjTypeID = ObjectType.eSolarSystem Then
                                        oSys = CType(oParent, SolarSystem)
                                    End If
                                    If oSys Is Nothing = False Then
                                        If oSys.SystemType = SolarSystem.elSystemType.RespawnSystem Then
                                            oSys.SystemType = CByte(SolarSystem.elSystemType.UnlockedSystem)
                                            oSys.DataChanged()
                                            gbGalaxyMsgReady = False
                                            oSys.SaveObject()
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next X
    End Sub
#End Region

#Region "  Special Tech Bonuses  "
    Public oSpecials As New Player_Specials()
    Public SpecTechCostMult As Single = 1.0F
    Public GuaranteedSpecialTechID As Int32 = -1
#End Region

#Region "  Technology Management  "
    Public moAlloy() As AlloyTech
    Public mlAlloyIdx() As Int32
    Public mlAlloyUB As Int32 = -1
    Public moArmor() As ArmorTech
    Public mlArmorIdx() As Int32
    Public mlArmorUB As Int32 = -1
    Public moEngine() As EngineTech
    Public mlEngineIdx() As Int32
    Public mlEngineUB As Int32 = -1
    Public moHull() As HullTech
    Public mlHullIdx() As Int32
    Public mlHullUB As Int32 = -1
    Public moPrototype() As Prototype
    Public mlPrototypeIdx() As Int32
    Public mlPrototypeUB As Int32 = -1
    Public moRadar() As RadarTech
    Public mlRadarIdx() As Int32
    Public mlRadarUB As Int32 = -1
    Public moShield() As ShieldTech
    Public mlShieldIdx() As Int32
    Public mlShieldUB As Int32 = -1
    Public moWeapon() As BaseWeaponTech
    Public mlWeaponIdx() As Int32
    Public mlWeaponUB As Int32 = -1

    'Public moSpecialTech() As PlayerSpecialTech
    'Public mlSpecialTechIdx() As Int32
    'Public mlSpecialTechUB As Int32 = -1

    Public Sub RemoveTech(ByVal lID As Int32, ByVal iTypeID As Int16)
        Select Case iTypeID
            Case ObjectType.eAlloyTech
                For X As Int32 = 0 To mlAlloyUB
                    If mlAlloyIdx(X) = lID Then
                        mlAlloyIdx(X) = -1
                        moAlloy(X) = Nothing
                    End If
                Next X
            Case ObjectType.eArmorTech
                For X As Int32 = 0 To mlArmorUB
                    If mlArmorIdx(X) = lID Then
                        mlArmorIdx(X) = -1
                        moArmor(X) = Nothing
                    End If
                Next X
            Case ObjectType.eEngineTech
                For X As Int32 = 0 To mlEngineUB
                    If mlEngineIdx(X) = lID Then
                        mlEngineIdx(X) = -1
                        moEngine(X) = Nothing
                    End If
                Next X
            Case ObjectType.eHullTech
                For X As Int32 = 0 To mlHullUB
                    If mlHullIdx(X) = lID Then
                        mlHullIdx(X) = -1
                        moHull(X) = Nothing
                    End If
                Next X
            Case ObjectType.ePrototype
                For X As Int32 = 0 To mlPrototypeUB
                    If mlPrototypeIdx(X) = lID Then
                        mlPrototypeIdx(X) = -1
                        moPrototype(X) = Nothing
                    End If
                Next X
            Case ObjectType.eRadarTech
                For X As Int32 = 0 To mlRadarUB
                    If mlRadarIdx(X) = lID Then
                        mlRadarIdx(X) = -1
                        moRadar(X) = Nothing
                    End If
                Next X
            Case ObjectType.eShieldTech
                For X As Int32 = 0 To mlShieldUB
                    If mlShieldIdx(X) = lID Then
                        mlShieldIdx(X) = -1
                        moShield(X) = Nothing
                    End If
                Next X
            Case ObjectType.eSpecialTech
                For X As Int32 = 0 To oSpecials.mlSpecialTechUB
                    If oSpecials.mlSpecialTechIdx(X) = lID Then
                        oSpecials.mlSpecialTechIdx(X) = -1
                        oSpecials.moSpecialTech(X) = Nothing
                    End If
                Next X
            Case ObjectType.eWeaponTech
                For X As Int32 = 0 To mlWeaponUB
                    If mlWeaponIdx(X) = lID Then
                        mlWeaponIdx(X) = -1
                        moWeapon(X) = Nothing
                    End If
                Next X
        End Select
    End Sub

    Public Function GetTech(ByVal lID As Int32, ByVal iTypeID As Int16) As Epica_Tech
        Select Case iTypeID
            Case ObjectType.eAlloyTech
                For X As Int32 = 0 To mlAlloyUB
                    If mlAlloyIdx(X) = lID Then
                        Return moAlloy(X)
                    End If
                Next X
            Case ObjectType.eArmorTech
                For X As Int32 = 0 To mlArmorUB
                    If mlArmorIdx(X) = lID Then
                        Return moArmor(X)
                    End If
                Next X
            Case ObjectType.eEngineTech
                For X As Int32 = 0 To mlEngineUB
                    If mlEngineIdx(X) = lID Then
                        Return moEngine(X)
                    End If
                Next X
            Case ObjectType.eHullTech
                For X As Int32 = 0 To mlHullUB
                    If mlHullIdx(X) = lID Then
                        Return moHull(X)
                    End If
                Next X
            Case ObjectType.ePrototype
                For X As Int32 = 0 To mlPrototypeUB
                    If mlPrototypeIdx(X) = lID Then
                        Return moPrototype(X)
                    End If
                Next X
            Case ObjectType.eRadarTech
                For X As Int32 = 0 To mlRadarUB
                    If mlRadarIdx(X) = lID Then
                        Return moRadar(X)
                    End If
                Next X
            Case ObjectType.eShieldTech
                For X As Int32 = 0 To mlShieldUB
                    If mlShieldIdx(X) = lID Then
                        Return moShield(X)
                    End If
                Next X
            Case ObjectType.eSpecialTech
                For X As Int32 = 0 To oSpecials.mlSpecialTechUB
                    If oSpecials.mlSpecialTechIdx(X) = lID Then
                        Return oSpecials.moSpecialTech(X)
                    End If
                Next X
            Case ObjectType.eWeaponTech
                For X As Int32 = 0 To mlWeaponUB
                    If mlWeaponIdx(X) = lID Then
                        Return moWeapon(X)
                    End If
                Next X
        End Select

        'Ok, if we are here, let's try the initial player object
        If Me.ObjectID <> 0 AndAlso goInitialPlayer Is Nothing = False Then
            Return goInitialPlayer.GetTech(lID, iTypeID)
        End If

        Return Nothing
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

    '    If Me.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return

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
    '    Dim lAddLinkChance As Int32 = 100 \ oSpecials.yMultipleLinkChance
    '    For X As Int32 = 0 To mlSpecialTechUB
    '        If moSpecialTech(X).bLinked = True AndAlso moSpecialTech(X).bInTheTank = False AndAlso moSpecialTech(X).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then lActiveLinks += 1
    '    Next X
    '    If lActiveLinks >= lMaxLinks Then Return

    '    'Now, check for wormhole special tech
    '    If Me.lWormholeUB <> -1 AndAlso (Me.oSpecials.lSuperSpecials And Player_Specials.SuperSpecialID.eWormholesTraversable) = 0 Then
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

    'Private Function HandleLinkTest(ByRef oSpecTech As SpecialTech, ByVal lChance As Int32) As Byte
    '    'First, see if the item is already in the link list
    '    Dim lIdx As Int32 = -1
    '    For X As Int32 = 0 To mlSpecialTechUB
    '        If mlSpecialTechIdx(X) = oSpecTech.ObjectID Then

    '            'Is the tech link failed already?
    '            If moSpecialTech(X).LinkAttempts > oSpecTech.MaxLinkChanceAttempts Then Return 0
    '            If moSpecialTech(X).bLinked = True Then Return 0

    '            lIdx = X
    '            Exit For
    '        End If
    '    Next X

    '    'Ok, is it in the link list?
    '    If lIdx = -1 Then
    '        'no, so add it
    '        lIdx = mlSpecialTechUB + 1
    '        ReDim Preserve moSpecialTech(lIdx)
    '        ReDim Preserve mlSpecialTechIdx(lIdx)
    '        moSpecialTech(lIdx) = New PlayerSpecialTech()

    '        With moSpecialTech(lIdx)
    '            '.bLinked = False
    '            .yFlags = 0
    '            .ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eComponentDesign
    '            .CurrentSuccessChance = oSpecTech.InitialSuccessChance
    '            .ErrorReasonCode = 0
    '            .LinkAttempts = 0
    '            .ObjectID = oSpecTech.ObjectID
    '            .ObjTypeID = ObjectType.eSpecialTech
    '            .Owner = Me
    '            .RandomSeed = 1.0F
    '            .ResearchAttempts = 0
    '            .SuccessChanceIncrement = oSpecTech.IncrementalSuccess
    '        End With
    '        mlSpecialTechIdx(lIdx) = oSpecTech.ObjectID
    '        mlSpecialTechUB += 1
    '    End If

    '    With moSpecialTech(lIdx)
    '        'ok, if we are here, increment our link attempts
    '        '.bNeedsSaved = True
    '        .LinkAttempts += 1
    '        If Rnd() * 100.0F < lChance Then
    '            'ok, it linked!!! so call ComponentDesigned
    '            .ComponentDesigned()

    '            'Ok, send player message, even if offline
    '            Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(lIdx), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
    '            If lConnectedPrimaryID > -1 OrElse HasOnlineAliases(AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns) = True Then
    '                Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(lIdx).GetResearchCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
    '                Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moSpecialTech(lIdx).GetProductionCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
    '            End If
    '            .SaveObject()
    '            Return 1
    '        End If
    '        .SaveObject()
    '    End With

    '    Return 0

    'End Function
    Public Sub CalculateSpecialTechs()
        For X As Int32 = 0 To oSpecials.mlSpecialTechUB
            If oSpecials.mlSpecialTechIdx(X) <> -1 AndAlso oSpecials.moSpecialTech(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then

                Select Case Me.ObjectID
                    Case 367
                        TestAndAddPlayerKnownProperty(eMinPropID.ChemicalReactance, 2, True, Me)
                    Case 368
                        TestAndAddPlayerKnownProperty(eMinPropID.Combustiveness, 2, True, Me)
                    Case 369
                        TestAndAddPlayerKnownProperty(eMinPropID.Compressibility, 2, True, Me)
                        TestAndAddPlayerKnownProperty(eMinPropID.Malleable, 2, True, Me)
                    Case 370
                        TestAndAddPlayerKnownProperty(eMinPropID.MagneticProduction, 2, True, Me)
                        TestAndAddPlayerKnownProperty(eMinPropID.MagneticReaction, 2, True, Me)
                    Case 371
                        TestAndAddPlayerKnownProperty(eMinPropID.ElectricalResist, 2, True, Me)
                    Case 372
                        TestAndAddPlayerKnownProperty(eMinPropID.Psych, 2, True, Me)
                    Case 373
                        TestAndAddPlayerKnownProperty(eMinPropID.Quantum, 2, True, Me)
                    Case 374
                        TestAndAddPlayerKnownProperty(eMinPropID.Reflection, 2, True, Me)
                    Case 375
                        TestAndAddPlayerKnownProperty(eMinPropID.Refraction, 2, True, Me)
                    Case 376
                        TestAndAddPlayerKnownProperty(eMinPropID.SuperconductivePoint, 2, True, Me)
                    Case 377
                        TestAndAddPlayerKnownProperty(eMinPropID.TemperatureSensitivity, 2, True, Me)
                        TestAndAddPlayerKnownProperty(eMinPropID.ThermalConductance, 2, True, Me)
                        TestAndAddPlayerKnownProperty(eMinPropID.ThermalExpansion, 2, True, Me)
                    Case 378
                        TestAndAddPlayerKnownProperty(eMinPropID.Toxicity, 2, True, Me)
                End Select

                oSpecials.ProcessTechResearched(oSpecials.moSpecialTech(X).oTech)
            End If
        Next X
    End Sub
    Private Shared Sub TestAndAddPlayerKnownProperty(ByVal lPropID As Int32, ByVal yDiscovery As Byte, ByVal bAlert As Boolean, ByRef oPlayer As Player)
        For X As Int32 = 0 To oPlayer.mlMinPropertyUB
            If oPlayer.moMinProperties(X).lPropertyID = lPropID Then
                If oPlayer.moMinProperties(X).Discovered <> yDiscovery Then
                    oPlayer.AddKnownProperty(lPropID, yDiscovery, bAlert, True)
                End If
                Return
            End If
        Next X

        'if we are here add it
        oPlayer.AddKnownProperty(lPropID, yDiscovery, bAlert, True)
    End Sub

    'Public Sub LoadTechProdCosts()
    '    For X As Int32 = 0 To mlAlloyUB
    '        If mlAlloyIdx(X) <> -1 Then moAlloy(X).FinalizeTechCostsLoad()
    '    Next X
    '    For X As Int32 = 0 To mlArmorUB
    '        If mlArmorIdx(X) <> -1 Then moArmor(X).FinalizeTechCostsLoad()
    '    Next X
    '    For X As Int32 = 0 To mlEngineUB
    '        If mlEngineIdx(X) <> -1 Then moEngine(X).FinalizeTechCostsLoad()
    '    Next X
    '    For X As Int32 = 0 To mlHullUB
    '        If mlHullIdx(X) <> -1 Then moHull(X).FinalizeTechCostsLoad()
    '    Next X
    '    For X As Int32 = 0 To mlPrototypeUB
    '        If mlPrototypeIdx(X) <> -1 Then moPrototype(X).FinalizeTechCostsLoad()
    '    Next X
    '    For X As Int32 = 0 To mlRadarUB
    '        If mlRadarIdx(X) <> -1 Then moRadar(X).FinalizeTechCostsLoad()
    '    Next X
    '    For X As Int32 = 0 To mlShieldUB
    '        If mlShieldIdx(X) <> -1 Then moShield(X).FinalizeTechCostsLoad()
    '    Next X
    '    For X As Int32 = 0 To mlWeaponUB
    '        If mlWeaponIdx(X) <> -1 Then moWeapon(X).FinalizeTechCostsLoad()
    '    Next X
    'End Sub

    Public Sub GetRandomComponentTech(ByRef lTargetID As Int32, ByRef iTargetTypeID As Int16)
        Dim lType As Int32 = CInt(Math.Floor(Rnd() * 8))
        Select Case lType
            Case 0      'alloy
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlAlloyUB
                    If mlAlloyIdx(X) <> -1 Then lCnt += 1
                Next X

                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlAlloyUB
                    If mlAlloyIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

                If lVal > -1 AndAlso lVal <= mlAlloyUB Then
                    If mlAlloyIdx(lVal) <> -1 Then
                        lTargetID = mlAlloyIdx(lVal)
                        iTargetTypeID = moAlloy(lVal).ObjTypeID
                    End If
                End If
            Case 1      'armor
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlArmorUB
                    If mlArmorIdx(X) <> -1 Then lCnt += 1
                Next X

                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlArmorUB
                    If mlArmorIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

                If lVal > -1 AndAlso lVal <= mlArmorUB Then
                    If mlArmorIdx(lVal) <> -1 Then
                        lTargetID = mlArmorIdx(lVal)
                        iTargetTypeID = moArmor(lVal).ObjTypeID
                    End If
                End If
            Case 2      'engine
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlEngineUB
                    If mlEngineIdx(X) <> -1 Then lCnt += 1
                Next X

                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlEngineUB
                    If mlEngineIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

                If lVal > -1 AndAlso lVal <= mlEngineUB Then
                    If mlEngineIdx(lVal) <> -1 Then
                        lTargetID = mlEngineIdx(lVal)
                        iTargetTypeID = moEngine(lVal).ObjTypeID
                    End If
                End If
            Case 3      'hull
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlHullUB
                    If mlHullIdx(X) <> -1 Then lCnt += 1
                Next X

                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlHullUB
                    If mlHullIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

                If lVal > -1 AndAlso lVal <= mlHullUB Then
                    If mlHullIdx(lVal) <> -1 Then
                        lTargetID = mlHullIdx(lVal)
                        iTargetTypeID = moHull(lVal).ObjTypeID
                    End If
                End If
            Case 4      'prototype
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlPrototypeUB
                    If mlPrototypeIdx(X) <> -1 Then lCnt += 1
                Next X

                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlPrototypeUB
                    If mlPrototypeIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

                If lVal > -1 AndAlso lVal <= mlPrototypeUB Then
                    If mlPrototypeIdx(lVal) <> -1 Then
                        lTargetID = mlPrototypeIdx(lVal)
                        iTargetTypeID = moPrototype(lVal).ObjTypeID
                    End If
                End If
            Case 5      'radar
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlRadarUB
                    If mlRadarIdx(X) <> -1 Then lCnt += 1
                Next X

                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlRadarUB
                    If mlRadarIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

                If lVal > -1 AndAlso lVal <= mlRadarUB Then
                    If mlRadarIdx(lVal) <> -1 Then
                        lTargetID = mlRadarIdx(lVal)
                        iTargetTypeID = moRadar(lVal).ObjTypeID
                    End If
                End If
            Case 6      'shield
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlShieldUB
                    If mlShieldIdx(X) <> -1 Then lCnt += 1
                Next X

                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlShieldUB
                    If mlShieldIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

                If lVal > -1 AndAlso lVal <= mlShieldUB Then
                    If mlShieldIdx(lVal) <> -1 Then
                        lTargetID = mlShieldIdx(lVal)
                        iTargetTypeID = moShield(lVal).ObjTypeID
                    End If
                End If
            Case Else   'weapon
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlWeaponUB
                    If mlWeaponIdx(X) <> -1 Then lCnt += 1
                Next X
                Dim lIdxs(lCnt - 1) As Int32
                Dim lTempIdx As Int32 = -1
                For X As Int32 = 0 To mlWeaponUB
                    If mlWeaponIdx(X) <> -1 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                Next X

                Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
                If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)
                If lVal > -1 AndAlso lVal <= mlWeaponUB Then
                    If mlWeaponIdx(lVal) <> -1 Then
                        lTargetID = mlWeaponIdx(lVal)
                        iTargetTypeID = moWeapon(lVal).ObjTypeID
                    End If
                End If
        End Select
    End Sub
#End Region

#Region "  Player Relationship Management  "
    Private moPlayerRels() As PlayerRel
    Private mlPlayerRelIdx() As Int32
    Public PlayerRelUB As Int32 = -1

    Public Sub SetPlayerRel(ByVal lTargetPlayerID As Int32, ByRef oPlayerRel As PlayerRel, ByVal bSaveRel As Boolean)
        Dim X As Int32
        Dim lIdx As Int32 = -1
        Dim oPI As PlayerIntel = Nothing

        'before proceeding, save our rel
        If bSaveRel = True AndAlso oPlayerRel.SaveObject() = False Then
            LogEvent(LogEventType.CriticalError, "SetPlayerRel: unable to save incoming rel.")
        End If
        If bSaveRel = True Then oPlayerRel.StorePlayerRelHistory()

        For X = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) = lTargetPlayerID Then
                moPlayerRels(X) = oPlayerRel
                'CheckForAndAddPlayerIntel(lTargetPlayerID)
                oPI = GetOrAddPlayerIntel(lTargetPlayerID, False)
                If oPI Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPI, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents Or AliasingRights.eViewDiplomacy)
                Return
            ElseIf lIdx = -1 AndAlso mlPlayerRelIdx(X) = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            PlayerRelUB += 1
            ReDim Preserve moPlayerRels(PlayerRelUB)
            ReDim Preserve mlPlayerRelIdx(PlayerRelUB)
            lIdx = PlayerRelUB
        End If

        moPlayerRels(lIdx) = oPlayerRel
        mlPlayerRelIdx(lIdx) = lTargetPlayerID
        'CheckForAndAddPlayerIntel(lTargetPlayerID)
        oPI = GetOrAddPlayerIntel(lTargetPlayerID, False)
        If oPI Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPI, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents Or AliasingRights.eViewDiplomacy)
    End Sub

    Public Function GetPlayerRelScore(ByVal lPlayerID As Int32) As Byte
        Dim lCurUB As Int32 = -1
        If mlPlayerRelIdx Is Nothing = False Then lCurUB = Math.Min(mlPlayerRelIdx.GetUpperBound(0), PlayerRelUB)
        For X As Int32 = 0 To lCurUB
            If mlPlayerRelIdx(X) = lPlayerID Then
                Return moPlayerRels(X).WithThisScore
            End If
        Next X
        Return elRelTypes.eNeutral
    End Function

    Public Function GetPlayerRelByIndex(ByVal lIndex As Int32) As PlayerRel
        If lIndex < 0 OrElse lIndex > PlayerRelUB Then Return Nothing
        Return moPlayerRels(lIndex)
    End Function

    Public Function GetPlayerRel(ByVal lPlayerID As Int32) As PlayerRel
        Dim lCurUB As Int32 = -1
        If mlPlayerRelIdx Is Nothing = False Then lCurUB = Math.Min(mlPlayerRelIdx.GetUpperBound(0), PlayerRelUB)
        For X As Int32 = 0 To lCurUB
            If mlPlayerRelIdx(X) = lPlayerID Then
                Return moPlayerRels(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Sub AddTradeIncomeItemToAllies(ByVal lAmt As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
        For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) <> -1 AndAlso moPlayerRels(X).WithThisScore > elRelTypes.ePeace AndAlso moPlayerRels(X).oPlayerRegards.ObjectID = Me.ObjectID AndAlso moPlayerRels(X).oThisPlayer.AccountStatus = AccountStatusType.eActiveAccount Then
                'Ok, give this player some money
                'If moPlayerRels(X).oThisPlayer.bInFullLockDown = False Then moPlayerRels(X).oThisPlayer.blCredits += lAmt
                'And indicate as much in the Budget window
                moPlayerRels(X).oThisPlayer.oBudget.AddTraderIncome(lEnvirID, iEnvirTypeID, Me.ObjectID, lAmt)
            End If
        Next X

        ''And to my guild mates if necessary
        'If oGuild Is Nothing = False Then
        '    With oGuild
        '        Dim oMeMember As GuildMember = .GetMember(Me.ObjectID)
        '        If oMeMember Is Nothing = False AndAlso (oMeMember.yMemberState And GuildMemberState.Approved) <> 0 Then
        '            If .oMembers Is Nothing = False Then
        '                Dim lCurUB As Int32 = Math.Min(.lMemberUB, .oMembers.GetUpperBound(0))
        '                For X As Int32 = 0 To lCurUB
        '                    Dim oMember As GuildMember = .oMembers(X)
        '                    If oMember Is Nothing = False AndAlso (oMember.yMemberState And GuildMemberState.Approved) <> 0 AndAlso oMember.oMember Is Nothing = False AndAlso oMember.oMember.ObjectID <> Me.ObjectID Then

        '                        If GetPlayerRelScore(oMember.lMemberID) < elRelTypes.eAlly AndAlso oMember.oMember.yPlayerPhase = Me.yPlayerPhase Then
        '                            'Ok, give this player some money
        '                            'If oMember.oMember.bInFullLockDown = False Then oMember.oMember.blCredits += lAmt
        '                            'and indicate as much in the budget window
        '                            oMember.oMember.oBudget.AddTraderIncome(lEnvirID, iEnvirTypeID, Me.ObjectID, lAmt)
        '                        End If

        '                    End If
        '                Next X
        '            End If
        '        End If
        '    End With
        'End If
    End Sub

    Public Function PlayerIsAtWar() As Boolean
        For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) <> -1 Then
                If moPlayerRels(X).WithThisScore <= elRelTypes.eWar Then
                    If moPlayerRels(X).oThisPlayer.PlayerIsDead = False Then Return True
                ElseIf moPlayerRels(X).oThisPlayer.GetPlayerRelScore(Me.ObjectID) <= elRelTypes.eWar Then
                    If moPlayerRels(X).oThisPlayer.PlayerIsDead = False Then Return True
                End If
            End If
        Next X
        Return False
    End Function

    Public Sub RemovePlayerRel(ByVal lWithPlayerID As Int32, ByVal bNoRecursion As Boolean)

        Dim yMsg(9) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRemovePlayerRel).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 6)

        Dim lShiftIdx As Int32 = -1
        'Remove all relationships
        For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) <> -1 AndAlso (lWithPlayerID = -1 OrElse lWithPlayerID = mlPlayerRelIdx(X)) Then
                If moPlayerRels(X).oThisPlayer Is Nothing = False AndAlso moPlayerRels(X).oPlayerRegards Is Nothing = False Then
                    If moPlayerRels(X).oThisPlayer.ObjectID = Me.ObjectID Then
                        'ok, wrong direction
                        If bNoRecursion = False Then moPlayerRels(X).oPlayerRegards.RemovePlayerRel(Me.ObjectID, True)
                        System.BitConverter.GetBytes(moPlayerRels(X).oPlayerRegards.ObjectID).CopyTo(yMsg, 2)
                        moPlayerRels(X).oPlayerRegards.SendPlayerMessage(yMsg, False, AliasingRights.eViewDiplomacy)
                    Else
                        If bNoRecursion = False Then moPlayerRels(X).oThisPlayer.RemovePlayerRel(Me.ObjectID, True)
                        System.BitConverter.GetBytes(moPlayerRels(X).oThisPlayer.ObjectID).CopyTo(yMsg, 2)
                        moPlayerRels(X).oThisPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewDiplomacy)
                    End If
                End If

                If lWithPlayerID <> -1 Then
                    lShiftIdx = X
                    Exit For
                End If
            End If
        Next X
        If lWithPlayerID = -1 Then
            PlayerRelUB = -1

            If bNoRecursion = False Then
                'Now, delete the records 
                Dim oComm As OleDb.OleDbCommand = Nothing
                Try
                    oComm = New OleDb.OleDbCommand("DELETE FROM tblPlayerToPlayerRel WHERE Player1ID = " & Me.ObjectID & " OR Player2ID = " & Me.ObjectID, goCN)
                    oComm.ExecuteNonQuery()
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "Unable to Delete P2PRels for lWithPlayerId = -1. Reason: " & ex.Message)
                Finally
                    If oComm Is Nothing = False Then oComm.Dispose()
                    oComm = Nothing
                End Try
            End If

        Else
            If bNoRecursion = False Then
                'Now, delete the records 
                Dim oComm As OleDb.OleDbCommand = Nothing
                Try
                    oComm = New OleDb.OleDbCommand("DELETE FROM tblPlayerToPlayerRel WHERE (Player1ID = " & Me.ObjectID & " AND Player2ID = " & lWithPlayerID & ") OR (Player1ID = " & lWithPlayerID & " AND Player2ID = " & Me.ObjectID & ")", goCN)
                    oComm.ExecuteNonQuery()
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "Unable to Delete P2PRels for lWithPlayerId = " & lWithPlayerID & ". Reason: " & ex.Message)
                Finally
                    If oComm Is Nothing = False Then oComm.Dispose()
                    oComm = Nothing
                End Try
            End If

            'Clear the item
            For X As Int32 = lShiftIdx To PlayerRelUB - 1
                moPlayerRels(X) = moPlayerRels(X + 1)
                mlPlayerRelIdx(X) = mlPlayerRelIdx(X + 1)
            Next X
            PlayerRelUB -= 1
        End If


    End Sub

    Public Function GetPlayerWarCount() As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) <> -1 Then
                If moPlayerRels(X).WithThisScore <= elRelTypes.eWar Then
                    If moPlayerRels(X).oThisPlayer.PlayerIsDead = False Then lCnt += 1
                ElseIf moPlayerRels(X).oThisPlayer.GetPlayerRelScore(Me.ObjectID) <= elRelTypes.eWar Then
                    If moPlayerRels(X).oThisPlayer.PlayerIsDead = False Then lCnt += 1
                End If
            End If
        Next X
        Return lCnt
    End Function
#End Region

#Region "  Mineral, Property, Knowledge Management  "

    Public Function CheckFirstContactWithMineral(ByVal lMineralID As Int32) As Boolean
        Dim bResult As Boolean = True

        For X As Int32 = 0 To lPlayerMineralUB
            If oPlayerMinerals(X).lMineralID = lMineralID Then
                bResult = False
                Exit For
            End If
        Next X

        If bResult = True Then
            'Ok, first contact... add the PlayerMineral object
            Dim lIdx As Int32 = Me.AddPlayerMineral(lMineralID, False, -1, False)
            Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me.oPlayerMinerals(lIdx), GlobalMessageCode.eAddObjectCommand), True, AliasingRights.eViewMining Or AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)

            'CheckForNewMineralPropertyTechs(lMineralID, 4)
        End If

        Return bResult
    End Function

    Public Function GetMineralPropertyKnowledge(ByVal lMineralID As Int32, ByVal lPropertyID As Int32) As Int32
        For X As Int32 = 0 To lPlayerMineralUB
            If oPlayerMinerals(X).lMineralID = lMineralID Then
                'Ok, found the mineral...
                Return oPlayerMinerals(X).GetPropertyValue(lPropertyID)
            End If
        Next X
        Return 0
    End Function

    'RETURNS THE INDEX OF THE PLAYERMINERAL ARRAY
    Public Function AddPlayerMineral(ByVal lMineralID As Int32, ByVal bDiscovered As Boolean, ByVal lPlayerMineralID As Int32, ByVal bNoSave As Boolean) As Int32
        If lPlayerMineralID = -1 Then
            For X As Int32 = 0 To lPlayerMineralUB
                If oPlayerMinerals(X).lMineralID = lMineralID Then
                    oPlayerMinerals(X).bDiscovered = bDiscovered
                    oPlayerMinerals(X).oPlayer = Me

                    If bNoSave = False Then oPlayerMinerals(X).SaveObject(Me.ObjectID)
                    Return X
                End If
            Next X
        End If

        lPlayerMineralUB += 1
        ReDim Preserve oPlayerMinerals(lPlayerMineralUB)
        oPlayerMinerals(lPlayerMineralUB) = New PlayerMineral
        With oPlayerMinerals(lPlayerMineralUB)
            .PlayerMineralID = lPlayerMineralID
            .lMineralID = lMineralID
            .bDiscovered = bDiscovered
            .DiscoveryNumber = oPlayerMinerals.Length       'TODO: a bit of a hack...
            .oPlayer = Me

            If bNoSave = False Then .SaveObject(Me.ObjectID)
        End With
        Return lPlayerMineralUB
    End Function

    Public Shared Function GetPlayerMineralIDWhereClause(ByRef oPlayer As Player) As String
        Dim sResult As String = ""
        For X As Int32 = 0 To oPlayer.lPlayerMineralUB
            If sResult = "" Then
                sResult = oPlayer.oPlayerMinerals(X).PlayerMineralID.ToString
            Else : sResult &= ", " & oPlayer.oPlayerMinerals(X).PlayerMineralID.ToString
            End If
        Next X
        Return sResult
    End Function

    Public Sub SetPlayerMineralProperty(ByVal lMineralPropertyValueID As Int32, ByVal lPlayerMineralID As Int32, ByVal lPropertyID As Int32, ByVal lValue As Int32)
        For X As Int32 = 0 To lPlayerMineralUB
            If oPlayerMinerals(X).PlayerMineralID = lPlayerMineralID Then
                oPlayerMinerals(X).SetPropertyValue(lMineralPropertyValueID, lPropertyID, lValue)
                Return
            End If
        Next X
    End Sub

    ''' <summary>
    ''' Adds a property to the player and sets its discover level. If this property is new and bAlertPlayer is set,
    '''   a notification is sent to the player object. 
    ''' </summary>
    ''' <param name="lPropertyID"></param>
    ''' <param name="yDiscovered"> The discovery level of the property </param>
    ''' <param name="bAlertPlayer"> Sends the resulting mineral property to the player (Recommend True) </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddKnownProperty(ByVal lPropertyID As Int32, ByVal yDiscovered As Byte, ByVal bAlertPlayer As Boolean, ByVal bSave As Boolean) As Int32
        Dim lIdx As Int32 = -1
        'Check if known property exists
        For X As Int32 = 0 To mlMinPropertyUB
            If moMinProperties(X).lPropertyID = lPropertyID Then
                moMinProperties(X).Discovered = yDiscovered
                lIdx = X
                If bSave = True Then moMinProperties(X).SaveObject()
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlMinPropertyUB += 1
            ReDim Preserve moMinProperties(mlMinPropertyUB)
            moMinProperties(mlMinPropertyUB) = New PlayerMineralProperty
            With moMinProperties(mlMinPropertyUB)
                .lPropertyID = lPropertyID
                .lPlayerID = Me.ObjectID
                .Discovered = yDiscovered
                If bSave = True Then .SaveObject()
            End With
            lIdx = mlMinPropertyUB
        End If

        If bAlertPlayer = True AndAlso (Me.lConnectedPrimaryID > -1 OrElse Me.HasOnlineAliases(AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns) = True) Then
            Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moMinProperties(lIdx), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
            If yDiscovered = 0 Then
                For X As Int32 = 0 To Me.lPlayerMineralUB
                    Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPlayerMinerals(X), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                Next
            End If
        End If

        Return lIdx
    End Function

    Public Function IsMineralDiscovered(ByVal lMineralID As Int32) As Boolean
        For X As Int32 = 0 To lPlayerMineralUB
            If oPlayerMinerals(X).lMineralID = lMineralID Then
                Return oPlayerMinerals(X).bDiscovered
            End If
        Next X
        Return False
    End Function

    Public Function IsMineralPropertyKnown(ByVal lPropID As Int32) As Boolean
        For X As Int32 = 0 To mlMinPropertyUB
            If moMinProperties(X).lPropertyID = lPropID Then
                Return moMinProperties(X).Discovered > 0
            End If
        Next X
        Return False
    End Function

    Public Sub HandleAutomaticMineralMovement()
        Dim lMinToMove As Int32 = Me.oSpecials.yAutomaticMinMove

        'Now, go through our colonies...
        Dim lTmpColUB As Int32 = mlColonyUB
        If mlColonyIdx Is Nothing = False Then lTmpColUB = Math.Min(lTmpColUB, mlColonyIdx.GetUpperBound(0))
        For X As Int32 = 0 To lTmpColUB
            If mlColonyIdx(X) <> -1 Then
                If mlColonyID(X) = glColonyIdx(mlColonyIdx(X)) Then
                    Dim oColony As Colony = goColony(mlColonyIdx(X))
                    If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = Me.ObjectID Then
                        'ok, we're good... now, loop through facilities
                        Dim lTmpChildUB As Int32 = oColony.ChildrenUB
                        If oColony.lChildrenIdx Is Nothing = False Then lTmpChildUB = Math.Min(lTmpChildUB, oColony.lChildrenIdx.GetUpperBound(0))
                        If oColony.oChildren Is Nothing = False Then lTmpChildUB = Math.Min(lTmpChildUB, oColony.oChildren.GetUpperBound(0))
                        With oColony
                            Dim bColonyFull As Boolean = False
                            For Y As Int32 = 0 To lTmpChildUB
                                If .lChildrenIdx(Y) <> -1 Then
                                    Dim oFac As Facility = .oChildren(Y)
                                    If oFac Is Nothing = False AndAlso oFac.yProductionType = ProductionType.eMining AndAlso oFac.Active = True Then
                                        'ok, this mine transports any materials from its cargo hold to the colony (if possible)
                                        Dim lTmpCargoUB As Int32 = oFac.lCargoUB
                                        If oFac.lCargoIdx Is Nothing = False Then lTmpCargoUB = Math.Min(lTmpCargoUB, oFac.lCargoIdx.GetUpperBound(0))
                                        If oFac.oCargoContents Is Nothing = False Then lTmpCargoUB = Math.Min(lTmpCargoUB, oFac.oCargoContents.GetUpperBound(0))
                                        For Z As Int32 = 0 To lTmpCargoUB
                                            If oFac.lCargoIdx(Z) <> -1 Then
                                                If oFac.oCargoContents(Z).ObjTypeID = ObjectType.eMineralCache Then
                                                    Dim oCache As MineralCache = CType(oFac.oCargoContents(Z), MineralCache)
                                                    If oCache Is Nothing = False AndAlso oCache.oMineral Is Nothing = False Then
                                                        Dim lTmpQty As Int32 = Math.Min(lMinToMove, oCache.Quantity)
                                                        Dim lResult As Int32 = .AddObjectCaches(oCache.oMineral.ObjectID, ObjectType.eMineral, lTmpQty, False)
                                                        If lResult <> 0 Then
                                                            lResult = lTmpQty - lResult
                                                            oCache.Quantity -= lResult
                                                            bColonyFull = True
                                                        End If
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        Next Z

                                        'ok, is the colony full?
                                        If bColonyFull = True Then Exit For
                                    End If
                                End If
                            Next Y
                        End With
                    End If
                End If
            End If
        Next X
    End Sub

    Public Function GetRandomKnownMineral() As Int32
        Dim lIdx As Int32 = CInt(Rnd() * lPlayerMineralUB)
        If lIdx > -1 AndAlso lIdx < lPlayerMineralUB Then
            If oPlayerMinerals(lIdx) Is Nothing = False Then Return oPlayerMinerals(lIdx).lMineralID
        End If
        Return -1
    End Function

    Public Function ResearchedMineralCount() As Int32
        Dim lCnt As Int32 = 0

        For X As Int32 = 0 To lPlayerMineralUB
            If oPlayerMinerals(X) Is Nothing = False Then
                If oPlayerMinerals(X).bDiscovered = True Then
                    If oPlayerMinerals(X).Mineral.lAlloyTechID < 1 Then lCnt += 1
                End If
            End If
        Next X

        Return lCnt
    End Function

#End Region

#Region "  Player Communication (Email and Socket)  "
    Public EmailFolders() As PlayerCommFolder
    Public EmailFolderIdx() As Int32
    Public EmailFolderUB As Int32 = -1

    Private muEnvirAlerts() As EnvirAlert
    Private mlEnvirAlertUB As Int32 = -1

    Public Function GetOrAddEmailFolder(ByVal lPCF_ID As Int32) As Int32
        Dim lIdx As Int32 = -1

        'Ok, find our folder first...
        For X As Int32 = 0 To EmailFolderUB
            If EmailFolderIdx(X) = lPCF_ID Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            Dim sName As String = "Unnamed"
            Select Case lPCF_ID
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF
                    sName = "Deleted"
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eDrafts_PCF
                    sName = "Drafts"
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF
                    sName = "Inbox"
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eOutbox_PCF
                    sName = "Outbox"
            End Select
            'Now, add the folder
            lIdx = AddFolder(lPCF_ID, sName)
            If lIdx <> -1 Then
                EmailFolderIdx(lIdx) = lIdx
                EmailFolders(lIdx).PlayerID = Me.ObjectID
            End If
        End If
        EmailFolders(lIdx).PlayerID = Me.ObjectID
        Return lIdx
    End Function

    Public Function AddEmailMsg(ByVal lPC_ID As Int32, ByVal lPCF_ID As Int32, ByVal iEncryption As Short, ByVal sBody As String, ByVal sTitle As String, ByVal lSentByID As Int32, ByVal lSentOn As Int32, ByVal bRead As Boolean, ByVal sRecipients As String, ByVal uAttachments() As PlayerComm.WPAttachment) As PlayerComm
        If Me.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return Nothing

        If Me.InMyDomain = False AndAlso Me.ObjectID <> 19326 Then
            'ok, need to do a passthru...
            Dim lBodyLen As Int32 = 0
            Dim lTitleLen As Int32 = 0
            Dim lRecipientLen As Int32 = 0
            Dim lAttachCnt As Int32 = 0
            If sBody Is Nothing = False Then lBodyLen = sBody.Length
            If sTitle Is Nothing = False Then lTitleLen = sTitle.Length
            If sRecipients Is Nothing = False Then lRecipientLen = sRecipients.Length
            If uAttachments Is Nothing = False Then lAttachCnt = uAttachments.Length

            Dim yMsg(38 + lBodyLen + lTitleLen + lRecipientLen + (lAttachCnt * 35)) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPC_ID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(ObjectType.ePlayerComm).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPCF_ID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iEncryption).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lSentByID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSentOn).CopyTo(yMsg, lPos) : lPos += 4
            If bRead = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
            lPos += 1
            System.BitConverter.GetBytes(lBodyLen).CopyTo(yMsg, lPos) : lPos += 4
            If sBody Is Nothing = False Then
                StringToBytes(sBody).CopyTo(yMsg, lPos) : lPos += lBodyLen
            End If
            System.BitConverter.GetBytes(lTitleLen).CopyTo(yMsg, lPos) : lPos += 4
            If sTitle Is Nothing = False Then
                StringToBytes(sTitle).CopyTo(yMsg, lPos) : lPos += lTitleLen
            End If
            System.BitConverter.GetBytes(lRecipientLen).CopyTo(yMsg, lPos) : lPos += 4
            If sRecipients Is Nothing = False Then
                StringToBytes(sRecipients).CopyTo(yMsg, lPos) : lPos += lRecipientLen
            End If
            System.BitConverter.GetBytes(lAttachCnt).CopyTo(yMsg, lPos) : lPos += 4
            If uAttachments Is Nothing = False Then
                For X As Int32 = 0 To uAttachments.GetUpperBound(0)
                    With uAttachments(X)
                        yMsg(lPos) = .AttachNumber : lPos += 1
                        System.BitConverter.GetBytes(.EnvirID).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.EnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
                        System.BitConverter.GetBytes(.LocX).CopyTo(yMsg, lPos) : lPos += 4
                        System.BitConverter.GetBytes(.LocZ).CopyTo(yMsg, lPos) : lPos += 4

                        If .sWPName Is Nothing OrElse .sWPName = "" Then .sWPName = "WP " & .AttachNumber
                        If .sWPName.Length > 20 Then .sWPName = Mid(.sWPName, 1, 20)
                        If .yWPNameBytes Is Nothing OrElse .yWPNameBytes.GetUpperBound(0) <> 19 Then
                            ReDim .yWPNameBytes(19)
                            StringToBytes(.sWPName).CopyTo(.yWPNameBytes, 0)
                        End If
                        .yWPNameBytes.CopyTo(yMsg, lPos) : lPos += 20
                    End With
                Next X
            End If

            goMsgSys.SendPassThruMsg(yMsg, Me.ObjectID, Me.ObjectID, ObjectType.ePlayer)
            Return Nothing
        End If

        Dim lIdx As Int32 = GetOrAddEmailFolder(lPCF_ID)
        If lIdx <> -1 Then
            Dim oPC As PlayerComm = EmailFolders(lIdx).AddPlayerComm(lPC_ID, iEncryption, sBody, sTitle, lSentByID, lSentOn, bRead, sRecipients)
            If oPC Is Nothing = False Then
                If uAttachments Is Nothing = False Then
                    For X As Int32 = 0 To uAttachments.GetUpperBound(0)
                        With uAttachments(X)
                            oPC.AddEmailAttachment(.AttachNumber, .EnvirID, .EnvirTypeID, .LocX, .LocZ, .sWPName)
                        End With
                    Next X
                    oPC.SaveObject(Me.ObjectID)
                End If
            End If

            Return oPC ' EmailFolders(lIdx).AddPlayerComm(lPC_ID, iEncryption, sBody, sTitle, lSentByID, lSentOn, bRead, sRecipients)
        End If
        Return Nothing
    End Function

    Public Function AddFolder(ByVal lBaseID As Int32, ByVal sName As String) As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1

        'Ok, first, see if the base ID is 0 or -1
        If lBaseID = 0 OrElse lBaseID = -1 Then
            'Ok, no base ID... so, we'll go through and look for an empty slot
            For X As Int32 = 0 To EmailFolderUB
                If EmailFolderIdx(X) = -1 Then
                    lFirstIdx = X
                    Exit For
                End If
            Next X
        Else
            'check if we already have it...
            For X As Int32 = 0 To EmailFolderUB
                If EmailFolderIdx(X) = lBaseID Then
                    lIdx = X
                    Exit For
                ElseIf lFirstIdx = -1 AndAlso EmailFolderIdx(X) = -1 Then
                    lFirstIdx = X
                End If
            Next X
        End If

        'Now, check our result
        If lIdx = -1 Then
            'Did not find it... did we find an empty slot?
            If lFirstIdx = -1 Then
                'No, so expand our array
                EmailFolderUB += 1
                ReDim Preserve EmailFolders(EmailFolderUB)
                ReDim Preserve EmailFolderIdx(EmailFolderUB)
                lIdx = EmailFolderUB
            Else : lIdx = lFirstIdx
            End If
            EmailFolders(lIdx) = New PlayerCommFolder()
        End If

        'Now, make sure the folder has our stuff
        If sName.Length > 20 Then sName = sName.Substring(0, 20)

        If EmailFolders(lIdx) Is Nothing Then EmailFolders(lIdx) = New PlayerCommFolder
        With EmailFolders(lIdx)
            StringToBytes(sName).CopyTo(.FolderName, 0)
            If lBaseID <> 0 AndAlso lBaseID <> -1 Then .PCF_ID = lBaseID Else .PCF_ID = -1
            .PlayerID = Me.ObjectID
            If .PCF_ID = -1 Then
                .SaveFolder()
            End If
            EmailFolderIdx(lIdx) = .PCF_ID
        End With

        Return lIdx
    End Function

    Public Function DeleteFolder(ByVal lPCF_ID As Int32) As Boolean
        If lPCF_ID > 0 Then
            For X As Int32 = 0 To EmailFolderUB
                If EmailFolderIdx(X) = lPCF_ID Then
                    Dim oFolder As PlayerCommFolder = EmailFolders(X)
                    If oFolder Is Nothing = False Then oFolder.DeleteFolder()
                    EmailFolderIdx(X) = -1
                    EmailFolders(X) = Nothing
                    Return True
                End If
            Next X
        End If
        Return False
    End Function

    Public Sub RemoveEmailMsg(ByVal lMsgID As Int32, ByVal lPCF_ID As Int32)
        Dim lFolderIdx As Int32 = -1
        Dim lTrashFolderIdx As Int32 = GetOrAddEmailFolder(PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF)

        For X As Int32 = 0 To EmailFolderUB
            If EmailFolderIdx(X) = lPCF_ID Then
                lFolderIdx = X
                Exit For
            End If
        Next X

        If lFolderIdx <> -1 Then
            With EmailFolders(lFolderIdx)

                'Ok, are we removing this message from the trash can?
                If lPCF_ID = PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF Then
                    'yes, so, delete it
                    For X As Int32 = 0 To .PlayerMsgsUB
                        If .PlayerMsgsIdx(X) <> -1 AndAlso (lMsgID = -1 OrElse .PlayerMsgsIdx(X) = lMsgID) Then
                            'Ok, found the message...
                            .PlayerMsgs(X).DeleteComm()
                            .PlayerMsgsIdx(X) = -1
                            .PlayerMsgs(X) = Nothing
                            If lMsgID <> -1 Then Exit For
                        End If
                    Next X
                Else
                    'No, so put the message in the trash can
                    If lMsgID = -1 Then
                        For X As Int32 = 0 To .PlayerMsgsUB
                            If .PlayerMsgsIdx(X) <> -1 Then
                                .MoveMessageToFolder(.PlayerMsgs(X).ObjectID, EmailFolders(lTrashFolderIdx))
                            End If
                        Next X
                    Else
                        .MoveMessageToFolder(lMsgID, EmailFolders(lTrashFolderIdx))
                    End If
                End If
            End With
        End If
    End Sub

    Public Sub SendPlayerMessage(ByRef yMsg() As Byte, ByVal bAddEmail As Boolean, ByVal lRights As AliasingRights)
        If Me.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return

        For X As Int32 = 0 To lAliasUB
            Try
                If lAliasIdx(X) <> -1 AndAlso (lRights = 0 OrElse (uAliasLogin(X).lRights And lRights) <> 0) AndAlso oAliases(X) Is Nothing = False AndAlso oAliases(X).lAliasingPlayerID = Me.ObjectID AndAlso oAliases(X).lConnectedPrimaryID > -1 Then
                    'If oAliases(X).bInPlayerRequestDetails = False Then oAliases(X).oSocket.SendData(yMsg)
                    If oAliases(X).bInPlayerRequestDetails = False Then oAliases(X).CrossPrimarySafeSendMsg(yMsg)
                End If
            Catch
            End Try
        Next X

        'Changed this block, to the following two lines:  When a player was partially disconnected it would not send the Email msg
        'If lConnectedPrimaryID = glServerID Then
        'If oSocket Is Nothing = False Then
        'If oSocket.IsConnected = True Then
        'If bInPlayerRequestDetails = False Then oSocket.SendData(yMsg)
        'Else : oSocket = Nothing
        'End If
        'End If
        If oSocket Is Nothing = False AndAlso oSocket.IsConnected = False Then oSocket = Nothing
        If lConnectedPrimaryID = glServerID AndAlso oSocket Is Nothing = False AndAlso oSocket.IsConnected = True Then
            If bInPlayerRequestDetails = False Then oSocket.SendData(yMsg)
        ElseIf ((lConnectedPrimaryID <> glServerID AndAlso lConnectedPrimaryID > -1) OrElse Me.InMyDomain = False) AndAlso Me.ObjectID <> 19326 Then
            CrossPrimarySafeSendMsg(yMsg)
        ElseIf bAddEmail = True Then
            Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yMsg, 0)

            Dim sBody As String
            Dim sTitle As String
            Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim lPos As Int32 = 2

            Select Case iMsgCode
                Case GlobalMessageCode.eOutBidAlert
                    'ok, an outbid alert.... facid, factypeid, envirid, envirtypeid, newmaxbid
                    If (Me.iInternalEmailSettings And eEmailSettings.eMineralOutbid) <> 0 Then
                        'send an email alert
                        Dim lFacID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                        Dim iFacTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
                        Dim oFac As Facility = GetEpicaFacility(lFacID)
                        If oFac Is Nothing = False AndAlso oFac.oMiningBid Is Nothing = False AndAlso oFac.oMiningBid.oMineralCache Is Nothing = False AndAlso oFac.oMiningBid.oMineralCache.oMineral Is Nothing = False Then
                            Dim lBidState As Int32 = oFac.oMiningBid.GetPlayerBidSlot(Me.ObjectID)
                            Dim oGUID As Epica_GUID = oFac.GetRootParentEnvir()
                            Dim sSysName As String = ""
                            If oGUID Is Nothing = False Then
                                If oGUID.ObjTypeID = ObjectType.ePlanet Then
                                    sSysName = BytesToString(CType(oGUID, Planet).PlanetName)
                                ElseIf oGUID.ObjTypeID = ObjectType.eSolarSystem Then
                                    sSysName = BytesToString(CType(oGUID, SolarSystem).SystemName)
                                End If
                            End If
                            If sSysName <> "" Then sSysName = " in " & sSysName & "." Else sSysName = "."
                            Select Case lBidState
                                Case 0
                                    oSB.AppendLine("You are no longer bidding at a mining facility" & sSysName)
                                Case 1
                                    Return
                                Case 2
                                    oSB.AppendLine("You are now in second place for bidding at a mining facility" & sSysName)
                                Case 3
                                    oSB.AppendLine("You are now in third place for bidding at a mining facility" & sSysName)
                                Case 4
                                    oSB.AppendLine("You are now in fourth place for bidding at a mining facility" & sSysName)
                                Case -1
                                    oSB.AppendLine("You no longer meet the minimum bid at a mining facility" & sSysName)
                                Case 10
                                    oSB.AppendLine("You are no longer in the top for for bidding at a mining facility" & sSysName)
                            End Select
                            oSB.AppendLine("The mining facility is mining " & BytesToString(oFac.oMiningBid.oMineralCache.oMineral.MineralName) & " on " & BytesToString(oFac.EntityName) & " at a rate of " & oFac.oMiningBid.lPreviousProductionRate & " minerals.")
                            oSB.AppendLine("The minimum bid for the mine is " & oFac.oMiningBid.lMinBidAmt & ".")

                            Dim lFirstPlaceBidderID As Int32 = oFac.oMiningBid.GetFirstPlaceBidderID
                            Dim lFirstPlaceBid As Int32 = oFac.oMiningBid.GetFirstPlaceBid
                            If lFirstPlaceBidderID > 0 Then
                                Dim oFirst As Player = GetEpicaPlayer(lFirstPlaceBidderID)
                                If oFirst Is Nothing = False Then
                                    oSB.AppendLine("The winning bid for the mine is " & oFirst.sPlayerNameProper & " at " & lFirstPlaceBid & ".")
                                End If
                            End If

                            Dim uWP(0) As PlayerComm.WPAttachment
                            With uWP(0)
                                .AttachNumber = 1
                                .EnvirID = CType(oFac.ParentObject, Epica_GUID).ObjectID
                                .EnvirTypeID = CType(oFac.ParentObject, Epica_GUID).ObjTypeID
                                .LocX = oFac.LocX
                                .LocZ = oFac.LocZ
                                .yWPNameBytes = oFac.EntityName
                                .sWPName = BytesToString(oFac.EntityName)
                            End With

                            Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Mining Bid Alert", Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName), uWP)
                            If oPC Is Nothing = False AndAlso (iEmailSettings And eEmailSettings.eMineralOutbid) <> 0 Then goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, lFacID, -1, -1, -1, -1, -1, "Respond with 'SET BID' and an amount.")

                        End If
                    End If
                Case GlobalMessageCode.eAddObjectCommand
                    Dim lObjID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
                    If (iInternalEmailSettings And eEmailSettings.eResearchComplete) <> 0 Then
                        Select Case iObjTypeID
                            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
                                Dim oTech As Epica_Tech = Me.GetTech(lObjID, iObjTypeID)
                                If oTech Is Nothing = False Then
                                    oSB.Length = 0
                                    oSB.AppendLine("Your scientists want to inform you")
                                    oSB.AppendLine("that research has been completed on")

                                    Dim sTypeStr As String = ""
                                    Select Case iObjTypeID
                                        Case ObjectType.eAlloyTech : sTypeStr = " (Alloy)"
                                        Case ObjectType.eArmorTech : sTypeStr = " (Armor)"
                                        Case ObjectType.eEngineTech : sTypeStr = " (Engine)"
                                        Case ObjectType.eHullTech : sTypeStr = " (Hull)"
                                        Case ObjectType.ePrototype : sTypeStr = " (Prototype)"
                                        Case ObjectType.eRadarTech : sTypeStr = " (Radar)"
                                        Case ObjectType.eShieldTech : sTypeStr = " (Shield)"
                                        Case ObjectType.eSpecialTech : sTypeStr = " (Special Project)"
                                        Case ObjectType.eWeaponTech : sTypeStr = " (Weapon)"
                                    End Select
                                    oSB.AppendLine(BytesToString(oTech.GetTechName) & sTypeStr)
                                    sBody = oSB.ToString
                                    sTitle = "Research Complete"

                                    Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName), Nothing)
                                    If oPC Is Nothing = False AndAlso (iEmailSettings And eEmailSettings.eResearchComplete) <> 0 Then goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, 0, 0, 0, 0, 0, 0, "")
                                End If
                            Case Else
                                'TODO: what else? for now return
                                Return
                        End Select
                    End If
                Case GlobalMessageCode.eSetPlayerRel
                    If (iInternalEmailSettings And eEmailSettings.ePlayerRels) <> 0 Then
                        oSB.Length = 0
                        oSB.AppendLine("Player Relationship Altered" & vbCrLf)
                        lPos = 2
                        Dim lPlayer1ID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                        Dim lPlayer2ID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                        Dim yScore As Byte = yMsg(lPos) : lPos += 1
                        Dim lExt1 As Int32
                        Dim bSendExt As Boolean = (iEmailSettings And eEmailSettings.ePlayerRels) <> 0

                        If Me.ObjectID = lPlayer1ID Then
                            Dim oPlayer As Player = GetEpicaPlayer(lPlayer2ID)
                            lExt1 = lPlayer2ID
                            If oPlayer Is Nothing = False Then
                                oSB.AppendLine("New Relationship Value with " & oPlayer.sPlayerName & ": " & yScore)
                            End If
                        Else
                            Dim oPlayer As Player = GetEpicaPlayer(lPlayer1ID)
                            lExt1 = lPlayer1ID
                            If oPlayer Is Nothing = False Then
                                oSB.AppendLine("New Relationship Value with " & oPlayer.sPlayerName & ": " & yScore)
                            End If
                        End If

                        Dim oSBExtra As New System.Text.StringBuilder()
                        oSBExtra.Length = 0
                        If bSendExt = True Then
                            oSBExtra.AppendLine()
                            oSBExtra.AppendLine("How would you like to respond?" & vbCrLf)
                            oSBExtra.AppendLine("Set Relationship to <enter value>")
                            oSBExtra.AppendLine("Raise Full Invulnerability")
                        End If

                        sBody = oSB.ToString
                        sTitle = "Foreign Policy Change"
                        Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False AndAlso bSendExt = True Then goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, lExt1, yScore, 0, 0, 0, 0, oSBExtra.ToString)
                    End If
                Case GlobalMessageCode.eRequestDockFail
                    oSB.Length = 0
                    oSB.AppendLine("Docking Request Rejected")
                    oSB.AppendLine()


                    Dim lObjID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
                    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
                    lPos += 8       'loc
                    Dim yReason As Byte = yMsg(lPos) : lPos += 1

                    'Now, construct the rest
                    If iEnvirTypeID = ObjectType.ePlanet Then
                        Dim oP As Planet = GetEpicaPlanet(lEnvirID)
                        If oP Is Nothing = False Then
                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
                        End If
                    ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
                        Dim oS As SolarSystem = GetEpicaSystem(lEnvirID)
                        If oS Is Nothing = False Then
                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
                        End If
                    End If
                    Select Case yReason
                        Case DockRejectType.eDockeeMoving
                            oSB.AppendLine("Reason: Dock Target was Moving")
                        Case DockRejectType.eDockerCannotEnterEnvir
                            oSB.AppendLine("Reason: Undock Environment was inhospitable")
                        Case DockRejectType.eDoorSizeExceeded
                            oSB.AppendLine("Reason: Target's Hangar Bay Doors were not big enough")
                        Case DockRejectType.eHangarInoperable
                            oSB.AppendLine("Reason: Target's Hangar was inoperable")
                        Case DockRejectType.eInsufficientHangarCap
                            oSB.AppendLine("Reason: Target's Hangar Capacity was Insufficient")
                    End Select

                    sBody = oSB.ToString
                    sTitle = "Docking Request Rejection"

                    Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName), Nothing)
                Case GlobalMessageCode.eColonyLowResources
                    If (iInternalEmailSettings And eEmailSettings.eLowResources) = 0 Then Return
                    Dim yTemp(19) As Byte
                    Array.Copy(yMsg, lPos, yTemp, 0, 20)
                    lPos += 20
                    Dim sColonyName As String = BytesToString(yTemp)
                    Dim lColonyID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    lPos += 2       'colony guid
                    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
                    Dim yReason As Byte = yMsg(lPos) : lPos += 1

                    Dim lItemID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2

                    Dim lProdID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2

                    Dim bSendExt As Boolean = (iEmailSettings And eEmailSettings.eLowResources) <> 0

                    Dim lExt1 As Int32 = iTypeID

                    oSB.Length = 0
                    Select Case yReason
                        Case ProductionType.eColonists
                            oSB.AppendLine("Insufficient Colonists")
                            lExt1 = Int32.MinValue
                        Case ProductionType.eEnlisted
                            oSB.AppendLine("Insufficient Enlisted")
                            lExt1 = ObjectType.eEnlisted
                        Case ProductionType.eFood
                            oSB.AppendLine("Insufficient Food Supplies")
                            lExt1 = Int32.MinValue
                        Case ProductionType.eOfficers
                            oSB.AppendLine("Insufficient Officers")
                            lExt1 = ObjectType.eOfficers
                        Case ProductionType.ePowerCenter
                            oSB.AppendLine("Insufficient Power")
                            lExt1 = Int32.MinValue
                        Case ProductionType.eCredits
                            oSB.AppendLine("Insufficient Credits")
                            lExt1 = Int32.MinValue
                        Case ProductionType.eWareHouse
                            oSB.AppendLine("Insufficient Cargo Capacity (Colony-wide)")
                            lExt1 = Int32.MinValue
                        Case ProductionType.eProduction
                            oSB.AppendLine("Insufficient Hangar Capacity")
                            lExt1 = Int32.MinValue
                        Case Else
                            oSB.AppendLine("Insufficient Resources")

                            Dim oObj As Object = GetEpicaObject(lItemID, iTypeID)
                            If oObj Is Nothing = False Then
                                oSB.AppendLine("Not enough " & BytesToString(GetEpicaObjectName(iTypeID, oObj)))
                                If CType(oObj, Epica_GUID).ObjTypeID = ObjectType.eMineral Then
                                    If CType(oObj, Mineral).lAlloyTechID <> -1 Then
                                        lExt1 = ObjectType.eAlloyTech
                                    End If
                                End If
                            End If
                            oObj = Nothing
                    End Select
                    oSB.AppendLine()
                    If sColonyName <> "" Then
                        oSB.AppendLine("Colony: " & sColonyName)
                    End If
                    If iEnvirTypeID = ObjectType.ePlanet Then
                        Dim oP As Planet = GetEpicaPlanet(lEnvirID)
                        If oP Is Nothing = False Then
                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
                        End If
                    ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
                        Dim oS As SolarSystem = GetEpicaSystem(lEnvirID)
                        If oS Is Nothing = False Then
                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
                        End If
                    End If

                    If iProdTypeID = ObjectType.eEnlisted Then
                        oSB.AppendLine("Attempting to Build Enlisted")
                    ElseIf iProdTypeID = ObjectType.eOfficers Then
                        oSB.AppendLine("Attempting to Build Officers")
                    ElseIf iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eFacilityDef Then
                        Dim oUnit As Epica_Entity_Def = GetEpicaUnitDef(lProdID)
                        If oUnit Is Nothing = False Then
                            oSB.AppendLine("Attempting to Build " & BytesToString(oUnit.DefName))
                        ElseIf iProdTypeID = ObjectType.eUnitDef Then
                            oSB.AppendLine("Attempting to Build a Unit")
                        Else : oSB.AppendLine("Attempting to Build a Facility")
                        End If
                    Else
                        Select Case iProdTypeID
                            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHangarTech, ObjectType.eHullTech, ObjectType.eMineralTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
                                Dim oTech As Epica_Tech = GetTech(lProdID, iProdTypeID)
                                If oTech Is Nothing = False Then
                                    oSB.AppendLine("Attempting to Build a " & BytesToString(oTech.GetTechName))
                                Else : oSB.AppendLine("Attempting to Build a Component")
                                End If
                        End Select
                    End If

                    Dim oSBExtra As New System.Text.StringBuilder
                    oSBExtra.Length = 0
                    If bSendExt = True AndAlso lExt1 <> Int32.MinValue Then
                        oSBExtra.AppendLine()
                        oSBExtra.AppendLine("How would you like to respond?" & vbCrLf)
                        oSBExtra.AppendLine("Build/Gather <enter number>")
                    End If

                    sBody = oSB.ToString
                    sTitle = "Low Resources"

                    Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName), Nothing)
                    If oPC Is Nothing = False AndAlso bSendExt = True Then goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, iTypeID, lColonyID, lItemID, 0, 0, 0, oSBExtra.ToString)
                Case GlobalMessageCode.ePlayerAlert
                    'MsgCode, Type(1), EntityGUID(6), EnvirGUID(6), PlayerID(4), EnemyID(4), Optional Name(20)
                    Dim yType As Byte = yMsg(lPos) : lPos += 1
                    Dim lObjID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
                    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yMsg, lPos) : lPos += 2
                    lPos += 4       'playerid
                    Dim lEnemyID As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

                    'Now, check our type
                    Select Case yType
                        Case PlayerAlertType.eColonyLost, PlayerAlertType.eFacilityLost, PlayerAlertType.eUnitLost
                            If yType = PlayerAlertType.eColonyLost AndAlso (iInternalEmailSettings And eEmailSettings.eColonyLost) = 0 Then Return
                            If (yType = PlayerAlertType.eFacilityLost OrElse yType = PlayerAlertType.eUnitLost) AndAlso (iInternalEmailSettings And eEmailSettings.eFacilityLost) = 0 Then Return
                            'If yType = PlayerAlertType.eUnitLost AndAlso (iInternalEmailSettings And eEmailSettings.eUnitLost) = 0 Then Return

                            'Ok, for these, a name will be appended to the message
                            Dim yTemp(19) As Byte
                            Array.Copy(yMsg, lPos, yTemp, 0, 20)
                            lPos += 20
                            Dim sName As String = BytesToString(yTemp)

                            Dim lLocX As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                            Dim lLocZ As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

                            Dim lIdx As Int32 = CreateEnvirAlert(lEnvirID, iEnvirTypeID, yType)
                            If lIdx = -1 Then Return

                            muEnvirAlerts(lIdx).AddName(sName, lLocX, lLocZ, lEnemyID)

                            If muEnvirAlerts(lIdx).OriginalAlert = 0 Then
                                muEnvirAlerts(lIdx).OriginalAlert = GetDateAsNumber(Now)
                            End If
                        Case PlayerAlertType.eEngagedEnemy, PlayerAlertType.eUnderAttack, PlayerAlertType.eBuyOrderAccepted
                            Dim lIdx As Int32 = CreateEnvirAlert(lEnvirID, iEnvirTypeID, yType)
                            If lIdx = -1 Then Return
                            If muEnvirAlerts(lIdx).EmailSent = False Then
                                'Ok, let's construct our email
                                oSB.Length = 0

                                Dim oPlayer As Player = GetEpicaPlayer(lEnemyID)

                                Dim bSendExtEmail As Boolean = False
                                Dim oSBExtra As New System.Text.StringBuilder
                                oSBExtra.Length = 0

                                If yType = PlayerAlertType.eEngagedEnemy Then
                                    If (Me.iInternalEmailSettings And eEmailSettings.eEngagedEnemy) = 0 Then Return
                                    If oPlayer Is Nothing = False Then
                                        oSB.AppendLine("Our forces are reporting an engagement with the enemy! (" & BytesToString(oPlayer.PlayerName) & ")")
                                    Else : oSB.AppendLine("Our forces are reporting an engagement with the enemy!")
                                    End If
                                    sTitle = "Enemy Engaged"

                                    bSendExtEmail = (iEmailSettings And eEmailSettings.eEngagedEnemy) <> 0
                                    'ElseIf yType = PlayerAlertType.eBuyOrderAccepted Then
                                    '    If (Me.iInternalEmailSettings And eEmailSettings.eBuyOrderAccepted) = 0 Then Return
                                    '    If oPlayer Is Nothing = False Then
                                    '        oSB.AppendLine(BytesToString(oPlayer.PlayerName) & " has accepted one of our buy order.")
                                    '    Else : oSB.AppendLine("One of our buy orders have been accepted.")
                                    '    End If
                                    '    sTitle = "Buy Order Accepted"
                                    '    bSendExtEmail = (iEmailSettings And eEmailSettings.eBuyOrderAccepted) <> 0
                                Else
                                    If (Me.iInternalEmailSettings And eEmailSettings.eUnderAttack) = 0 Then Return
                                    If oPlayer Is Nothing = False Then
                                        oSB.AppendLine(BytesToString(oPlayer.PlayerName) & " is attacking our forces!")
                                    Else : oSB.AppendLine("Our forces are under attack!")
                                    End If
                                    sTitle = "We Are Under Attack"

                                    bSendExtEmail = (iEmailSettings And eEmailSettings.eUnderAttack) <> 0

                                    If bSendExtEmail = True Then
                                        oSBExtra.AppendLine()
                                        oSBExtra.AppendLine("How would you like to respond?" & vbCrLf)

                                        For X As Int32 = 0 To glUnitGroupUB
                                            If glUnitGroupIdx(X) <> -1 Then
                                                'AndAlso goUnitGroup(X).oOwner.ObjectID = Me.ObjectID Then
                                                Try
                                                    Dim oGroup As UnitGroup = goUnitGroup(X)
                                                    If oGroup.oOwner.ObjectID = Me.ObjectID Then
                                                        oSBExtra.AppendLine("Attack Using Battlegroup: " & BytesToString(oGroup.UnitGroupName))
                                                    End If
                                                Catch
                                                End Try
                                            End If
                                        Next X

                                        oSBExtra.AppendLine("Launch To Attack")
                                        oSBExtra.AppendLine("Attack Using All")
                                        oSBExtra.AppendLine("Attack Using Half")
                                        oSBExtra.AppendLine("Raise Full Invulnerability")

                                    End If
                                End If

                                If yType <> PlayerAlertType.eBuyOrderAccepted Then
                                    If iEnvirTypeID = ObjectType.ePlanet Then
                                        Dim oP As Planet = GetEpicaPlanet(lEnvirID)
                                        If oP Is Nothing = False Then
                                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
                                        End If
                                    ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
                                        Dim oS As SolarSystem = GetEpicaSystem(lEnvirID)
                                        If oS Is Nothing = False Then
                                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
                                        End If
                                    End If
                                Else
                                    oSB.AppendLine(goGTC.GetBuyOrderAcceptedEmailBody(lObjID))
                                End If
                                sBody = oSB.ToString

                                muEnvirAlerts(lIdx).EmailSent = True

                                Dim lLocX As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4
                                Dim lLocZ As Int32 = System.BitConverter.ToInt32(yMsg, lPos) : lPos += 4

                                Dim uAttach(0) As PlayerComm.WPAttachment
                                With uAttach(0)
                                    .AttachNumber = 0
                                    .EnvirID = lEnvirID
                                    .EnvirTypeID = iEnvirTypeID
                                    .LocX = lLocX
                                    .LocZ = lLocZ
                                    .sWPName = "Event Location"
                                End With
                                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName), uAttach)
                                If oPC Is Nothing = False Then
                                    'oPC.AddEmailAttachment(0, lEnvirID, iEnvirTypeID, lLocX, lLocZ, "Event Location")
                                    If bSendExtEmail = True Then

                                        'If iObjTypeID = ObjectType.eUnit Then
                                        '    Dim oUnit As Unit = GetEpicaUnit(lObjID)
                                        '    If oUnit Is Nothing = False Then
                                        '        lLocX = oUnit.LocX
                                        '        lLocZ = oUnit.LocZ
                                        '    End If
                                        'ElseIf iObjTypeID = ObjectType.eFacility Then
                                        '    Dim oFacility As Facility = GetEpicaFacility(lObjID)
                                        '    If oFacility Is Nothing = False Then
                                        '        lLocX = oFacility.LocX
                                        '        lLocZ = oFacility.LocZ
                                        '    End If
                                        'End If

                                        goMsgSys.SendOutboundEmail(oPC, Me, CShort(yType), yType, lEnemyID, lEnvirID, iEnvirTypeID, lLocX, lLocZ, oSBExtra.ToString)
                                    End If
                                End If

                            End If
                    End Select

                    'TODO: However, if the player has msg alerts set up for external devices, we need to send them now

                Case Else
                    'TODO: Queue the message up to send to the player when they come online... for now, return
            End Select
        ElseIf (iEmailSettings And eEmailSettings.eAllInternalEmails) <> 0 Then
        'Ok, if this is an AddObjectMessage of a PlayerComm type, then we will want to check the sender...
        If yMsg Is Nothing OrElse yMsg.Length < 2 Then Return
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yMsg, 0)
        If iMsgCode = GlobalMessageCode.eAddObjectCommand Then
            'ok, check the guid
            Dim lID As Int32 = System.BitConverter.ToInt32(yMsg, 2)
            Dim iTypeID As Int16 = System.BitConverter.ToInt16(yMsg, 6)

            If iTypeID = ObjectType.ePlayerComm Then


                Dim oPC As PlayerComm = Nothing
                For X As Int32 = 0 To EmailFolderUB
                    If EmailFolders(X) Is Nothing = False Then
                        For Y As Int32 = 0 To EmailFolders(X).PlayerMsgsUB
                            If EmailFolders(X).PlayerMsgsIdx(Y) > -1 Then
                                Dim oMsg As PlayerComm = EmailFolders(X).PlayerMsgs(Y)
                                If oMsg Is Nothing = False AndAlso oMsg.ObjectID = lID Then
                                    oPC = oMsg
                                    Exit For
                                End If
                            End If
                        Next Y
                        If oPC Is Nothing = False Then Exit For
                    End If
                Next X

                If oPC Is Nothing = False Then
                    Dim sSentBy As String = ""
                    Dim oTmp As Player = GetEpicaPlayer(oPC.SentByID)
                    If oTmp Is Nothing = False Then sSentBy = vbCrLf & "Sent By: " & oTmp.sPlayerNameProper
                    goMsgSys.SendOutboundEmail(oPC, Me, iMsgCode, ObjectType.ePlayerComm, oPC.SentByID, -1, -1, -1, -1, sSentBy & vbCrLf & "Max Response Length: 5000")
                End If

            End If

        End If
        End If
    End Sub

    Public Sub CrossPrimarySafeSendMsg(ByVal yData() As Byte)
        If lConnectedPrimaryID = -1 Then Return
        'ok, the player is connected... is it to me?
        If lConnectedPrimaryID = glServerID Then
            'yes, send directly to the player
            Try
                If oSocket Is Nothing = False Then
                    oSocket.SendData(yData)
                End If
            Catch ex As Exception
                LogEvent(LogEventType.Warning, "CrossPrimarySafeSendMsg (" & Me.ObjectID & "): " & ex.Message)
                If oSocket Is Nothing = False AndAlso oSocket.IsConnected = False Then oSocket = Nothing
            End Try
        Else
            'Ok, forward the msg to the operator encapsulated in a msg (ie: ForwardToPlayerAtPrimary) 
            Dim yForward(9 + yData.Length) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eForwardToPlayerAtPrimary).CopyTo(yForward, lPos) : lPos += 2
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(yData.Length).CopyTo(yForward, lPos) : lPos += 4
            yData.CopyTo(yForward, lPos) : lPos += yData.Length

            'Operator will receive it and forward to the appropriate primary
            'Receiving Primary will remove the encapsulation and send to the player

            goMsgSys.SendMsgToOperator(yForward)
        End If
    End Sub

    Private Function CreateEnvirAlert(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal yType As Byte) As Int32
        'Now, we store this specifically in an muEnvirAlert object...
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlEnvirAlertUB
            With muEnvirAlerts(X)
                If .EnvirID = lEnvirID AndAlso .EnvirTypeID = iEnvirTypeID AndAlso .AlertType = yType Then
                    lIdx = X
                    Exit For
                End If
            End With
        Next X

        If lIdx = -1 Then
            mlEnvirAlertUB += 1
            ReDim Preserve muEnvirAlerts(mlEnvirAlertUB)
            lIdx = mlEnvirAlertUB
            With muEnvirAlerts(lIdx)
                .EnvirID = lEnvirID
                .EnvirTypeID = iEnvirTypeID
                .AlertType = yType
                .lNameUB = -1
            End With
        End If
        Return lIdx
    End Function

    Public Sub CreateAndSendPlayerAlert(ByVal yType As PlayerAlertType, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lEnemyID As Int32, ByVal sName As String, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal sAlternateName As String)
        'MsgCode, Type(1), EntityGUID(6), EnvirGUID(6), PlayerID(4), EnemyID (4) , Name (20), LocX (4), LocZ (4)
        Dim yMsg(70) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAlert).CopyTo(yMsg, lPos) : lPos += 2
        yMsg(lPos) = yType : lPos += 1
        System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lEnemyID).CopyTo(yMsg, lPos) : lPos += 4

        If sName Is Nothing = False AndAlso sName <> "" Then
            Dim sTemp As String
            If sName.Length > 20 Then sTemp = sName.Substring(0, 20) Else sTemp = sName
            StringToBytes(sTemp).CopyTo(yMsg, lPos)
        End If
        lPos += 20

        System.BitConverter.GetBytes(lLocX).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lLocZ).CopyTo(yMsg, lPos) : lPos += 4

        If sAlternateName Is Nothing = False AndAlso sAlternateName <> "" Then
            If sAlternateName.Length > 20 Then sAlternateName = sAlternateName.Substring(0, 20)
            StringToBytes(sAlternateName).CopyTo(yMsg, lPos)
        End If
        lPos += 20

        Me.SendPlayerMessage(yMsg, True, 0)
    End Sub

    Public Sub CancelIronCurtain()
        'ok, remove the iron curtain (if one exists)
        CancelIronCurtainEvents(Me.ObjectID)
        If Me.lIronCurtainPlanet < 1 Then Me.lIronCurtainPlanet = Me.lStartedEnvirID
        If Me.lIronCurtainPlanet > 0 Then
            Dim oPlanet As Planet = GetEpicaPlanet(Me.lIronCurtainPlanet)
            If oPlanet Is Nothing = False Then
                Dim yMsg(10) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetIronCurtain).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(oPlanet.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = 0 : lPos += 1
                'oPlanet.oDomain.DomainSocket.SendData(yMsg)
                goMsgSys.BroadcastToDomains(yMsg)
            End If
        End If
    End Sub

    Public Sub FinalizeEnvirAlerts()
        'Ok, go through all environment alerts
        For X As Int32 = 0 To mlEnvirAlertUB
            If muEnvirAlerts(X).EmailSent = False Then
                With muEnvirAlerts(X)
                    .EmailSent = True
                    Dim sBody As String
                    Dim sTitle As String

                    Dim oSB As New System.Text.StringBuilder

                    If .EnvirTypeID = ObjectType.ePlanet Then
                        Dim oP As Planet = GetEpicaPlanet(.EnvirID)
                        If oP Is Nothing = False Then
                            oSB.AppendLine("Location: " & BytesToString(oP.PlanetName) & " (Planet side)")
                        End If
                    ElseIf .EnvirTypeID = ObjectType.eSolarSystem Then
                        Dim oS As SolarSystem = GetEpicaSystem(.EnvirID)
                        If oS Is Nothing = False Then
                            oSB.AppendLine("Location: " & BytesToString(oS.SystemName) & " (System)")
                        End If
                    End If

                    Dim bIncludeKilledBy As Boolean = True
                    Select Case .AlertType
                        Case PlayerAlertType.eColonyLost
                            oSB.AppendLine("The following Colonies were lost:")
                            sTitle = "Colony Lost"
                            bIncludeKilledBy = False
                        Case PlayerAlertType.eFacilityLost
                            oSB.AppendLine("The following Facilities were lost:")
                            sTitle = "Facility Lost"
                        Case PlayerAlertType.eUnitLost
                            oSB.AppendLine("The following Units were lost:")
                            sTitle = "Unit Lost"
                        Case Else
                            Continue For
                    End Select

                    oSB.AppendLine()
                    oSB.Append(.GetNameList(bIncludeKilledBy))

                    sBody = oSB.ToString

                    Dim lAttachUB As Int32 = Math.Min(.lNameUB, 254)
                    Dim uAttach() As PlayerComm.WPAttachment = Nothing
                    If lAttachUB > -1 Then
                        ReDim uAttach(lAttachUB)
                        Dim lAttachIdx As Int32 = -1
                        For Y As Int32 = 0 To Math.Min(.lNameUB, 254)
                            Dim sName As String = .GetNameItem(Y)
                            If sName <> "" Then
                                Dim ptLoc As Point = .GetNameLoc(Y)
                                If ptLoc <> Point.Empty Then
                                    'oPC.AddEmailAttachment(CByte(Y + 1), .EnvirID, .EnvirTypeID, ptLoc.X, ptLoc.Y, sName)
                                    lAttachIdx += 1
                                    uAttach(lAttachIdx).AttachNumber = CByte(lAttachIdx)
                                    uAttach(lAttachIdx).EnvirID = .EnvirID
                                    uAttach(lAttachIdx).EnvirTypeID = .EnvirTypeID
                                    uAttach(lAttachIdx).LocX = ptLoc.X
                                    uAttach(lAttachIdx).LocZ = ptLoc.Y
                                    uAttach(lAttachIdx).sWPName = sName
                                End If
                            End If
                        Next Y
                        ReDim Preserve uAttach(lAttachIdx)
                    End If

                    Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, sTitle, Me.ObjectID, .OriginalAlert, False, BytesToString(Me.PlayerName), uAttach)
                    If oPC Is Nothing = False Then
                        oPC.SaveObject(Me.ObjectID)
                    End If
                End With
            End If
        Next X

    End Sub

    Private Structure EnvirAlert
        Public EnvirID As Int32
        Public EnvirTypeID As Int16
        Public AlertType As Byte

        Public EmailSent As Boolean

        Public OriginalAlert As Int32

        Private msNames() As String
        Private mlLocX() As Int32
        Private mlLocZ() As Int32
        Private mlKilledBy() As Int32
        Public lNameUB As Int32

        Public Sub AddName(ByVal sName As String, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal lKilledByID As Int32)
            If msNames Is Nothing Then
                ReDim msNames(-1)
                lNameUB = -1
            End If

            lNameUB += 1
            ReDim Preserve msNames(lNameUB)
            ReDim Preserve mlLocX(lNameUB)
            ReDim Preserve mlLocZ(lNameUB)
            ReDim Preserve mlKilledBy(lNameUB)
            msNames(lNameUB) = Mid$(sName, 1, 20)
            mlLocX(lNameUB) = lLocX
            mlLocZ(lNameUB) = lLocZ
            mlKilledBy(lNameUB) = lKilledByID
        End Sub

        Public Function GetNameList(ByVal bIncludeKilledBy As Boolean) As String
            Dim oSB As New System.Text.StringBuilder
            For X As Int32 = 0 To lNameUB
                Dim sVal As String = msNames(X)
                If bIncludeKilledBy = True Then
                    Dim oKilledBy As Player = GetEpicaPlayer(mlKilledBy(X))
                    If oKilledBy Is Nothing = False Then
                        sVal &= " (" & oKilledBy.sPlayerNameProper & ")"
                        'Else
                        '   sVal &= " (Unknown)"
                    End If
                End If
                oSB.AppendLine(sVal)
            Next X
            Return oSB.ToString
        End Function

        Public Function GetNameLoc(ByVal lIndex As Int32) As Point
            Try
                Return New Point(mlLocX(lIndex), mlLocZ(lIndex))
            Catch ex As Exception
                Return Point.Empty
            End Try
        End Function
        Public Function GetNameItem(ByVal lIndex As Int32) As String
            Try
                Return msNames(lIndex)
            Catch ex As Exception
                Return ""
            End Try
        End Function
    End Structure

    Public Sub AddEmailAttachment(ByVal PC_ID As Int32, ByVal PCF_ID As Int32, ByVal AttachNumber As Byte, ByVal EnvirID As Int32, ByVal EnvirTypeID As Int16, ByVal LocX As Int32, ByVal LocZ As Int32, ByVal sName As String)
        For X As Int32 = 0 To EmailFolderUB
            If EmailFolderIdx(X) = PCF_ID Then
                With EmailFolders(X)
                    For Y As Int32 = 0 To .PlayerMsgsUB
                        If .PlayerMsgsIdx(Y) = PC_ID Then
                            .PlayerMsgs(Y).AddEmailAttachment(AttachNumber, EnvirID, EnvirTypeID, LocX, LocZ, sName)
                            Exit For
                        End If
                    Next
                End With
                Exit For
            End If
        Next X
    End Sub
#End Region

    '#Region "  Chat Room Management  "
    '	'Player Chat Rooms
    '	Public ChatRooms() As Int32
    '	Public ChatRoomAlias() As String
    '	Public ChatRoomUB As Int32 = -1

    '	Public Sub AssociateChatRoom(ByVal lChatRoomID As Int32)
    '		Dim X As Int32
    '		Dim lIdx As Int32 = -1

    '		For X = 0 To ChatRoomUB
    '			If lIdx = -1 AndAlso ChatRooms(X) = -1 Then
    '				lIdx = X
    '			ElseIf ChatRooms(X) = lChatRoomID Then
    '				Exit Sub
    '			End If
    '		Next X

    '		If lIdx = -1 Then
    '			ChatRoomUB += 1
    '			ReDim Preserve ChatRooms(ChatRoomUB)
    '			ReDim Preserve ChatRoomAlias(ChatRoomUB)
    '			lIdx = ChatRoomUB
    '		End If

    '		ChatRooms(lIdx) = lChatRoomID
    '	End Sub

    '	Public Sub RemoveChatRoom(ByVal lChatRoomID As Int32)
    '		Dim X As Int32

    '		For X = 0 To ChatRoomUB
    '			If ChatRooms(X) = lChatRoomID Then
    '				ChatRooms(X) = -1
    '			End If
    '		Next X
    '	End Sub

    '	Public Sub AddChatRoomAlias(ByVal lChatRoomID As Int32, ByVal sAlias As String)
    '		Dim X As Int32

    '		For X = 0 To ChatRoomUB
    '			If ChatRooms(X) = lChatRoomID Then
    '				ChatRoomAlias(X) = UCase$(sAlias)
    '				Exit For
    '			End If
    '		Next X
    '    End Sub

    '    Public Function InChatRoom(ByVal lIdx As Int32) As Boolean
    '        For X As Int32 = 0 To ChatRoomUB
    '            If ChatRooms(X) = lIdx Then Return True
    '        Next X
    '        Return False
    '    End Function

    '    Public Function GetChatRoomIndex(ByVal lID As Int32, ByVal sChannel As String) As Int32
    '        sChannel = sChannel.ToUpper
    '        For X As Int32 = 0 To ChatRoomUB
    '            If ChatRooms(X) > -1 Then
    '                With guChatRooms(ChatRooms(X))
    '                    If .lID = lID OrElse .sRoomName.ToUpper = sChannel Then
    '                        Return ChatRooms(X)
    '                    End If
    '                End With
    '            End If
    '        Next X
    '        Return -1
    '    End Function
    '#End Region

#Region "  Player Scores  "
    Public yPlayerTitle As Byte = 0
    Public yExTitleNew As Byte = 0          'title the player will be once the 7 day ex-rank duration ends
    Public yCustomTitle As Byte = 0
    Public lCustomTitlePermission As Int32 = 0
    Public dtExTitleEnd As Date = Date.MinValue

    'Last Global Scores Reported...
    Public lLGTechScore As Int32 = 0
    Public lLGDiplomacyScore As Int32 = 0
    Public lLGPopulationScore As Int32 = 0
    Public lLGMilitaryScore As Int32 = 0
    Public lLGProductionScore As Int32 = 0
    Public lLGWealthScore As Int32 = 0
    Public lLGTotalScore As Int32 = 0
    Public lLastGlobalRequestTurnIn As Int32 = 0
    Public lTotalSenateVoteStrength As Int32 = 0

    Public ReadOnly Property TechnologyScore() As Int32
        Get
            Dim lResult As Int32 = 0
            For X As Int32 = 0 To mlAlloyUB
                If mlAlloyIdx(X) <> -1 AndAlso moAlloy(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moAlloy(X).TechnologyScore
            Next X
            For X As Int32 = 0 To mlArmorUB
                If mlArmorIdx(X) <> -1 AndAlso moArmor(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moArmor(X).TechnologyScore
            Next X
            For X As Int32 = 0 To mlEngineUB
                If mlEngineIdx(X) <> -1 AndAlso moEngine(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moEngine(X).TechnologyScore
            Next X
            For X As Int32 = 0 To mlHullUB
                If mlHullIdx(X) <> -1 AndAlso moHull(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moHull(X).TechnologyScore
            Next X
            For X As Int32 = 0 To mlPrototypeUB
                If mlPrototypeIdx(X) <> -1 AndAlso moPrototype(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moPrototype(X).TechnologyScore
            Next X
            For X As Int32 = 0 To mlRadarUB
                If mlRadarIdx(X) <> -1 AndAlso moRadar(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moRadar(X).TechnologyScore
            Next X
            For X As Int32 = 0 To mlShieldUB
                If mlShieldIdx(X) <> -1 AndAlso moShield(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moShield(X).TechnologyScore
            Next X
            For X As Int32 = 0 To mlWeaponUB
                If mlWeaponIdx(X) <> -1 AndAlso moWeapon(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += moWeapon(X).TechnologyScore
            Next X
            For X As Int32 = 0 To oSpecials.mlSpecialTechUB
                If oSpecials.mlSpecialTechIdx(X) <> -1 AndAlso oSpecials.moSpecialTech(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then lResult += oSpecials.moSpecialTech(X).TechnologyScore
            Next X
            Return lResult \ 10 '100
        End Get
    End Property

    Public ReadOnly Property DiplomacyScore() As Int32
        Get
            Dim lResult As Int32 = 0

            For X As Int32 = 0 To Me.PlayerRelUB
                If Me.mlPlayerRelIdx(X) <> -1 Then
                    lResult += (CInt(moPlayerRels(X).WithThisScore) - 40I)
                End If
            Next X

            If lGuildID <> -1 AndAlso Me.oGuild Is Nothing = False Then
                If (oGuild.lBaseGuildFlags And elGuildFlags.RequirePeaceBetweenMembers) <> 0 Then
                    lResult += (oGuild.GetApprovedMemberCount() * 1000)
                Else : lResult += 1000
                End If
            End If

            Return lResult * 20     'added the *20
        End Get
    End Property

    'Military Score needs to be managed on its own (in AddUnit and Entity.DeleteEntity)
    Public lMilitaryScore As Int32 = 0

    Public ReadOnly Property PopulationScore() As Int32
        Get
            Dim lResult As Int32 = 0
            For X As Int32 = 0 To mlColonyUB
                If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                    Dim oColony As Colony = goColony(mlColonyIdx(X))
                    If oColony Is Nothing = False Then lResult += (oColony.Population \ 10000)
                End If
            Next X
            Return lResult * 10     'added the x10
        End Get
    End Property

    Public ReadOnly Property ProductionScore() As Int32
        Get
            Dim lResult As Int32 = 0
            For X As Int32 = 0 To mlColonyUB
                If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                    Dim oColony As Colony = goColony(mlColonyIdx(X))
                    If oColony Is Nothing = False Then
                        With oColony
                            For Y As Int32 = 0 To .ChildrenUB
                                If .lChildrenIdx(Y) <> -1 AndAlso (.oChildren(Y).yProductionType And ProductionType.eProduction) <> 0 Then
                                    'lResult += .oChildren(Y).mlProdPoints
                                    lResult += .oChildren(Y).EntityDef.ProdFactor
                                End If
                            Next Y
                        End With
                    End If
                End If
            Next X
            Return lResult \ 5          'added the \5
        End Get
    End Property

    Public ReadOnly Property WealthScore() As Int32
        Get
            'Dim blCashFlow As Int64 = oBudget.GetCashFlow
            'Dim blVal As Int64 = (Me.blCredits \ (blCashFlow * 17280)) * 10000
            'If Me.blCredits < 0 AndAlso blCashFlow < 0 Then blVal = -(Math.Abs(blVal))

            If blCredits < 0 Then Return 0

            Dim dVal As Double = Math.Pow(blCredits, 0.3F) * 3      'added the x3
            If dVal > Int32.MaxValue Then Return Int32.MaxValue
            If dVal < Int32.MinValue Then Return Int32.MinValue
            Return CInt(dVal)
        End Get
    End Property

    Public ReadOnly Property TotalScore() As Int32
        Get
            Return (TechnologyScore + DiplomacyScore + PopulationScore + ProductionScore + WealthScore + (lMilitaryScore \ 50)) \ 6
        End Get
    End Property

    Private myLastRetestTitleVal As Byte = 0
    Public Sub ReTestTitle()
        Dim blPop As Int64 = 0

        If Me.AccountStatus <> AccountStatusType.eActiveAccount OrElse Me.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
            For X As Int32 = 0 To mlColonyUB
                If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                    Dim oColony As Colony = goColony(mlColonyIdx(X))
                    If oColony Is Nothing Then Continue For
                    blPop += oColony.Population
                End If
            Next X
            If blPop > blMaxPopulation Then blMaxPopulation = blPop
            Return
        End If

        'Ok, the layer by default has Magistrate
        Dim yLevel As Byte = 0

        'Dim bSpaceColony As Boolean = False

        Dim lOwnedPlanets As Int32 = 0
        Dim lPlanetColonies As Int32 = 0

        Dim lSystemID() As Int32 = Nothing      'ID of the system (for grouping)
        Dim lSystemPCnt() As Int32 = Nothing    'Total planets in the system
        Dim lSystemOwned() As Int32 = Nothing   'Total planets in the system that I own
        Dim oSystem() As SolarSystem = Nothing  'pointer to the system
        Dim lSystemUB As Int32 = -1

        Dim bNeedToTestCustomTitle As Boolean = False

        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing Then Continue For

                blPop += oColony.Population
                If CType(oColony.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                    'bSpaceColony = True
                Else

                    If CType(oColony.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.ePlanet Then Continue For
                    lPlanetColonies += 1

                    Dim oPlanet As Planet = CType(oColony.ParentObject, Planet)
                    Dim bGood As Boolean = True
                    Dim lMyPop As Int32 = oColony.Population

                    If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then

                        Dim lIdx As Int32 = -1
                        For Y As Int32 = 0 To lSystemUB
                            If oPlanet.ParentSystem.ObjectID = lSystemID(Y) Then
                                lIdx = Y
                                Exit For
                            End If
                        Next Y
                        If lIdx = -1 Then
                            lSystemUB += 1
                            ReDim Preserve lSystemID(lSystemUB)
                            ReDim Preserve lSystemPCnt(lSystemUB)
                            ReDim Preserve lSystemOwned(lSystemUB)
                            ReDim Preserve oSystem(lSystemUB)
                            lSystemID(lSystemUB) = oPlanet.ParentSystem.ObjectID
                            lSystemPCnt(lSystemUB) = oPlanet.ParentSystem.mlPlanetUB + 1
                            lSystemOwned(lSystemUB) = 0
                            oSystem(lSystemUB) = oPlanet.ParentSystem
                            lIdx = lSystemUB
                        End If

                        Dim blTotalPop As Int64 = lMyPop
                        For Y As Int32 = 0 To oPlanet.lColonysHereUB
                            If oPlanet.lColonysHereIdx(Y) <> -1 AndAlso oPlanet.lColonysHereIdx(Y) <> mlColonyIdx(X) Then
                                If glColonyIdx(oPlanet.lColonysHereIdx(Y)) <> -1 Then
                                    Dim oTmpColony As Colony = goColony(oPlanet.lColonysHereIdx(Y))
                                    If oTmpColony Is Nothing = False Then
                                        If oTmpColony.ParentObject Is Nothing = False Then
                                            With CType(oTmpColony.ParentObject, Epica_GUID)
                                                If .ObjectID = oPlanet.ObjectID AndAlso .ObjTypeID = oPlanet.ObjTypeID Then
                                                    blTotalPop += oTmpColony.Population
                                                Else
                                                    oPlanet.lColonysHereIdx(Y) = -1
                                                End If
                                            End With
                                        End If
                                    End If
                                End If
                            End If
                        Next Y

                        If blTotalPop = 0 Then
                            bGood = False
                        Else
                            bGood = CSng(lMyPop / blTotalPop) > 0.75F
                            If oPlanet.OwnerID = Me.ObjectID AndAlso CSng(lMyPop / blTotalPop) > 0.7F Then bGood = True
                            oColony.GovScore = CByte(Math.Floor((lMyPop / blTotalPop) * 100.0F))
                        End If

                        If bGood = True Then

                            If oPlanet.OwnerID <> Me.ObjectID AndAlso yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso oPlanet.ParentSystem Is Nothing = False AndAlso oPlanet.ParentSystem.SystemType <> SolarSystem.elSystemType.RespawnSystem AndAlso oPlanet.ParentSystem.SystemType <> SolarSystem.elSystemType.SpawnSystem Then
                                Dim yGNS(118) As Byte
                                Dim lPos As Int32 = 0
                                System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
                                yGNS(lPos) = NewsItemType.ePlanetControlShifts : lPos += 1
                                oPlanet.ParentSystem.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                                oPlanet.PlanetName.CopyTo(yGNS, lPos) : lPos += 20
                                oPlanet.ParentSystem.SystemName.CopyTo(yGNS, lPos) : lPos += 20
                                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
                                System.BitConverter.GetBytes(oPlanet.OwnerID).CopyTo(yGNS, lPos) : lPos += 4

                                Me.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                                yGNS(lPos) = Me.yGender : lPos += 1
                                Me.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

                                If oPlanet.OwnerID > 0 Then
                                    Dim oTmpPlayer As Player = GetEpicaPlayer(oPlanet.OwnerID)
                                    If oTmpPlayer Is Nothing = False Then
                                        oTmpPlayer.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                                        yGNS(lPos) = oTmpPlayer.yGender : lPos += 1
                                    Else : lPos += 21
                                    End If
                                Else : lPos += 21
                                End If

                                goMsgSys.SendToEmailSrvr(yGNS)
                            End If

                            oPlanet.OwnerID = Me.ObjectID
                            oPlanet.QueueMeToSave()
                            bNeedToTestCustomTitle = True

                            'Now, send the ownership change to the Operator
                            Dim yOpMsg(9) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlanetOwnership).CopyTo(yOpMsg, 0)
                            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yOpMsg, 2)
                            System.BitConverter.GetBytes(oPlanet.ObjectID).CopyTo(yOpMsg, 6)
                            goMsgSys.SendMsgToOperator(yOpMsg)

                            lOwnedPlanets += 1
                            lSystemOwned(lIdx) += 1
                        ElseIf oPlanet.OwnerID = Me.ObjectID Then
                            lOwnedPlanets += 1
                            lSystemOwned(lIdx) += 1
                        End If

                    End If
                End If
            End If
        Next X

        If blPop > 200000 Then yLevel = 1
        If yLevel = 0 AndAlso yPlayerTitle = 1 AndAlso blPop > 190000 Then yLevel = 1
        If blPop > 500000 Then yLevel = 2 ' AndAlso bSpaceColony = True Then yLevel = 2
        If yLevel = 1 AndAlso yPlayerTitle = 2 AndAlso blPop > 490000 Then yLevel = 2
        If lOwnedPlanets > 0 AndAlso lPlanetColonies > 1 Then yLevel = 3 'AndAlso bSpaceColony = True 
        If lOwnedPlanets > 1 AndAlso lPlanetColonies > 2 Then yLevel = 4 'AndAlso bSpaceColony = True 
        If blPop > blMaxPopulation Then blMaxPopulation = blPop
        blCurrentPopulation = blPop

        'Dim lPrevOwnerID As Int32 = -1
        lTotalSenateVoteStrength = lOwnedPlanets

        If lOwnedPlanets > 0 Then
            For X As Int32 = 0 To lSystemUB
                If lSystemOwned(X) = lSystemPCnt(X) Then
                    If lOwnedPlanets > lSystemOwned(X) Then
                        yLevel = 6
                        mlPrevOwnerID = oSystem(X).OwnerID
                        Exit For
                    Else
                        yLevel = 5
                        mlPrevOwnerID = oSystem(X).OwnerID
                    End If
                    If mlPrevOwnerID <> Me.ObjectID Then
                        oSystem(X).OwnerID = Me.ObjectID
                        oSystem(X).QueueMeToSave()
                    End If
                Else
                    Dim fTemp As Single = CSng(lSystemOwned(X) / lSystemPCnt(X))
                    If fTemp > 0.5F Then
                        yLevel = 5
                        mlPrevOwnerID = oSystem(X).OwnerID
                    End If
                End If
            Next X
        End If

        Dim yTestTitle As Byte = yExTitleNew        'yExTitleNew should NEVER be Ex-Title anything
        If (yExTitleNew And Player.PlayerRank.ExRankShift) <> 0 Then
            yExTitleNew = yExTitleNew Xor Player.PlayerRank.ExRankShift
            yTestTitle = yTestTitle Xor Player.PlayerRank.ExRankShift
        End If

        If yTestTitle <> yLevel Then
            'Ok, we have been either promoted or demoted...
            yExTitleNew = yLevel        'set our extitlenew
            If yLevel > yTestTitle Then
                'we have been promoted... ok, determine if we are ex-ranked
                Dim bDoCelebration As Boolean = True
                If (yPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 Then
                    'we are ex-ranked... what were we?
                    Dim yTmp As Byte = yPlayerTitle Xor Player.PlayerRank.ExRankShift
                    'is our ex-rank > new rank?
                    If yTmp > yLevel Then
                        'ok, we do not do a celebration
                        bDoCelebration = False
                    End If
                End If

                If bDoCelebration = True Then

                    If yCustomTitle = yPlayerTitle Then yCustomTitle = yExTitleNew

                    yPlayerTitle = yExTitleNew
                    dtExTitleEnd = Date.MinValue

                    lCelebrationEnds = glCurrentCycle + 2592000
                    Dim oPC As PlayerComm = Nothing
                    If (iInternalEmailSettings And eEmailSettings.eTitleChange) <> 0 Then
                        Dim oSB As New System.Text.StringBuilder
                        oSB.AppendLine(Me.sPlayerNameProper & ",")
                        oSB.AppendLine()
                        oSB.AppendLine("  In honor of your recent title recognition, festivities have begun at all major colonies! The festivities will last for 24 hours. This is only good news for our civilization and we are honored to be alive during your reign.")
                        oSB.AppendLine()
                        oSB.AppendLine()
                        oSB.AppendLine("Your Loyal Citizens and Servants")
                        Dim sSubject As String = "Title Promotion to " & GetPlayerTitle(yLevel, Me.yGender <> 2) & " from " & GetPlayerTitle(yTestTitle, Me.yGender <> 2) & "."
                        oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, sSubject, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                    End If
                    If lConnectedPrimaryID > -1 OrElse HasOnlineAliases(AliasingRights.eViewEmail) = True Then
                        If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)

                        Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                        CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
                    End If

                    'Now, submit our GNS msg...
                    Dim oHomePlanet As Planet = GetEpicaPlanet(lStartedEnvirID)
                    If oHomePlanet Is Nothing = False AndAlso yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
                        Dim yGNS(99) As Byte
                        Dim lPos As Int32 = 0

                        System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2

                        If yLevel <= PlayerRank.Overseer Then
                            yGNS(lPos) = NewsItemType.eTitlePromotionLow : lPos += 1
                        ElseIf yLevel <= PlayerRank.Baron Then
                            yGNS(lPos) = NewsItemType.eTitlePromotionMed : lPos += 1
                        Else : yGNS(lPos) = NewsItemType.eTitlePromotionHi : lPos += 1
                        End If
                        'yGNS(lPos) = NewsItemType.eTitlePromotion : lPos += 1

                        If yLevel <= PlayerRank.Baron Then
                            'System lvl
                            oHomePlanet.ParentSystem.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                        Else
                            'galactic lvl
                            oHomePlanet.ParentSystem.ParentGalaxy.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                        End If
                        yGNS(lPos) = yLevel : lPos += 1
                        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
                        System.BitConverter.GetBytes(mlPrevOwnerID).CopyTo(yGNS, lPos) : lPos += 4

                        oHomePlanet.PlanetName.CopyTo(yGNS, lPos) : lPos += 20

                        Me.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                        yGNS(lPos) = Me.yGender : lPos += 1
                        Me.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

                        If mlPrevOwnerID > 0 Then
                            Dim oTmpPlayer As Player = GetEpicaPlayer(mlPrevOwnerID)
                            If oTmpPlayer Is Nothing = False Then
                                oTmpPlayer.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                                yGNS(lPos) = oTmpPlayer.yGender : lPos += 1
                            Else : lPos += 21
                            End If
                        Else : lPos += 21
                        End If

                        goMsgSys.SendToEmailSrvr(yGNS)
                    End If


                End If
            Else
                'ok, demotion
                If (yPlayerTitle And Player.PlayerRank.ExRankShift) = 0 Then
                    If yPlayerTitle = yCustomTitle Then yCustomTitle = yPlayerTitle Or Player.PlayerRank.ExRankShift
                    yPlayerTitle = yPlayerTitle Or Player.PlayerRank.ExRankShift
                    dtExTitleEnd = Now.AddDays(7)

                    Dim oPC As PlayerComm = Nothing
                    If (iInternalEmailSettings And eEmailSettings.eTitleChange) <> 0 Then
                        Dim oSB As New System.Text.StringBuilder
                        oSB.AppendLine(Me.sPlayerNameProper & ",")
                        oSB.AppendLine()
                        oSB.AppendLine("  The Galactic Senate has deemed it necessary to strip you of your title of " & GetPlayerTitle(yTestTitle, Me.yGender <> 2) & " because you no longer meet the criteria. Because of this demotion, the Galactic Senate will refer to you as " & GetPlayerTitle(yPlayerTitle, Me.yGender <> 2) & " for seven days. During this time, you will be unable to be a member of a faction or have faction members.")
                        oSB.AppendLine()
                        oSB.AppendLine()
                        oSB.AppendLine("Galactic Senate Title Administration")
                        Dim sSubject As String = "Title Demotion to " & GetPlayerTitle(yLevel, Me.yGender <> 2) & " from " & GetPlayerTitle(yTestTitle, Me.yGender <> 2) & "."
                        oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, sSubject, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                    End If

                    Dim yNews(54) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yNews, lPos) : lPos += 2
                    yNews(lPos) = 24 : lPos += 1
                    System.BitConverter.GetBytes(1I).CopyTo(yNews, lPos) : lPos += 4
                    System.BitConverter.GetBytes(1S).CopyTo(yNews, lPos) : lPos += 2
                    yNews(lPos) = yLevel : lPos += 1
                    System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yNews, lPos) : lPos += 4
                    Me.PlayerName.CopyTo(yNews, lPos) : lPos += 20
                    yNews(lPos) = Me.yGender : lPos += 1
                    Me.EmpireName.CopyTo(yNews, lPos) : lPos += 20
                    goMsgSys.SendToEmailSrvr(yNews)
                    

                    If lConnectedPrimaryID > -1 OrElse Me.HasOnlineAliases(AliasingRights.eViewEmail) = True Then
                        If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)

                        Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                        CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
                    End If
                End If
            End If

            Dim yMsg(6) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerTitleChange).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
            yMsg(6) = yPlayerTitle
            Me.SendPlayerMessage(yMsg, False, 0)

            goMsgSys.SendMsgToOperator(yMsg)

            'Reverify our slot configurations...
            ReverifySlots()
        ElseIf (yPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 Then
            If Now > dtExTitleEnd Then
                dtExTitleEnd = Date.MinValue
                If yCustomTitle = yPlayerTitle Then yCustomTitle = yExTitleNew
                yPlayerTitle = yExTitleNew
                Dim yMsg(6) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerTitleChange).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                yMsg(6) = yPlayerTitle
                Me.SendPlayerMessage(yMsg, False, 0)

                Dim oPC As PlayerComm = Nothing
                If (iInternalEmailSettings And eEmailSettings.eTitleChange) <> 0 Then
                    Dim oSB As New System.Text.StringBuilder
                    oSB.AppendLine(Me.sPlayerNameProper & ",")
                    oSB.AppendLine()
                    oSB.AppendLine("  The seven day period for your recent title demotion has ended.")
                    oSB.AppendLine()
                    oSB.AppendLine()
                    oSB.AppendLine("Galactic Senate Title Administration")
                    Dim sSubject As String = "Important News About Your Title"
                    oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, sSubject, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                End If
                If lConnectedPrimaryID > -1 OrElse Me.HasOnlineAliases(AliasingRights.eViewEmail) = True Then
                    If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                    CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
                End If
            End If
        End If

        'If yLevel <> myLastRetestTitleVal Then
        '    'Ok, the last time we retested title, I had a different value then now
        '    Dim yMsg(11) As Byte
        '    System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
        '    System.BitConverter.GetBytes(ObjectID).CopyTo(yMsg, 2)
        '    System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.ePlayerTitle).CopyTo(yMsg, 6)
        '    System.BitConverter.GetBytes(CInt(yLevel)).CopyTo(yMsg, 8)

        '    myLastRetestTitleVal = yLevel

        '    goMsgSys.SendMsgToOperator(yMsg)
        'End If
        'myLastRetestTitleVal = yLevel

        If bNeedToTestCustomTitle = True Then
            TestCustomTitlePermissions_Allies()
        End If

    End Sub

    Private mlPrevOwnerID As Int32 = -1
    Public Sub OperatorSetTitle(ByVal yLevel As Byte, ByVal yNewExTitleVal As Byte)

        If Me.InMyDomain = False Then
            If yCustomTitle = yPlayerTitle Then yCustomTitle = yLevel
            yPlayerTitle = yLevel
            yExTitleNew = yNewExTitleVal
        Else
            Dim yTestTitle As Byte = yPlayerTitle
            If (yTestTitle And PlayerRank.ExRankShift) <> 0 Then yTestTitle = yTestTitle Xor PlayerRank.ExRankShift

            If yTestTitle <> yLevel Then
                'Ok, we have been either promoted or demoted...
                yExTitleNew = yLevel        'set our extitlenew
                If yLevel > yTestTitle Then
                    'we have been promoted... ok, determine if we are ex-ranked
                    Dim bDoCelebration As Boolean = True
                    If (yPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 Then
                        'we are ex-ranked... what were we?
                        Dim yTmp As Byte = yPlayerTitle Xor Player.PlayerRank.ExRankShift
                        'is our ex-rank > new rank?
                        If yTmp > yLevel Then
                            'ok, we do not do a celebration
                            bDoCelebration = False
                        End If
                    End If

                    If bDoCelebration = True Then

                        If yCustomTitle = yPlayerTitle Then yCustomTitle = yExTitleNew

                        yPlayerTitle = yExTitleNew
                        dtExTitleEnd = Date.MinValue

                        lCelebrationEnds = glCurrentCycle + 2592000
                        Dim oPC As PlayerComm = Nothing
                        If (iInternalEmailSettings And eEmailSettings.eTitleChange) <> 0 Then
                            Dim oSB As New System.Text.StringBuilder
                            oSB.AppendLine(Me.sPlayerNameProper & ",")
                            oSB.AppendLine()
                            oSB.AppendLine("  In honor of your recent title recognition, festivities have begun at all major colonies! The festivities will last for 24 hours. This is only good news for our civilization and we are honored to be alive during your reign.")
                            oSB.AppendLine()
                            oSB.AppendLine()
                            oSB.AppendLine("Your Loyal Citizens and Servants")
                            Dim sSubject As String = "Title Promotion to " & GetPlayerTitle(yLevel, Me.yGender <> 2) & " from " & GetPlayerTitle(yTestTitle, Me.yGender <> 2) & "."
                            oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, sSubject, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                        End If
                        If lConnectedPrimaryID > -1 OrElse HasOnlineAliases(AliasingRights.eViewEmail) = True Then
                            If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)

                            Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                            CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
                        End If

                        'Now, submit our GNS msg...
                        Dim oHomePlanet As Planet = GetEpicaPlanet(lStartedEnvirID)
                        If oHomePlanet Is Nothing = False AndAlso yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
                            Dim yGNS(99) As Byte
                            Dim lPos As Int32 = 0

                            System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2

                            If yLevel <= PlayerRank.Overseer Then
                                yGNS(lPos) = NewsItemType.eTitlePromotionLow : lPos += 1
                            ElseIf yLevel <= PlayerRank.Baron Then
                                yGNS(lPos) = NewsItemType.eTitlePromotionMed : lPos += 1
                            Else : yGNS(lPos) = NewsItemType.eTitlePromotionHi : lPos += 1
                            End If
                            'yGNS(lPos) = NewsItemType.eTitlePromotion : lPos += 1

                            If yLevel <= PlayerRank.Baron Then
                                'System lvl
                                oHomePlanet.ParentSystem.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                            Else
                                'galactic lvl
                                oHomePlanet.ParentSystem.ParentGalaxy.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                            End If
                            yGNS(lPos) = yLevel : lPos += 1
                            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
                            System.BitConverter.GetBytes(mlPrevOwnerID).CopyTo(yGNS, lPos) : lPos += 4

                            oHomePlanet.PlanetName.CopyTo(yGNS, lPos) : lPos += 20

                            Me.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                            yGNS(lPos) = Me.yGender : lPos += 1
                            Me.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

                            If mlPrevOwnerID > 0 Then
                                Dim oTmpPlayer As Player = GetEpicaPlayer(mlPrevOwnerID)
                                If oTmpPlayer Is Nothing = False Then
                                    oTmpPlayer.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                                    yGNS(lPos) = oTmpPlayer.yGender : lPos += 1
                                Else : lPos += 21
                                End If
                            Else : lPos += 21
                            End If

                            goMsgSys.SendToEmailSrvr(yGNS)
                        End If


                    End If
                Else
                    'ok, demotion
                    If (yPlayerTitle And Player.PlayerRank.ExRankShift) = 0 Then
                        If yPlayerTitle = yCustomTitle Then yCustomTitle = yPlayerTitle Or Player.PlayerRank.ExRankShift
                        yPlayerTitle = yPlayerTitle Or Player.PlayerRank.ExRankShift
                        dtExTitleEnd = Now.AddDays(7)

                        Dim oPC As PlayerComm = Nothing
                        If (iInternalEmailSettings And eEmailSettings.eTitleChange) <> 0 Then
                            Dim oSB As New System.Text.StringBuilder
                            oSB.AppendLine(Me.sPlayerNameProper & ",")
                            oSB.AppendLine()
                            oSB.AppendLine("  The Galactic Senate has deemed it necessary to strip you of your title of " & GetPlayerTitle(yTestTitle, Me.yGender <> 2) & " because you no longer meet the criteria. Because of this demotion, the Galactic Senate will refer to you as " & GetPlayerTitle(yPlayerTitle, Me.yGender <> 2) & " for seven days. During this time, you will be unable to be a member of a faction or have faction members.")
                            oSB.AppendLine()
                            oSB.AppendLine()
                            oSB.AppendLine("Galactic Senate Title Administration")
                            Dim sSubject As String = "Title Demotion to " & GetPlayerTitle(yLevel, Me.yGender <> 2) & " from " & GetPlayerTitle(yTestTitle, Me.yGender <> 2) & "."
                            oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, sSubject, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                        End If
                        If lConnectedPrimaryID > -1 OrElse Me.HasOnlineAliases(AliasingRights.eViewEmail) = True Then
                            If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)

                            Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                            CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
                        End If
                    End If
                End If

                Dim yMsg(6) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerTitleChange).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                yMsg(6) = yPlayerTitle
                Me.SendPlayerMessage(yMsg, False, 0)

                'Reverify our slot configurations...
                ReverifySlots()
            ElseIf (yPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 Then
                If Now > dtExTitleEnd Then
                    dtExTitleEnd = Date.MinValue
                    If yCustomTitle = yPlayerTitle Then yCustomTitle = yExTitleNew
                    yPlayerTitle = yExTitleNew
                    Dim yMsg(6) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerTitleChange).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                    yMsg(6) = yPlayerTitle
                    Me.SendPlayerMessage(yMsg, False, 0)

                    Dim oPC As PlayerComm = Nothing
                    If (iInternalEmailSettings And eEmailSettings.eTitleChange) <> 0 Then
                        Dim oSB As New System.Text.StringBuilder
                        oSB.AppendLine(Me.sPlayerNameProper & ",")
                        oSB.AppendLine()
                        oSB.AppendLine("  The seven day period for your recent title demotion has ended.")
                        oSB.AppendLine()
                        oSB.AppendLine()
                        oSB.AppendLine("Galactic Senate Title Administration")
                        Dim sSubject As String = "Important News About Your Title"
                        oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, sSubject, Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                    End If
                    If lConnectedPrimaryID > -1 OrElse Me.HasOnlineAliases(AliasingRights.eViewEmail) = True Then
                        If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                        CreateAndSendPlayerSpecialAttributeSetting(ePlayerSpecialAttributeSetting.eCelebrationPeriod, lTemp)
                    End If
                End If
            End If

        End If

    End Sub

    Public Shared Function GetPlayerTitle(ByVal pyTitle As Byte, ByVal pbIsMale As Boolean) As String
        Dim sTemp As String = ""
        Dim bExRank As Boolean = False

        If (pyTitle And PlayerRank.ExRankShift) <> 0 Then
            bExRank = True
            pyTitle = pyTitle Xor PlayerRank.ExRankShift
        End If

        Select Case pyTitle
            Case PlayerRank.Baron
                If pbIsMale = True Then sTemp = "Baron " Else sTemp = "Baroness "
            Case PlayerRank.Duke
                If pbIsMale = True Then sTemp = "Duke " Else sTemp = "Duchess "
            Case PlayerRank.Emperor
                If pbIsMale = True Then sTemp = "Emperor " Else sTemp = "Empress "
            Case PlayerRank.Governor
                If pbIsMale = True Then sTemp = "Governor " Else sTemp = "Governess "
            Case PlayerRank.King
                If pbIsMale = True Then sTemp = "King " Else sTemp = "Queen "
            Case PlayerRank.Magistrate
                sTemp = "Magistrate "
            Case PlayerRank.Overseer
                sTemp = "Overseer "
        End Select

        If bExRank = True Then sTemp = "Ex-" & sTemp

        Return sTemp '& sName
    End Function

    Private Function GetTradepostCountInSystem(ByVal lSystemID As Int32) As Int32
        Dim lCnt As Int32 = 0

        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) > -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    Dim oParent As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                    Dim bGood As Boolean = False
                    If oParent Is Nothing = False Then
                        If oParent.ObjTypeID = ObjectType.ePlanet Then
                            bGood = CType(oParent, Planet).ParentSystem.ObjectID = lSystemID
                        ElseIf oParent.ObjTypeID = ObjectType.eSolarSystem Then
                            bGood = CType(oParent, SolarSystem).ObjectID = lSystemID
                        ElseIf oParent.ObjTypeID = ObjectType.eFacility Then
                            oParent = CType(CType(oParent, Facility).ParentObject, Epica_GUID)
                            If oParent Is Nothing = False Then
                                If (oParent.ObjTypeID = ObjectType.eSolarSystem AndAlso oParent.ObjectID = lSystemID) OrElse _
                                    (oParent.ObjTypeID = ObjectType.ePlanet AndAlso CType(oParent, Planet).ParentSystem.ObjectID = lSystemID) Then
                                    bGood = True
                                End If
                            End If
                        End If
                    End If

                    If bGood = True Then
                        Try
                            For Y As Int32 = 0 To oColony.ChildrenUB
                                If oColony.lChildrenIdx(Y) > -1 Then
                                    Dim oTmpFac As Facility = oColony.oChildren(Y)
                                    If oTmpFac Is Nothing = False Then
                                        If oTmpFac.yProductionType = ProductionType.eTradePost Then
                                            lCnt += 1
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next Y
                        Catch
                        End Try
                    End If

                    If lCnt > 4 Then Exit For
                End If
            End If
        Next X

        Return lCnt
    End Function
    Private Function GetResearchLabCount(ByRef oColony As Colony, ByVal lMinProd As Int32) As Int32
        Dim lCnt As Int32 = 0
        Try
            With oColony
                For X As Int32 = 0 To .ChildrenUB
                    If .lChildrenIdx(X) > -1 Then
                        Dim oTmpFac As Facility = .oChildren(X)
                        If oTmpFac Is Nothing = False AndAlso oTmpFac.yProductionType = ProductionType.eResearch Then
                            If oTmpFac.EntityDef.ProdFactor > lMinProd Then
                                lCnt += 1
                            End If
                        End If
                    End If
                Next X
            End With
        Catch
        End Try
        Return lCnt
    End Function
    Public Sub TestCustomTitlePermissions_Facility(ByRef oFac As Facility)
        If AccountStatus <> AccountStatusType.eActiveAccount OrElse yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return

        Dim lOriginalPermission As Int32 = Me.lCustomTitlePermission

        If oFac.yProductionType = ProductionType.eTradePost Then
            'ok, first, trader
            If (Me.lCustomTitlePermission And elCustomRankPermissions.Trader) = 0 Then
                Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Trader
            End If
            'Next is merchant
            If (Me.lCustomTitlePermission And elCustomRankPermissions.Merchant) = 0 Then
                'ok, let's go check the system...
                Dim lSystemID As Int32 = -1
                If oFac.ParentObject Is Nothing = False Then
                    If oFac.ParentObject Is Nothing = False Then
                        Select Case CType(oFac.ParentObject, Epica_GUID).ObjTypeID
                            Case ObjectType.ePlanet
                                lSystemID = CType(oFac.ParentObject, Planet).ParentSystem.ObjectID
                            Case ObjectType.eSolarSystem
                                lSystemID = CType(oFac.ParentObject, SolarSystem).ObjectID
                            Case ObjectType.eFacility
                                Dim oParent As Epica_GUID = CType(CType(oFac.ParentObject, Facility).ParentObject, Epica_GUID)
                                If oParent Is Nothing = False Then
                                    If oParent.ObjTypeID = ObjectType.eSolarSystem Then
                                        lSystemID = oParent.ObjectID
                                    ElseIf oParent.ObjTypeID = ObjectType.ePlanet Then
                                        lSystemID = CType(oParent, Planet).ParentSystem.ObjectID
                                    End If
                                End If
                        End Select
                    End If
                End If

                If lSystemID <> -1 Then
                    Dim lCnt As Int32 = GetTradepostCountInSystem(lSystemID)
                    If lCnt > 4 Then
                        Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Merchant
                    End If
                End If
            End If
        ElseIf oFac.yProductionType = ProductionType.eResearch Then
            'ok, first, explorer
            If (Me.lCustomTitlePermission And elCustomRankPermissions.Explorer) = 0 Then
                Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Explorer
            End If
            'Now, we call discovermineral test for everything else
            TestCustomTitlePermissions_DiscoverMineral()
        End If
        ProcessCustomTitlePermissionChange(lOriginalPermission)
    End Sub
    Public Sub TestCustomTitlePermissions_DiscoverMineral()
        If Me.AccountStatus <> AccountStatusType.eActiveAccount OrElse yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return

        Dim lOriginalPermission As Int32 = Me.lCustomTitlePermission

        'test scientist
        If (Me.lCustomTitlePermission And elCustomRankPermissions.Scientist) = 0 Then
            'need 10 discovered and 2 reslabs in a colony
            Dim lDiscoveredMins As Int32 = Me.ResearchedMineralCount()
            If lDiscoveredMins > 9 Then
                For X As Int32 = 0 To mlColonyUB
                    If mlColonyIdx(X) > -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                        Dim oColony As Colony = goColony(mlColonyIdx(X))
                        If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = Me.ObjectID Then
                            Dim lResLabCnt As Int32 = GetResearchLabCount(oColony, 0)
                            If lResLabCnt > 1 Then
                                Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Scientist
                                Exit For
                            End If
                        End If
                    End If
                Next X
            End If
        End If
        If (Me.lCustomTitlePermission And elCustomRankPermissions.MasterScientist) = 0 Then
            'need 30 discovered and 1 alloy and 2 labs > 400 prod in a colony
            Dim lDiscoveredMins As Int32 = Me.ResearchedMineralCount
            If lDiscoveredMins > 29 Then
                'does the player have 1 alloy researched?
                Dim bAlloyGood As Boolean = False
                For X As Int32 = 0 To mlAlloyUB
                    If mlAlloyIdx(X) > -1 Then
                        Dim oAlloy As AlloyTech = moAlloy(X)
                        If oAlloy Is Nothing = False AndAlso oAlloy.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                            bAlloyGood = True
                            Exit For
                        End If
                    End If
                Next X

                If bAlloyGood = True Then
                    For X As Int32 = 0 To mlColonyUB
                        If mlColonyIdx(X) > -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                            Dim oColony As Colony = goColony(mlColonyIdx(X))
                            If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = Me.ObjectID Then
                                Dim lResLabCnt As Int32 = GetResearchLabCount(oColony, 400)
                                If lResLabCnt > 1 Then
                                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.MasterScientist
                                    Exit For
                                End If
                            End If
                        End If
                    Next X
                End If
            End If
        End If

        ProcessCustomTitlePermissionChange(lOriginalPermission)
    End Sub
    Public Sub TestCustomTitlePermissions_Research(ByVal iResTypeID As Int16)
        If Me.AccountStatus <> AccountStatusType.eActiveAccount OrElse yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return

        If (Me.lCustomTitlePermission And elCustomRankPermissions.MasterScientist) = 0 Then
            If iResTypeID = ObjectType.eAlloyTech Then TestCustomTitlePermissions_DiscoverMineral()
        End If

        If (Me.lCustomTitlePermission And (elCustomRankPermissions.Inquisitor Or elCustomRankPermissions.Preeminence Or elCustomRankPermissions.Transcendent)) <> (elCustomRankPermissions.Inquisitor Or elCustomRankPermissions.Preeminence Or elCustomRankPermissions.Transcendent) Then

            Dim lOriginal As Int32 = Me.lCustomTitlePermission

            Dim lMinCntToStayIn As Int32 = 1
            If (Me.lCustomTitlePermission And elCustomRankPermissions.Inquisitor) <> 0 Then
                lMinCntToStayIn = 2
                If (Me.lCustomTitlePermission And elCustomRankPermissions.Preeminence) <> 0 Then
                    lMinCntToStayIn = 5
                End If
            End If

            Dim lAlloy As Int32 = 0
            Dim lArmor As Int32 = 0
            Dim lEngine As Int32 = 0
            Dim lHull As Int32 = 0
            Dim lPrototype As Int32 = 0
            Dim lRadar As Int32 = 0
            Dim lShield As Int32 = 0
            Dim lWeapon As Int32 = 0
            Dim lSpecials As Int32 = 0
            Dim lNearPerfect As Int32 = 0

            'put in order of occurrence in player designs during beta (lowest to highest)

            For X As Int32 = 0 To mlShieldUB
                If mlShieldIdx(X) > -1 Then
                    Dim oShield As ShieldTech = moShield(X)
                    If oShield Is Nothing = False AndAlso oShield.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        lShield += 1
                        If (oShield.MajorDesignFlaw Or Epica_Tech.eComponentDesignFlaw.eGoodDesign) <> 0 Then lNearPerfect += 1
                    End If
                End If
            Next X
            If lShield < lMinCntToStayIn Then Return

            For X As Int32 = 0 To mlRadarUB
                If mlRadarIdx(X) > -1 Then
                    Dim oRadar As RadarTech = moRadar(X)
                    If oRadar Is Nothing = False AndAlso oRadar.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        If oRadar.OptimumRange > 30 OrElse oRadar.MaximumRange > 1 OrElse oRadar.JamEffect > 0 Then
                            lRadar += 1
                            If (oRadar.MajorDesignFlaw Or Epica_Tech.eComponentDesignFlaw.eGoodDesign) <> 0 Then lNearPerfect += 1
                        End If
                    End If
                End If
            Next X
            If lRadar < lMinCntToStayIn Then Return

            For X As Int32 = 0 To mlWeaponUB
                If mlWeaponIdx(X) > -1 Then
                    Dim oWeapon As BaseWeaponTech = moWeapon(X)
                    If oWeapon Is Nothing = False AndAlso oWeapon.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        lWeapon += 1
                        If (oWeapon.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eGoodDesign) <> 0 Then lNearPerfect += 1
                    End If
                End If
            Next X
            If lWeapon < lMinCntToStayIn Then Return

            For X As Int32 = 0 To mlArmorUB
                If mlArmorIdx(X) > -1 Then
                    Dim oArmor As ArmorTech = moArmor(X)
                    If oArmor Is Nothing = False AndAlso oArmor.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        If oArmor.lHPPerPlate > 10 AndAlso oArmor.lHPPerPlate / Math.Max(1, oArmor.lHullUsagePerPlate) > 10 Then
                            lArmor += 1
                            If (oArmor.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eGoodDesign) <> 0 Then lNearPerfect += 1
                        End If
                    End If
                End If
            Next X
            If lArmor < lMinCntToStayIn Then Return

            For X As Int32 = 0 To mlEngineUB
                If mlEngineIdx(X) > -1 Then
                    Dim oEngine As EngineTech = moEngine(X)
                    If oEngine Is Nothing = False AndAlso oEngine.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        If oEngine.Thrust = 0 Then
                            If oEngine.PowerProd > 2000 Then
                                lEngine += 1
                                If (oEngine.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eGoodDesign) <> 0 Then lNearPerfect += 1
                            End If
                        Else
                            Dim blTemp As Int64 = CLng(oEngine.Thrust) * CLng(oEngine.Maneuver) * CLng(oEngine.Speed) * CLng(oEngine.PowerProd)
                            If blTemp > 1025000 Then
                                lEngine += 1
                                If (oEngine.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eGoodDesign) <> 0 Then lNearPerfect += 1
                            End If
                        End If
                    End If
                End If
            Next X
            If lEngine < lMinCntToStayIn Then Return

            For X As Int32 = 0 To mlAlloyUB
                If mlAlloyIdx(X) > -1 Then
                    Dim oAlloy As AlloyTech = moAlloy(X)
                    If oAlloy Is Nothing = False AndAlso oAlloy.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        lAlloy += 1
                        If lAlloy > 4 Then Exit For
                    End If
                End If
            Next X
            If lAlloy < lMinCntToStayIn Then Return

            For X As Int32 = 0 To mlHullUB
                If mlHullIdx(X) > -1 Then
                    Dim oHull As HullTech = moHull(X)
                    If oHull Is Nothing = False AndAlso oHull.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        lHull += 1
                        If lHull > 4 Then Exit For
                    End If
                End If
            Next X
            If lHull < lMinCntToStayIn Then Return

            For X As Int32 = 0 To mlPrototypeUB
                If mlPrototypeIdx(X) > -1 Then
                    Dim oPrototype As Prototype = moPrototype(X)
                    If oPrototype Is Nothing = False AndAlso oPrototype.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        lPrototype += 1
                        If lPrototype > 4 Then Exit For
                    End If
                End If
            Next X
            If lPrototype < lMinCntToStayIn Then Return

            For X As Int32 = 0 To oSpecials.mlSpecialTechUB
                If oSpecials.mlSpecialTechIdx(X) > -1 Then
                    Dim oSpecial As PlayerSpecialTech = oSpecials.moSpecialTech(X)
                    If oSpecial Is Nothing = False AndAlso oSpecial.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                        lSpecials += 1
                    End If
                End If
            Next X

            'Now, to test our results
            Dim lMinComponentCnt As Int32 = Math.Min(lHull, Math.Min(Math.Min(lAlloy, lArmor), lEngine))
            lMinComponentCnt = Math.Min(Math.Min(Math.Min(Math.Min(lMinComponentCnt, lPrototype), lRadar), lShield), lWeapon)

            If (Me.lCustomTitlePermission And elCustomRankPermissions.Inquisitor) = 0 Then
                If lMinComponentCnt > 0 AndAlso lSpecials > 0 Then
                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Inquisitor
                End If
            End If
            If (Me.lCustomTitlePermission And elCustomRankPermissions.Preeminence) = 0 Then
                If lMinComponentCnt > 1 AndAlso lSpecials > 19 AndAlso lNearPerfect > 1 Then
                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Preeminence
                End If
            End If
            If (Me.lCustomTitlePermission And elCustomRankPermissions.Transcendent) = 0 Then
                If lMinComponentCnt > 4 AndAlso lSpecials > 49 AndAlso lNearPerfect > 9 Then
                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Transcendent
                End If
            End If

            ProcessCustomTitlePermissionChange(lOriginal)
        End If
    End Sub
    Public Sub TestCustomTitlePermissions_Allies()
        If Me.AccountStatus <> AccountStatusType.eActiveAccount OrElse yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return
        If (Me.lCustomTitlePermission And (elCustomRankPermissions.Diplomat Or elCustomRankPermissions.Arbiter Or elCustomRankPermissions.Chancellor Or elCustomRankPermissions.SupremeChancellor)) = (elCustomRankPermissions.Diplomat Or elCustomRankPermissions.Arbiter Or elCustomRankPermissions.Chancellor Or elCustomRankPermissions.SupremeChancellor) Then Return

        Dim oTmpGuild As Guild = oGuild
        Dim lAllies As Int32 = 0
        Dim lBloodAllies As Int32 = 0

        Dim lSystemID(10) As Int32
        Dim lOwnedPlanetCnt(10) As Int32
        Dim lSystemPCnt(10) As Int32
        Dim lSysUB As Int32 = -1

        For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) > -1 Then
                Dim oRel As PlayerRel = moPlayerRels(X)
                If oRel.oPlayerRegards Is Nothing = False AndAlso oRel.oThisPlayer Is Nothing = False Then
                    If oRel.oPlayerRegards.ObjectID = Me.ObjectID AndAlso oRel.oThisPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso oRel.WithThisScore > elRelTypes.ePeace Then
                        If oTmpGuild Is Nothing = False Then
                            Dim oMember As GuildMember = oTmpGuild.GetMember(oRel.oThisPlayer.ObjectID)
                            If oMember Is Nothing = False AndAlso (oMember.yMemberState And GuildMemberState.Approved) <> 0 Then
                                Continue For
                            End If
                        End If

                        'ok, does the player share the same rel?
                        Dim yTheirRelWithMe As Byte = oRel.oThisPlayer.GetPlayerRelScore(Me.ObjectID)
                        If yTheirRelWithMe > elRelTypes.ePeace Then
                            'ok, increment our allies
                            lAllies += 1
                            'now, is our relationship with them blood ally?
                            If oRel.WithThisScore > elRelTypes.eAlly AndAlso yTheirRelWithMe > elRelTypes.eAlly Then
                                'yes... increment blood allies
                                lBloodAllies += 1
                                'ok, get their system and ownernership lists
                                Dim oPlayer As Player = oRel.oThisPlayer
                                For Y As Int32 = 0 To oPlayer.mlColonyUB
                                    'loop thru the colony list of the player
                                    If oPlayer.mlColonyIdx(Y) > -1 AndAlso glColonyIdx(oPlayer.mlColonyIdx(Y)) = oPlayer.mlColonyID(Y) Then
                                        Dim oColony As Colony = goColony(oPlayer.mlColonyIdx(Y))
                                        If oColony Is Nothing = False Then
                                            Dim oParent As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                                            'parent a valid planet?
                                            If oParent Is Nothing = False AndAlso oParent.ObjTypeID = ObjectType.ePlanet Then
                                                Dim oPlanet As Planet = CType(oColony.ParentObject, Planet)
                                                'does the player 'own' the planet?
                                                If oPlanet.OwnerID = oPlayer.ObjectID AndAlso oPlanet.ParentSystem Is Nothing = False Then
                                                    Dim lPsys As Int32 = oPlanet.ParentSystem.ObjectID
                                                    Dim bFound As Boolean = False
                                                    For Z As Int32 = 0 To lSysUB
                                                        If lSystemID(Z) = lPsys Then
                                                            bFound = True
                                                            lOwnedPlanetCnt(Z) += 1
                                                            Exit For
                                                        End If
                                                    Next Z
                                                    If bFound = False Then
                                                        lSysUB += 1
                                                        If lSystemID.GetUpperBound(0) < lSysUB Then
                                                            ReDim Preserve lSystemID(lSysUB + 10)
                                                            ReDim Preserve lOwnedPlanetCnt(lSysUB + 10)
                                                            ReDim Preserve lSystemPCnt(lSysUB + 10)
                                                        End If
                                                        lSystemID(lSysUB) = lPsys
                                                        lOwnedPlanetCnt(lSysUB) = 1
                                                        lSystemPCnt(lSysUB) = oPlanet.ParentSystem.mlPlanetUB + 1
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Y

                            End If
                        End If

                    End If
                End If
            End If
        Next X

        'add in my ownership
        For X As Int32 = 0 To Me.mlColonyUB
            If mlColonyIdx(X) > -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    Dim oParent As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                    If oParent Is Nothing = False AndAlso oParent.ObjTypeID = ObjectType.ePlanet Then
                        Dim oPlanet As Planet = CType(oColony.ParentObject, Planet)
                        If oPlanet.OwnerID = Me.ObjectID AndAlso oPlanet.ParentSystem Is Nothing = False Then
                            Dim lPSys As Int32 = oPlanet.ParentSystem.ObjectID
                            For Y As Int32 = 0 To lSysUB
                                If lSystemID(Y) = lPSys Then
                                    lOwnedPlanetCnt(Y) += 1
                                    Exit For
                                End If
                            Next Y
                        End If
                    End If
                End If
            End If
        Next X

        'Now, we have everything we need to continue
        Dim lOriginal As Int32 = Me.lCustomTitlePermission
        If (Me.lCustomTitlePermission And elCustomRankPermissions.Diplomat) = 0 Then
            If lAllies > 0 Then Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Diplomat
        End If
        If (Me.lCustomTitlePermission And elCustomRankPermissions.Arbiter) = 0 Then
            If lAllies > 4 AndAlso lBloodAllies > 0 Then Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Arbiter
        End If
        If (Me.lCustomTitlePermission And (elCustomRankPermissions.Chancellor Or elCustomRankPermissions.SupremeChancellor)) <> (elCustomRankPermissions.Chancellor Or elCustomRankPermissions.SupremeChancellor) Then
            Dim lSystems As Int32 = 0
            For X As Int32 = 0 To lSysUB
                If lSystemID(X) > -1 Then
                    If lSystemPCnt(X) = lOwnedPlanetCnt(X) Then lSystems += 1
                End If
            Next X
            If (Me.lCustomTitlePermission And elCustomRankPermissions.Chancellor) = 0 Then
                If lSystems > 0 Then
                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Chancellor
                End If
            End If
            If (Me.lCustomTitlePermission And elCustomRankPermissions.SupremeChancellor) = 0 Then
                If lSystems > 4 Then
                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.SupremeChancellor
                End If
            End If
        End If

        ProcessCustomTitlePermissionChange(lOriginal)

    End Sub
    Public Sub TestCustomTitlePermissions_Trade()
        If Me.AccountStatus <> AccountStatusType.eActiveAccount OrElse yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return
        'broker, chief broker, master merchant, commerce czar
        If (Me.lCustomTitlePermission And (elCustomRankPermissions.Broker Or elCustomRankPermissions.ChiefBroker Or elCustomRankPermissions.MasterMerchant Or elCustomRankPermissions.CommerceCzar Or elCustomRankPermissions.TradeLord)) <> (elCustomRankPermissions.Broker Or elCustomRankPermissions.ChiefBroker Or elCustomRankPermissions.MasterMerchant Or elCustomRankPermissions.CommerceCzar Or elCustomRankPermissions.TradeLord) Then
            Dim lOriginal As Int32 = Me.lCustomTitlePermission

            Dim lSystemID(10) As Int32
            Dim lPCnt(10) As Int32
            Dim lOwnedCnt(10) As Int32
            Dim lSysUB As Int32 = -1

            For X As Int32 = 0 To mlColonyUB
                If mlColonyIdx(X) > -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                    Dim oColony As Colony = goColony(mlColonyIdx(X))
                    If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = Me.ObjectID Then
                        Dim oParent As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                        If oParent Is Nothing OrElse oParent.ObjTypeID <> ObjectType.ePlanet Then Continue For

                        For Y As Int32 = 0 To oColony.ChildrenUB
                            If oColony.lChildrenIdx(Y) > -1 Then
                                Dim oFac As Facility = oColony.oChildren(Y)
                                If oFac Is Nothing = False AndAlso oFac.yProductionType = ProductionType.eTradePost Then
                                    If oFac.lTradePostSellSlotsUsed > 4 Then
                                        Dim lSysID As Int32 = CType(oParent, Planet).ParentSystem.ObjectID
                                        Dim bFound As Boolean = False
                                        For Z As Int32 = 0 To lSysUB
                                            If lSystemID(Z) = lSysID Then
                                                lOwnedCnt(Z) += 1
                                                bFound = True
                                                Exit For
                                            End If
                                        Next Z
                                        If bFound = False Then
                                            lSysUB += 1
                                            If lSysUB > lSystemID.GetUpperBound(0) Then
                                                ReDim Preserve lSystemID(lSysUB + 10)
                                                ReDim Preserve lPCnt(lSysUB + 10)
                                                ReDim Preserve lOwnedCnt(lSysUB + 10)
                                            End If
                                            lSystemID(lSysUB) = lSysID
                                            lPCnt(lSysUB) = CType(oParent, Planet).ParentSystem.mlPlanetUB + 1
                                            lOwnedCnt(lSysUB) = 1
                                        End If
                                    End If
                                    If oFac.lTradePostSellSlotsUsed > 9 AndAlso (Me.lCustomTitlePermission And elCustomRankPermissions.Broker) = 0 Then
                                        'need 5 different types
                                        Dim lTypes As Int32 = goGTC.NumberOfMarketTypesForSell(oFac.ObjectID)
                                        If lTypes > 4 Then
                                            Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Broker
                                        End If
                                    End If
                                    If oFac.lTradePostSellSlotsUsed > 9 AndAlso (Me.lCustomTitlePermission And elCustomRankPermissions.TradeLord) = 0 Then
                                        Dim lTypes As Int32 = goGTC.NumberOfMarketTypesForSell(oFac.ObjectID)
                                        If lTypes > 9 Then
                                            Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.TradeLord
                                        End If
                                    End If
                                    If oFac.lTradePostSellSlotsUsed > 19 AndAlso (Me.lCustomTitlePermission And elCustomRankPermissions.ChiefBroker) = 0 Then
                                        'need 10 different types
                                        Dim lTypes As Int32 = goGTC.NumberOfMarketTypesForSell(oFac.ObjectID)
                                        If lTypes > 9 Then
                                            Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.ChiefBroker
                                        End If
                                    End If

                                    Exit For
                                End If
                            End If
                        Next Y
                    End If
                End If
            Next X

            'Now, check our other systems
            If (Me.lCustomTitlePermission And (elCustomRankPermissions.MasterMerchant Or elCustomRankPermissions.CommerceCzar)) <> (elCustomRankPermissions.MasterMerchant Or elCustomRankPermissions.CommerceCzar) Then
                Dim lSystems As Int32 = 0
                For X As Int32 = 0 To lSysUB
                    If lSystemID(X) > -1 AndAlso lOwnedCnt(X) = lPCnt(X) Then
                        lSystems += 1
                    End If
                Next X
                If lSystems > 0 Then
                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.MasterMerchant
                End If
                If lSystems > 2 Then
                    Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.CommerceCzar
                End If
            End If

            ProcessCustomTitlePermissionChange(lOriginal)
        End If
    End Sub
    Public Sub ProcessCustomTitlePermissionChange(ByVal lOriginal As Int32)
        If lOriginal <> Me.lCustomTitlePermission Then
            Dim lNewVals As Int32 = lOriginal Xor Me.lCustomTitlePermission
            If (lNewVals And elCustomRankPermissions.Arbiter) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                    "Due to your recent achievement of having 5 allies and 1 blood ally that are non-guild mates, you have been rewarded the title of Arbiter. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Broker) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having 10 concurrent sell orders of 5 different market types, you have been rewarded the title of Broker. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Chancellor) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having enough blood allies that, combined, you control and entire star system, you have been rewarded the title of Chancellor. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.ChiefBroker) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having 20 concurrent sell orders of 10 different market types, you have been rewarded the title of Chief Broker. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.ChiefScientist) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having multiple labs on multiple planets with 400 research production all assigned to the same project, you have been rewarded the title of Chief Scientist. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.CommerceCzar) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having a tradepost on every planet of 3 systems with at least 5 sell orders in each, you have been rewarded the title of Commerce Czar. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Counselor) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having 5 members in your faction, you have been rewarded the title of Counselor. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Diplomat) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having your first ally relationship, you have been rewarded the title of Diplomat. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Explorer) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of building your first research lab outside of the tutorial, you have been rewarded the title of Explorer. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.HighSenator) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of proposing a senate item that was passed in which all of your allies voted for the proposal, you have been rewarded the title of High Senator. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Inquisitor) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having researched 1 of each type of component, you have been rewarded the title of Inquisitor. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.MasterMerchant) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having a tradepost on every planet of a star system with at least 5 sell orders in each, you have been rewarded the title of Master Merchant. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.MasterScientist) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having 2 research labs with production of at least 400, discovering 30 minerals and researching 1 alloy, you have been rewarded the title of Master Scientist. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Merchant) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having 5 tradeposts in the same start system, you have been rewarded the title of Merchant. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Preeminence) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of researching 2 of each type of component, having 20 special projects researched and at least 2 near perfect designs, you have been rewarded the title of Preeminence. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Scientist) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of building 2 research labs and discovering at least 10 minerals, you have been rewarded the title of Scientist. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Senator) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having a vote that went the way your system went, you have been rewarded the title of Senator. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.SupremeChancellor) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of having blood allies that have a combined power to control at least 5 star systems, you have been rewarded the title of Supreme Chancellor. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.TradeLord) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of completing a direct trade that resulted in a large profit, you have been rewarded the title of Trade Lord. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Trader) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of building your first tradepost outside of the tutorial, you have been rewarded the title of Trader. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            If (lNewVals And elCustomRankPermissions.Transcendent) <> 0 Then
                Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                 "Due to your recent achievement of researching 5 of each type of component, having 50 special projects researched and at least 10 near perfect designs, you have been rewarded the title of Transcendent. You can change your title in the Diplomacy window.", "Achievement Reward", Me.ObjectID, GetDateAsNumber(Now), False, Me.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If

            Dim yMsg(10) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerCustomTitle).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(Me.lCustomTitlePermission).CopyTo(yMsg, 6)
            yMsg(10) = yCustomTitle
            Me.SendPlayerMessage(yMsg, False, 0)
        End If
    End Sub
#End Region

#Region "  Fast Agent Lookup Management  "
    Public mlAgentIdx() As Int32
    Public mlAgentID() As Int32
    Public mlAgentUB As Int32 = -1

    Public uImposedEffects() As ImposedAgentEffect
    Public lImposedEffectUB As Int32 = -1

    Public Sub AddAgentLookup(ByVal lAgentID As Int32, ByVal lAgentIdx As Int32)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlAgentUB
            If mlAgentIdx(X) = lAgentIdx Then
                mlAgentID(X) = lAgentID
                Return
            ElseIf lIdx = -1 AndAlso mlAgentIdx(X) = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            lIdx = mlAgentUB + 1
            ReDim Preserve mlAgentIdx(lIdx)
            ReDim Preserve mlAgentID(lIdx)
            mlAgentUB += 1
        End If
        mlAgentIdx(lIdx) = lAgentIdx
        mlAgentID(lIdx) = lAgentID
    End Sub

    Public Sub RemoveAgentLookup(ByVal lAgentID As Int32)
        For X As Int32 = 0 To mlAgentUB
            If mlAgentID(X) = lAgentID Then
                mlAgentIdx(X) = -1
                mlAgentID(X) = -1
            End If
        Next X
    End Sub

    Public Function GetRandomAgentID(ByVal bNewRecruit As Boolean) As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlAgentUB
            If mlAgentIdx(X) <> -1 AndAlso glAgentIdx(mlAgentIdx(X)) = mlAgentID(X) Then
                Dim oAgent As Agent = goAgent(mlAgentIdx(X))
                If oAgent Is Nothing = False Then
                    If bNewRecruit = True AndAlso (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
                        lCnt += 1
                    ElseIf bNewRecruit = False AndAlso (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.NewRecruit)) = 0 Then
                        lCnt += 1
                    End If
                End If
            End If
        Next X

        Dim lIdxs(lCnt - 1) As Int32
        Dim lTempIdx As Int32 = -1
        For X As Int32 = 0 To mlAgentUB
            If mlAgentIdx(X) <> -1 AndAlso glAgentIdx(mlAgentIdx(X)) = mlAgentID(X) Then
                Dim oAgent As Agent = goAgent(mlAgentIdx(X))
                If oAgent Is Nothing = False Then
                    If bNewRecruit = True AndAlso (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    ElseIf bNewRecruit = False AndAlso (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.NewRecruit)) = 0 Then
                        lTempIdx += 1
                        lIdxs(lTempIdx) = X
                    End If
                End If
            End If
        Next X
        If lCnt = 0 Then Return -1
        Dim lVal As Int32 = CInt(Math.Floor(Rnd() * lCnt))
        If lVal > -1 AndAlso lVal <= lIdxs.GetUpperBound(0) Then lVal = lIdxs(lVal)

        Try
            Return mlAgentID(lVal)
        Catch
            Return -1
        End Try
    End Function

    Public Sub AdjustAllAgentLoyalty(ByVal lValue As Int32)
        For X As Int32 = 0 To mlAgentUB
            If mlAgentIdx(X) > -1 AndAlso glAgentIdx(mlAgentIdx(X)) = mlAgentID(X) Then
                Dim oAgent As Agent = goAgent(mlAgentIdx(X))
                If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And AgentStatus.NewRecruit) = 0 Then
                    Dim lTemp As Int32 = oAgent.Loyalty
                    lTemp += lValue
                    If lTemp < 1 Then lTemp = 1
                    If lTemp > 100 Then lTemp = 100
                    oAgent.Loyalty = CByte(lTemp)
                End If
            End If
        Next X
    End Sub

    Public Sub AddImposedEffect(ByRef oEffect As AgentEffect, ByRef oTargetPlayer As Player, ByVal sSpecificName As String)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lImposedEffectUB
            If uImposedEffects(X).lStartCycle + uImposedEffects(X).lDuration > glCurrentCycle Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            ReDim uImposedEffects(lImposedEffectUB + 1)
            lImposedEffectUB += 1
            lIdx = lImposedEffectUB
        End If
        With uImposedEffects(lIdx)
            .bAmountAsPerc = oEffect.bAmountAsPerc
            .lAmount = oEffect.lAmount
            .lDuration = oEffect.lDuration
            .lStartCycle = oEffect.lStartCycle
            .oTargetPlayer = oTargetPlayer
            .sSpecificName = sSpecificName
            .yType = oEffect.yType
        End With
    End Sub

    Public Function GetImposedAgentListMsg() As Byte()
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lImposedEffectUB
            If uImposedEffects(X).lStartCycle + uImposedEffects(X).lDuration < glCurrentCycle Then
                lCnt += 1
            End If
        Next X

        Dim yMsg(5 + (lCnt * 34)) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetImposedAgentEffects).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

        For X As Int32 = 0 To lImposedEffectUB
            If uImposedEffects(X).lStartCycle + uImposedEffects(X).lDuration < glCurrentCycle Then
                lCnt -= 1
                If lCnt < 0 Then Exit For

                With uImposedEffects(X)
                    System.BitConverter.GetBytes(.oTargetPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    Dim lCyclesRemaining As Int32 = .lDuration - (glCurrentCycle - .lStartCycle)
                    System.BitConverter.GetBytes(lCyclesRemaining).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lAmount).CopyTo(yMsg, lPos) : lPos += 4
                    yMsg(lPos) = .yType : lPos += 1
                    If .bAmountAsPerc = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
                    lPos += 1
                    If .sSpecificName.Length > 20 Then .sSpecificName = Mid(.sSpecificName, 1, 20)
                    StringToBytes(.sSpecificName).CopyTo(yMsg, lPos) : lPos += 20
                End With
            End If
        Next X

        Return yMsg
    End Function
#End Region

#Region "  Player Intel  "
    Public moPlayerIntel() As PlayerIntel
    Public mlPlayerIntelIdx() As Int32      'the 'other' player id
    Public mlPlayerIntelUB As Int32 = -1

    Public moPlayerTechKnowledge() As PlayerTechKnowledge
    Public myPlayerTechKnowledgeUsed() As Byte
    Public mlPlayerTechKnowledgeUB As Int32 = -1

    Public Sub SetPlayerIntel(ByRef oPlayerIntel As PlayerIntel)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlPlayerIntelUB
            If mlPlayerIntelIdx(X) = oPlayerIntel.ObjectID Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            mlPlayerIntelUB += 1
            ReDim Preserve moPlayerIntel(mlPlayerIntelUB)
            ReDim Preserve mlPlayerIntelIdx(mlPlayerIntelUB)
            lIdx = mlPlayerIntelUB
        End If

        moPlayerIntel(lIdx) = oPlayerIntel
        mlPlayerIntelIdx(lIdx) = oPlayerIntel.ObjectID
    End Sub

    'Public Sub CheckForAndAddPlayerIntel(ByVal lPlayerID As Int32)
    '	For X As Int32 = 0 To mlPlayerIntelUB
    '		If mlPlayerIntelIdx(X) = lPlayerID Then
    '			Return
    '		End If
    '	Next X

    '	ReDim Preserve moPlayerIntel(mlPlayerIntelUB + 1)
    '	ReDim Preserve mlPlayerIntelIdx(mlPlayerIntelUB + 1)
    '	mlPlayerIntelIdx(mlPlayerIntelUB + 1) = lPlayerID
    '	moPlayerIntel(mlPlayerIntelUB + 1) = New PlayerIntel()
    '	mlPlayerIntelUB += 1

    '	With moPlayerIntel(mlPlayerIntelUB)
    '		.lPlayerID = Me.ObjectID
    '		.ObjectID = lPlayerID
    '		.ObjTypeID = ObjectType.ePlayerIntel
    '		.StaticVariables = 0
    '	End With

    '	If Me.oSocket Is Nothing = False OrElse Me.HasOnlineAliases(AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents) = True Then Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moPlayerIntel(mlPlayerIntelUB), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)
    'End Sub

    Public Sub AddPlayerTechKnowledge(ByRef oPTK As PlayerTechKnowledge, ByVal bNoSave As Boolean)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlPlayerTechKnowledgeUB
            If myPlayerTechKnowledgeUsed(X) <> 0 AndAlso moPlayerTechKnowledge(X).oTech.ObjectID = oPTK.oTech.ObjectID AndAlso moPlayerTechKnowledge(X).oTech.ObjTypeID = oPTK.oTech.ObjTypeID Then
                moPlayerTechKnowledge(X) = oPTK
                Return
            ElseIf lIdx = -1 AndAlso myPlayerTechKnowledgeUsed(X) = 0 Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            lIdx = mlPlayerTechKnowledgeUB + 1
            ReDim Preserve moPlayerTechKnowledge(lIdx)
            ReDim Preserve myPlayerTechKnowledgeUsed(lIdx)
            mlPlayerTechKnowledgeUB = lIdx
        End If
        moPlayerTechKnowledge(lIdx) = oPTK
        myPlayerTechKnowledgeUsed(lIdx) = 255
        If bNoSave = False Then moPlayerTechKnowledge(lIdx).SaveObject()
    End Sub

    Public Function HasTechKnowledge(ByVal lTechID As Int32, ByVal iTechTypeID As Int16, ByVal yKnowLvl As PlayerTechKnowledge.KnowledgeType) As Boolean
        For X As Int32 = 0 To mlPlayerTechKnowledgeUB
            If myPlayerTechKnowledgeUsed(X) <> 0 AndAlso moPlayerTechKnowledge(X).oTech.ObjectID = lTechID AndAlso moPlayerTechKnowledge(X).oTech.ObjTypeID = iTechTypeID Then
                If CByte(moPlayerTechKnowledge(X).yKnowledgeType) >= CByte(yKnowLvl) Then
                    Return True
                Else : Return False
                End If
            End If
        Next X
        Return False
    End Function

    Public Function GetPlayerTechKnowledgeTech(ByVal lTechID As Int32, ByVal iTechTypeID As Int16, ByVal yMinKnowLvl As PlayerTechKnowledge.KnowledgeType) As Epica_Tech
        For X As Int32 = 0 To mlPlayerTechKnowledgeUB
            If myPlayerTechKnowledgeUsed(X) <> 0 AndAlso moPlayerTechKnowledge(X).oTech.ObjectID = lTechID AndAlso moPlayerTechKnowledge(X).oTech.ObjTypeID = iTechTypeID Then
                If CByte(moPlayerTechKnowledge(X).yKnowledgeType) >= CByte(yMinKnowLvl) Then
                    Return moPlayerTechKnowledge(X).oTech
                Else : Return Nothing
                End If
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetPlayerTechKnowledge(ByVal lTechID As Int32, ByVal iTechTypeID As Int16) As PlayerTechKnowledge
        For X As Int32 = 0 To mlPlayerTechKnowledgeUB
            If myPlayerTechKnowledgeUsed(X) <> 0 AndAlso moPlayerTechKnowledge(X).oTech.ObjectID = lTechID AndAlso moPlayerTechKnowledge(X).oTech.ObjTypeID = iTechTypeID Then
                Return moPlayerTechKnowledge(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetOrAddPlayerIntel(ByVal lPlayerID As Int32, ByVal bNoAdd As Boolean) As PlayerIntel
        For X As Int32 = 0 To mlPlayerIntelUB
            If mlPlayerIntelIdx(X) = lPlayerID Then Return moPlayerIntel(X)
        Next X

        If bNoAdd = True Then Return Nothing

        ReDim Preserve moPlayerIntel(mlPlayerIntelUB + 1)
        ReDim Preserve mlPlayerIntelIdx(mlPlayerIntelUB + 1)
        mlPlayerIntelIdx(mlPlayerIntelUB + 1) = lPlayerID
        moPlayerIntel(mlPlayerIntelUB + 1) = New PlayerIntel()
        mlPlayerIntelUB += 1

        With moPlayerIntel(mlPlayerIntelUB)
            .lPlayerID = Me.ObjectID
            .ObjectID = lPlayerID
            .ObjTypeID = ObjectType.ePlayerIntel
            .StaticVariables = 0
            .SaveObject()
        End With
        Return moPlayerIntel(mlPlayerIntelUB)
    End Function
#End Region

#Region "  Player Item Intel Management  "
    Public moItemIntel() As PlayerItemIntel
    Public mlItemIntelUB As Int32 = -1

    Public Function AddPlayerItemIntel(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal yType As PlayerItemIntel.PlayerItemIntelType, ByVal lOtherPlayerID As Int32) As PlayerItemIntel
        Dim oII As PlayerItemIntel = Nothing
        For X As Int32 = 0 To mlItemIntelUB
            If moItemIntel(X).lItemID = lObjectID AndAlso moItemIntel(X).iItemTypeID = iObjTypeID AndAlso moItemIntel(X).lOtherPlayerID = lOtherPlayerID Then
                oII = moItemIntel(X)
                Exit For
            End If
        Next X

        If oII Is Nothing Then
            ReDim Preserve moItemIntel(mlItemIntelUB + 1)
            moItemIntel(mlItemIntelUB + 1) = New PlayerItemIntel
            mlItemIntelUB += 1
            oII = moItemIntel(mlItemIntelUB)
        End If

        With oII
            .dtTimeStamp = Now
            .iItemTypeID = iObjTypeID
            .lItemID = lObjectID
            .lOtherPlayerID = lOtherPlayerID
            .lPlayerID = Me.ObjectID
            .lValue = 0
            .yIntelType = yType
            .PopulateData()
            If .SaveObject() = False Then Return Nothing
        End With
        Return oII
    End Function

    Public Function HasItemIntel(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal lOtherPlayerID As Int32) As Boolean
        For X As Int32 = 0 To mlItemIntelUB
            If moItemIntel(X).lItemID = lObjectID AndAlso moItemIntel(X).iItemTypeID = iObjTypeID AndAlso moItemIntel(X).lOtherPlayerID = lOtherPlayerID Then
                Return True
            End If
        Next X
        Return False
    End Function

    Public Function GetPlayerItemIntel(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal lOtherPlayerID As Int32) As PlayerItemIntel
        For X As Int32 = 0 To mlItemIntelUB
            If moItemIntel(X).lItemID = lObjectID AndAlso moItemIntel(X).iItemTypeID = iObjTypeID AndAlso moItemIntel(X).lOtherPlayerID = lOtherPlayerID Then
                Return moItemIntel(X)
            End If
        Next X
        Return Nothing
    End Function
#End Region

#Region "  Trade History Management  "
    Private moTradeHistory() As TradeHistory
    Private mlTradeHistoryUB As Int32 = -1
    Public Sub AddTradeHistory(ByRef oTH As TradeHistory)
        mlTradeHistoryUB += 1
        ReDim Preserve moTradeHistory(mlTradeHistoryUB)
        moTradeHistory(mlTradeHistoryUB) = oTH
    End Sub
    Public Function GetTradeHistoryMsg() As Byte()
        Dim lUB As Int32 = mlTradeHistoryUB
        Dim lLen As Int32 = 0
        Dim lFinalCnt As Int32 = 0
        Dim dtNow As Date = Now
        Dim bExcludeByDate As Boolean = False

        For X As Int32 = lUB To 0 Step -1
            'If GetDateFromNumber(moTradeHistory(X).lTransactionDate).AddDays(15) > dtNow Then
            lLen += moTradeHistory(X).GetObjAsString().Length
            lFinalCnt += 1
            'End If
        Next X

        If lLen > 25000 Then
            lFinalCnt = 0
            lLen = 0
            bExcludeByDate = True
            For X As Int32 = lUB To 0 Step -1 '0 To lUB
                If GetDateFromNumber(moTradeHistory(X).lTransactionDate).AddDays(15) > dtNow Then
                    lLen += moTradeHistory(X).GetObjAsString().Length
                    lFinalCnt += 1
                    If lLen > 25000 Then Exit For
                End If
            Next X
        End If

        Dim yMsg(9 + lLen) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetTradeHistory).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lFinalCnt).CopyTo(yMsg, lPos) : lPos += 4
        Dim lTmpCnt As Int32 = 0
        For X As Int32 = lUB To 0 Step -1 '0 To lUB
            If bExcludeByDate = False OrElse GetDateFromNumber(moTradeHistory(X).lTransactionDate).AddDays(15) > dtNow Then
                moTradeHistory(X).GetObjAsString.CopyTo(yMsg, lPos) : lPos += moTradeHistory(X).GetObjAsString.Length
                lTmpCnt += 1
                If lTmpCnt = lFinalCnt Then Exit For
            End If
        Next X
        Return yMsg
    End Function
    Public Sub LoadTradeHistoryItem(ByVal lOtherPlayerID As Int32, ByVal lTransDate As Int32, ByVal sItemName As String, ByVal blQty As Int64, ByVal yType As TradeHistory.TradeHistoryItemType)
        For X As Int32 = 0 To mlTradeHistoryUB
            With moTradeHistory(X)
                If .lOtherPlayerID = lOtherPlayerID AndAlso .lTransactionDate = lTransDate Then
                    .AddTradeItem(StringToBytes(sItemName), blQty, yType)
                    Exit For
                End If
            End With
        Next X
    End Sub
    Public Sub DeleteTradeHistoryItem(ByVal lTransDate As Int32, ByVal lOtherPlayerID As Int32, ByVal yResult As Byte, ByVal yEventType As Byte)
        For X As Int32 = 0 To mlTradeHistoryUB
            Dim oTH As TradeHistory = moTradeHistory(X)
            If oTH Is Nothing = False Then
                If oTH.lTransactionDate = lTransDate AndAlso oTH.lOtherPlayerID = lOtherPlayerID AndAlso oTH.yTradeResult = yResult AndAlso oTH.yTradeEventType = yEventType Then
                    'ok, found it...
                    For Y As Int32 = X To mlTradeHistoryUB - 1
                        moTradeHistory(Y) = moTradeHistory(Y + 1)
                    Next Y
                    mlTradeHistoryUB -= 1
                    Exit For
                End If
            End If
        Next X

    End Sub
    Public Sub RemoveTradeHistoryItemsOfPlayer(ByVal lOther As Int32)
        Try
            For X As Int32 = 0 To mlTradeHistoryUB
                If moTradeHistory(X) Is Nothing = False Then
                    If moTradeHistory(X).lOtherPlayerID = lOther Then
                        moTradeHistory(X).lOtherPlayerID = -1
                        moTradeHistory(X).ResetMsg()
                        moTradeHistory(X).SaveObject()
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub
#End Region

#Region "  Wormhole Discoveries  "
    'Contains everything needed for wormhole knowledge management
    Public oWormholes() As Wormhole
    Public lWormholeUB As Int32 = -1

    Public Sub AddWormholeKnowledge(ByRef oWormhole As Wormhole, ByVal bAlertPlayer As Boolean, ByRef oSystem As SolarSystem, ByVal bSaveItem As Boolean)
        For X As Int32 = 0 To lWormholeUB
            If oWormholes(X) Is Nothing = False AndAlso oWormholes(X).ObjectID = oWormhole.ObjectID Then
                Return
            End If
        Next X

        lWormholeUB += 1
        ReDim Preserve oWormholes(lWormholeUB)
        oWormholes(lWormholeUB) = oWormhole

        If bSaveItem = True AndAlso SavePlayerWormhole(lWormholeUB) = False Then
            LogEvent(LogEventType.CriticalError, "Unable to save player's knowledge of wormhole: " & Me.ObjectID & ", " & oWormhole.ObjectID)
        End If

        If bAlertPlayer = True Then

            'If alert player is true, this is a genuine first discovery, let's inform the other servers
            Dim yMsg(15) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eForcePrimarySync).CopyTo(yMsg, lPos) : lPos += 2
            oWormhole.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(oSystem.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

            'Ok, send the player an email
            If Me.lConnectedPrimaryID > -1 OrElse Me.HasOnlineAliases(AliasingRights.eChangeEnvironment) = True Then
                Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oWormhole, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eChangeEnvironment)
            End If

            Dim oSB As New System.Text.StringBuilder()
            oSB.AppendLine("We have detected a spatial anomaly in " & BytesToString(oSystem.SystemName) & ".")
            oSB.AppendLine("Our scientists are anxious to begin studies on it.")

            If oSystem Is Nothing = False Then oSB.AppendLine("The location has been attached as a waypoint to this message.")

            Dim uAttach() As PlayerComm.WPAttachment = Nothing
            If oSystem Is Nothing = False Then
                ReDim uAttach(0)

                Dim lX As Int32
                Dim lZ As Int32
                If oWormhole.System1 Is Nothing = False AndAlso oWormhole.System1.ObjectID = oSystem.ObjectID Then
                    lX = oWormhole.LocX1
                    lZ = oWormhole.LocY1
                Else
                    lX = oWormhole.LocX2
                    lZ = oWormhole.LocY2
                End If

                With uAttach(0)
                    'oPC.AddEmailAttachment(0, oSystem.ObjectID, oSystem.ObjTypeID, lX, lZ, "Spatial Anomaly")
                    .AttachNumber = 0
                    .EnvirID = oSystem.ObjectID
                    .EnvirTypeID = oSystem.ObjTypeID
                    .LocX = lX
                    .LocZ = lZ
                    .sWPName = "Spatial Anomaly"
                End With
            End If

            Dim oPC As PlayerComm = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Anomaly Detected", Me.ObjectID, GetDateAsNumber(Now), False, BytesToString(Me.PlayerName), uAttach)
            If oPC Is Nothing = False Then
                Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eChangeEnvironment)
            End If
        End If
    End Sub
    Public Function HasWormholeKnowledge(ByVal lWormholeID As Int32) As Boolean
        For X As Int32 = 0 To lWormholeUB
            If oWormholes(X) Is Nothing = False AndAlso oWormholes(X).ObjectID = lWormholeID Then
                Return True
            End If
        Next X
        Return False
    End Function
#End Region

#Region "  Faction Slot Management  "
    Public lSlotID(4) As Int32
    Public ySlotState(4) As Byte
    Public lFactionID(2) As Int32       'for quick reference backwards...
    Public Function GetSlotStateMsg() As Byte()
        Dim yMsg(45) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateSlotStates).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

        For X As Int32 = 0 To 4
            System.BitConverter.GetBytes(lSlotID(X)).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = ySlotState(X) : lPos += 1
        Next X

        For X As Int32 = 0 To 2
            System.BitConverter.GetBytes(lFactionID(X)).CopyTo(yMsg, lPos) : lPos += 4
            Dim yVal As Byte = 0
            If lFactionID(X) > 0 Then
                Dim oTmpPlayer As Player = GetEpicaPlayer(lFactionID(X))
                If oTmpPlayer Is Nothing = False Then
                    For Y As Int32 = 0 To 4
                        If oTmpPlayer.lSlotID(Y) = Me.ObjectID Then
                            yVal = oTmpPlayer.ySlotState(Y)
                            Exit For
                        End If
                    Next Y
                End If
                oTmpPlayer = Nothing
            End If
            yMsg(lPos) = yVal : lPos += 1
        Next X

        Return yMsg
    End Function

    Private mlPreviousFactionResearchLookup As Int32 = 0
    Private mfFactionResearchTimeMult As Single = Single.MinValue
    Public ReadOnly Property fFactionResearchTimeMultiplier() As Single
        Get
            If glCurrentCycle - mlPreviousFactionResearchLookup > 150 OrElse mfFactionResearchTimeMult = Single.MinValue Then       'every 5 seconds
                mlPreviousFactionResearchLookup = glCurrentCycle

                Dim fMult As Single
                Dim yTestTitle As Byte = Me.yPlayerTitle
                If (yTestTitle And Player.PlayerRank.ExRankShift) = 0 Then
                    Select Case yTestTitle
                        Case Player.PlayerRank.Emperor
                            fMult = 0.85F
                        Case Player.PlayerRank.King
                            fMult = 0.76F
                        Case Player.PlayerRank.Baron
                            fMult = 0.81F
                        Case Player.PlayerRank.Duke
                            fMult = 0.96F
                        Case Player.PlayerRank.Overseer
                            fMult = 0.99F
                        Case Player.PlayerRank.Governor
                            fMult = 0.99F
                    End Select
                End If


                Dim fValue As Single = 1.0F
                For X As Int32 = 0 To 4
                    If lSlotID(X) > 0 AndAlso (ySlotState(X) And eySlotState.Accepted) <> 0 AndAlso ((eySlotState.RankTooHigh Or eySlotState.InsufficientFactionSlots Or eySlotState.WarNotJoined Or eySlotState.ExRankMember) And ySlotState(X)) = 0 Then
                        fValue *= fMult ' 0.9F
                    End If
                Next X
                For X As Int32 = 0 To 2
                    If lFactionID(X) > 0 Then
                        Dim oTmpPlayer As Player = GetEpicaPlayer(lFactionID(X))
                        If oTmpPlayer Is Nothing = False Then
                            For Y As Int32 = 0 To 4
                                If oTmpPlayer.lSlotID(Y) = Me.ObjectID Then
                                    Dim yValue As Byte = oTmpPlayer.ySlotState(Y)
                                    If (yValue And eySlotState.Accepted) <> 0 AndAlso ((eySlotState.FactionAtWar Or eySlotState.RankTooHigh Or eySlotState.InsufficientFactionSlots Or eySlotState.ExRankMember) And yValue) = 0 Then
                                        fValue *= 0.5F
                                    End If
                                    Exit For
                                End If
                            Next Y
                        End If
                    End If
                Next X

                Dim bRefresh As Boolean = mfFactionResearchTimeMult <> fValue
                mfFactionResearchTimeMult = fValue

                'now, go thru and recalc all facility researchers...
                If bRefresh = True Then
                    Dim lCurUB As Int32 = Math.Min(glFacilityUB, glFacilityIdx.GetUpperBound(0))
                    For X As Int32 = 0 To lCurUB
                        If glFacilityIdx(X) > 0 Then
                            Dim oFac As Facility = goFacility(X)
                            If oFac Is Nothing = False AndAlso oFac.yProductionType = ProductionType.eResearch AndAlso oFac.Owner.ObjectID = Me.ObjectID Then
                                oFac.RecalcProduction()
                            End If
                        End If
                    Next X
                End If
            End If

            Return mfFactionResearchTimeMult
        End Get
    End Property

    Public Function TradesAreFree(ByVal lPlayerID As Int32) As Boolean
        For X As Int32 = 0 To 4
            If lSlotID(X) = lPlayerID AndAlso (ySlotState(X) And eySlotState.Accepted) <> 0 AndAlso ((eySlotState.RankTooHigh Or eySlotState.InsufficientFactionSlots Or eySlotState.WarNotJoined Or eySlotState.ExRankMember) And ySlotState(X)) = 0 Then
                Return True
            End If
        Next X
        For X As Int32 = 0 To 2
            If lFactionID(X) = lPlayerID Then
                Dim oTmpPlayer As Player = GetEpicaPlayer(lFactionID(X))
                If oTmpPlayer Is Nothing = False Then
                    For Y As Int32 = 0 To 4
                        If oTmpPlayer.lSlotID(Y) = Me.ObjectID Then
                            Dim yValue As Byte = oTmpPlayer.ySlotState(Y)
                            If (yValue And eySlotState.Accepted) <> 0 AndAlso ((eySlotState.FactionAtWar Or eySlotState.RankTooHigh Or eySlotState.InsufficientFactionSlots Or eySlotState.ExRankMember) And yValue) = 0 Then
                                Return True
                            End If
                            Exit For
                        End If
                    Next Y
                End If
            End If
        Next X
        Return False
    End Function

    Public Sub ReverifySlots()

        Dim bIAmExRank As Boolean = (yPlayerTitle And PlayerRank.ExRankShift) <> 0

        Dim lMyTitle As Int32 = Me.yPlayerTitle + 1

        Dim lValidSlotCnt As Int32 = 0

        For X As Int32 = 0 To 4
            If lSlotID(X) > 0 AndAlso ySlotState(X) <> eySlotState.Unaccepted Then
                Dim yNewState As Byte = eySlotState.Accepted
                Dim oPlayer As Player = GetEpicaPlayer(lSlotID(X))
                If oPlayer Is Nothing = False Then
                    Dim bTheyAreExRank As Boolean = (oPlayer.yPlayerTitle And PlayerRank.ExRankShift) <> 0

                    If bIAmExRank = True OrElse bTheyAreExRank = True Then yNewState = yNewState Or eySlotState.ExRankMember

                    Dim lTheirTitle As Int32 = oPlayer.yPlayerTitle + 1
                    If lMyTitle + lTheirTitle > 9 OrElse lMyTitle <= lTheirTitle Then
                        'Emperors are allowed to have Overseers... so if either i am not an emperor or they are not an overseer, then the rank is too high
                        If (lMyTitle <> 7 OrElse lTheirTitle <> 3) Then
                            yNewState = yNewState Or eySlotState.RankTooHigh
                        End If
                    End If

                    'ok, here it is... 
                    'items in the SLOTS are Lesser empires and I would be the greater
                    'items in the FACTIONS are Greater empires and I am the lesser

                    'when the lesser empire is at war with an empire the greater is not, the greater loses the benefit (but the lesser retains)
                    For Y As Int32 = 0 To oPlayer.PlayerRelUB
                        Dim oTmpRel As PlayerRel = oPlayer.GetPlayerRelByIndex(Y)
                        If oTmpRel Is Nothing = False AndAlso oTmpRel.WithThisScore <= elRelTypes.eWar Then
                            'ok, make sure I am at war...
                            Dim lOtherP As Int32 = oTmpRel.oThisPlayer.ObjectID
                            If lOtherP = oPlayer.ObjectID Then lOtherP = oTmpRel.oPlayerRegards.ObjectID
                            If Me.GetPlayerRelScore(lOtherP) > elRelTypes.eWar Then
                                yNewState = yNewState Or eySlotState.WarNotJoined
                                Exit For
                            End If
                        End If
                    Next Y

                    'when the greater empire is at war with an empire that the lesser is slotted, the lesser loses the benefit (but the greater retains)
                    For Y As Int32 = 0 To 2
                        If oPlayer.lFactionID(Y) > 0 AndAlso oPlayer.lFactionID(Y) <> Me.ObjectID Then
                            'Ok, test our relationship
                            If Me.GetPlayerRelScore(oPlayer.lFactionID(Y)) <= elRelTypes.eWar Then
                                'Make sure they confirmed
                                Dim oTmpPlayer As Player = GetEpicaPlayer(oPlayer.lFactionID(Y))
                                If oTmpPlayer Is Nothing = False Then
                                    For Z As Int32 = 0 To 4
                                        If oTmpPlayer.lSlotID(Z) = oPlayer.ObjectID Then
                                            If oTmpPlayer.ySlotState(Z) <> eySlotState.Unaccepted Then
                                                yNewState = yNewState Or eySlotState.FactionAtWar
                                                Exit For
                                            End If
                                        End If
                                    Next Z
                                    If (yNewState And eySlotState.FactionAtWar) <> 0 Then Exit For
                                End If
                            End If
                        End If
                    Next Y

                    Dim lOtherCnt As Int32 = 0
                    Dim yOtherTitle As Byte = oPlayer.yPlayerTitle
                    If yOtherTitle = PlayerRank.Magistrate Then
                        lOtherCnt = 3
                    ElseIf yOtherTitle = PlayerRank.Governor Then
                        lOtherCnt = 2
                    ElseIf yOtherTitle = PlayerRank.Overseer OrElse yOtherTitle = PlayerRank.Duke Then
                        lOtherCnt = 1
                    End If
                    Dim lOtherFactionCnt As Int32 = 0
                    For Y As Int32 = 0 To 2
                        If oPlayer.lFactionID(Y) > 0 Then
                            Dim oTmpPlayer As Player = GetEpicaPlayer(oPlayer.lFactionID(Y))
                            For Z As Int32 = 0 To 4
                                If oTmpPlayer.lSlotID(Z) = oPlayer.ObjectID Then
                                    If oTmpPlayer.ySlotState(Z) <> eySlotState.Unaccepted Then
                                        lOtherFactionCnt += 1
                                        Exit For
                                    End If
                                End If
                            Next Z
                        End If
                    Next Y
                    If lOtherFactionCnt > lOtherCnt Then
                        yNewState = yNewState Or eySlotState.InsufficientFactionSlots
                    End If

                    ySlotState(X) = yNewState
                    If yNewState = eySlotState.Accepted Then lValidSlotCnt += 1
                Else
                    lSlotID(X) = -1
                    Continue For
                End If
            End If
        Next X

        Dim lFactionCnt As Int32 = 0
        For X As Int32 = 0 To 2
            If lFactionID(X) > 0 Then
                Dim oPlayer As Player = GetEpicaPlayer(lFactionID(X))
                If oPlayer Is Nothing = False Then
                    For Y As Int32 = 0 To 4
                        If oPlayer.lSlotID(Y) = Me.ObjectID Then
                            Dim bTheyAreExRank As Boolean = (oPlayer.yPlayerTitle And PlayerRank.ExRankShift) <> 0
                            If bTheyAreExRank = True OrElse bIAmExRank = True Then
                                oPlayer.ySlotState(Y) = oPlayer.ySlotState(Y) Or eySlotState.ExRankMember
                            End If
                            If (oPlayer.ySlotState(Y) And eySlotState.Accepted) <> 0 Then
                                lFactionCnt += 1
                            End If
                            Exit For
                        End If
                    Next Y
                End If
            End If
        Next X

        'Now, based on my title
        Dim lMaxCnt As Int32 = 0
        'Magistrate is permitted to be in 3 faction slots
        'Governor is permitted to be in 2 faction slots
        'Overseer and Duke are permitted to be in 1 faction slot
        Dim yTestTitle As Byte = yPlayerTitle
        If bIAmExRank = True Then yTestTitle = yTestTitle Xor PlayerRank.ExRankShift
        If yTestTitle = PlayerRank.Magistrate Then
            lMaxCnt = 3
        ElseIf yTestTitle = PlayerRank.Governor Then
            lMaxCnt = 2
        ElseIf yTestTitle = PlayerRank.Overseer OrElse yTestTitle = PlayerRank.Duke Then
            lMaxCnt = 1
        End If
        If lFactionCnt > lMaxCnt Then
            For X As Int32 = 0 To 2
                If lFactionID(X) > 0 Then
                    Dim oPlayer As Player = GetEpicaPlayer(lFactionID(X))
                    If oPlayer Is Nothing = False Then
                        For Y As Int32 = 0 To 4
                            If oPlayer.lSlotID(Y) = Me.ObjectID Then
                                oPlayer.ySlotState(Y) = oPlayer.ySlotState(Y) Or eySlotState.InsufficientFactionSlots
                                Exit For
                            End If
                        Next Y
                    End If
                End If
            Next X
        Else
            'ensure that no slots are marked as exceeded faction slots
            For X As Int32 = 0 To 2
                If lFactionID(X) > 0 Then
                    Dim oPlayer As Player = GetEpicaPlayer(lFactionID(X))
                    If oPlayer Is Nothing = False Then
                        For Y As Int32 = 0 To 4
                            If oPlayer.lSlotID(Y) = Me.ObjectID Then
                                If (oPlayer.ySlotState(Y) And eySlotState.InsufficientFactionSlots) <> 0 Then
                                    oPlayer.ySlotState(Y) = oPlayer.ySlotState(Y) Xor eySlotState.InsufficientFactionSlots
                                End If
                                Exit For
                            End If
                        Next Y
                    End If
                End If
            Next X
        End If

        'Ok, finally, send all players involved in slot management with me a message
        For X As Int32 = 0 To 4
            If lSlotID(X) > 0 AndAlso ySlotState(X) <> eySlotState.Unaccepted Then
                Dim oPlayer As Player = GetEpicaPlayer(lSlotID(X))
                If oPlayer Is Nothing = False Then
                    oPlayer.SendPlayerMessage(oPlayer.GetSlotStateMsg(), False, 0)
                End If
            End If
        Next X
        For X As Int32 = 0 To 2
            If lFactionID(X) > 0 Then
                Dim oPlayer As Player = GetEpicaPlayer(lFactionID(X))
                If oPlayer Is Nothing = False Then
                    oPlayer.SendPlayerMessage(oPlayer.GetSlotStateMsg(), False, 0)
                End If
            End If
        Next X

        'Now, the test for Counselor
        If (Me.lCustomTitlePermission And elCustomRankPermissions.Counselor) = 0 AndAlso lValidSlotCnt = 5 Then
            Dim lOriginal As Int32 = Me.lCustomTitlePermission
            Me.lCustomTitlePermission = Me.lCustomTitlePermission Or elCustomRankPermissions.Counselor
            ProcessCustomTitlePermissionChange(lOriginal)
        End If

        Me.SendPlayerMessage(Me.GetSlotStateMsg(), False, 0)
    End Sub

    Public Sub ClearFactionSlots()
        For X As Int32 = 0 To 4
            Dim lOther As Int32 = lSlotID(X)
            lSlotID(X) = 0
            If lOther > 0 Then
                Dim oOther As Player = GetEpicaPlayer(lOther)
                If oOther Is Nothing = False Then
                    For Y As Int32 = 0 To 2
                        If lFactionID(Y) = Me.ObjectID Then
                            lFactionID(Y) = -1
                        End If
                    Next Y
                    oOther.ReverifySlots()
                End If
            End If
        Next X
        For X As Int32 = 0 To 2
            Dim lOther As Int32 = lFactionID(X)
            lFactionID(X) = 0
            If lOther > 0 Then
                Dim oOther As Player = GetEpicaPlayer(lOther)
                If oOther Is Nothing = False Then
                    For Y As Int32 = 0 To 4
                        If lSlotID(Y) = Me.ObjectID Then
                            lSlotID(Y) = -1
                            ySlotState(Y) = 0
                        End If
                    Next Y
                    oOther.ReverifySlots()
                End If
            End If
        Next X
        ReverifySlots()
    End Sub
#End Region

#Region "  Player CP Penalty Class Management  "
    Public oCPPenalties() As PlayerCPPenalty
    Public lCPPenaltyUB As Int32 = -1

    Public Sub AddCPPenalty(ByRef oItem As PlayerCPPenalty)
        Dim lUB As Int32 = -1
        If oCPPenalties Is Nothing = False Then lUB = Math.Min(lCPPenaltyUB, oCPPenalties.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If oCPPenalties(X) Is Nothing = False Then
                If oCPPenalties(X).oDecPlayer.ObjectID = oItem.oDecPlayer.ObjectID Then Return
            End If
        Next X

        lCPPenaltyUB += 1
        ReDim Preserve oCPPenalties(lCPPenaltyUB)
        oCPPenalties(lCPPenaltyUB) = oItem
    End Sub
    Public Sub ClearCPPenaltyList()
        lCPPenaltyUB = -1
        ReDim oCPPenalties(-1)

        Dim sSQL As String = "DELETE FROM tblPlayerCPPenalty WHERE PlayerID = " & Me.ObjectID
        Dim oComm As OleDb.OleDbCommand = Nothing
        Try
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Player.ClearCPPenaltyList: " & ex.Message)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try
    End Sub
    Public Sub SendCPPenaltyList()
        Try
            If lConnectedPrimaryID > -1 OrElse HasOnlineAliases(AliasingRights.eViewUnitsAndFacilities) = True Then
                Dim lCnt As Int32 = 0
                Dim lUB As Int32 = -1
                If oCPPenalties Is Nothing = False Then lUB = Math.Min(lCPPenaltyUB, oCPPenalties.GetUpperBound(0))
                For X As Int32 = 0 To lUB
                    If oCPPenalties(X) Is Nothing = False Then lCnt += 1
                Next X
                If lCnt = 0 Then Return

                Dim yMsg(13 + (lCnt * 5)) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eGetCPPenaltyList).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(BadWarDecCPIncrease).CopyTo(yMsg, lPos) : lPos += 4

                Dim lCyclesRemain As Int32 = 0
                If BadWarDecCPIncreaseEndCycle > glCurrentCycle Then
                    lCyclesRemain = BadWarDecCPIncreaseEndCycle - glCurrentCycle
                Else : lCyclesRemain = -1
                End If

                System.BitConverter.GetBytes(lCyclesRemain).CopyTo(yMsg, lPos) : lPos += 4
                For X As Int32 = 0 To lUB
                    If oCPPenalties(X) Is Nothing = False Then
                        System.BitConverter.GetBytes(oCPPenalties(X).oDecPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        yMsg(lPos) = oCPPenalties(X).yPenalty : lPos += 1
                    End If
                Next X

                SendPlayerMessage(yMsg, False, AliasingRights.eViewUnitsAndFacilities)
            End If
        Catch ex As Exception
            LogEvent(LogEventType.Warning, "Player.SendCPPenaltyList: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "  Cargo Route Template Management  "
    Public oRouteTemplates() As RouteTemplate
    Public lRouteTemplateUB As Int32 = -1

    Public Function GetRouteTemplate(ByVal lTemplateID As Int32) As RouteTemplate
        Try
            For X As Int32 = 0 To lRouteTemplateUB
                If oRouteTemplates(X) Is Nothing = False AndAlso oRouteTemplates(X).lTemplateID = lTemplateID Then
                    Return oRouteTemplates(X)
                End If
            Next X
        Catch
        End Try
        Return Nothing
    End Function
    Public Sub RemoveTemplate(ByVal lID As Int32)
        For X As Int32 = 0 To lRouteTemplateUB
            If oRouteTemplates(X) Is Nothing = False AndAlso oRouteTemplates(X).lTemplateID = lID Then
                oRouteTemplates(X).DeleteMe()
                For Y As Int32 = X To lRouteTemplateUB - 1
                    oRouteTemplates(Y) = oRouteTemplates(Y + 1)
                Next Y
                lRouteTemplateUB -= 1
                Exit For
            End If
        Next X
    End Sub
    Public Function HandleGetRouteTemplateList() As Byte()

        Dim lCnt As Int32 = 0
        Dim lItemCnt As Int32 = 0
        For X As Int32 = 0 To lRouteTemplateUB
            If oRouteTemplates(X) Is Nothing = False Then
                lCnt += 1
                lItemCnt += (oRouteTemplates(X).lItemUB + 1)
            End If
        Next X

        Dim lTotalLen As Int32 = 5 + (lCnt * 28) + (lItemCnt * 27)

        If lTotalLen > 32000 Then Return Nothing
        Dim yMsg(lTotalLen) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetRouteTemplates).CopyTo(yMsg, lPos) : lPos += 2

        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        For X As Int32 = 0 To lRouteTemplateUB

            If oRouteTemplates(X) Is Nothing = False Then
                With oRouteTemplates(X)
                    System.BitConverter.GetBytes(.lTemplateID).CopyTo(yMsg, lPos) : lPos += 4
                    .TemplateName.CopyTo(yMsg, lPos) : lPos += 20

                    System.BitConverter.GetBytes(.lItemUB + 1).CopyTo(yMsg, lPos) : lPos += 4
                    For Y As Int32 = 0 To .lItemUB

                        With .uItems(Y)
                            If .oDest Is Nothing = False Then
                                .oDest.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                                If (.oDest.ObjTypeID = ObjectType.eUnit OrElse .oDest.ObjTypeID = ObjectType.eFacility) AndAlso CType(.oDest, Epica_Entity).ParentObject Is Nothing = False Then
                                    CType(CType(.oDest, Epica_Entity).ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                                Else
                                    System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                    System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
                                End If
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
                            End If
                            If .oLoadItem Is Nothing = False Then
                                .oLoadItem.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
                            End If
                            System.BitConverter.GetBytes(.lLocX).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.lLocZ).CopyTo(yMsg, lPos) : lPos += 4

                            yMsg(lPos) = .yExtraFlags : lPos += 1
                        End With
                    Next Y
                End With
            End If
        Next X

        Return yMsg
    End Function
    Public Sub AddTemplate(ByRef oTemplate As RouteTemplate)
        For X As Int32 = 0 To lRouteTemplateUB
            If oRouteTemplates(X) Is Nothing Then
                oRouteTemplates(X) = oTemplate
                Return
            End If
        Next X
        ReDim Preserve oRouteTemplates(lRouteTemplateUB + 1)
        oRouteTemplates(lRouteTemplateUB + 1) = oTemplate
        lRouteTemplateUB += 1
    End Sub
#End Region

#Region "  Transport Management  "
    Private moTransports() As Transport
    Private mlTransportIdx() As Int32
    Private mlTransportUB As Int32 = -1

    Private mdtNextTransportEvent As DateTime = DateTime.MinValue
    Private mbClearNextTransportEvent As Boolean = False
    Public Sub ProcessTransportEvents(ByVal dtNow As DateTime)
        If mdtNextTransportEvent = DateTime.MinValue Then
            Dim dtNewVal As DateTime = DateTime.MaxValue
            Dim lUB As Int32 = -1
            If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportUB, mlTransportIdx.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If mlTransportIdx(X) > -1 Then
                    Dim oT As Transport = moTransports(X)
                    If oT Is Nothing = False AndAlso (oT.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                        If oT.ETA <> DateTime.MinValue AndAlso oT.ETA < dtNewVal Then
                            dtNewVal = oT.ETA
                        End If
                    End If
                End If
            Next X

            mdtNextTransportEvent = dtNewVal
        ElseIf mdtNextTransportEvent < dtNow Then
            'Ok, check our list
            Dim lCnt As Int32 = 0
            Dim lUB As Int32 = -1
            If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportUB, mlTransportIdx.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If mlTransportIdx(X) > -1 Then
                    Dim oT As Transport = moTransports(X)
                    If oT Is Nothing = False AndAlso (oT.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 AndAlso oT.ETA < dtNow Then
                        'Do it...
                        oT.TransportArrived()
                        mdtNextTransportEvent = DateTime.MinValue
                        lCnt += 1
                        If lCnt > 10 Then Exit For
                    End If
                End If
            Next X
        End If

        If mbClearNextTransportEvent = True Then
            mdtNextTransportEvent = DateTime.MinValue
            mbClearNextTransportEvent = False
        End If
    End Sub
    Public Sub ClearNextTransportEvent()
        mbClearNextTransportEvent = True
    End Sub

    Public Function HandleRequestTransports() As Byte()
        Dim lUB As Int32 = -1
        If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportUB, mlTransportIdx.GetUpperBound(0))
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lUB
            If mlTransportIdx(X) > -1 Then
                Dim oT As Transport = moTransports(X)
                If oT Is Nothing = False Then
                    lCnt += 1
                End If
            End If
        Next X

        Dim yMsg(5 + (lCnt * 17)) As Byte
        Dim lPos As Int32 = 0
        Dim dtNow As DateTime = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime()

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransports).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        For X As Int32 = 0 To lUB
            If mlTransportIdx(X) > -1 Then
                Dim oT As Transport = moTransports(X)
                If oT Is Nothing = False Then
                    With oT
                        System.BitConverter.GetBytes(.TransportID).CopyTo(yMsg, lPos) : lPos += 4
                        Dim lSecs As Int32 = 0
                        If (.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                            If .ETA > dtNow Then
                                If .ETA.Subtract(dtNow).TotalSeconds > Int32.MaxValue \ 2 Then lSecs = Int32.MaxValue \ 2 Else lSecs = CInt((.ETA.Subtract(dtNow).TotalSeconds))
                            End If
                        End If
                        System.BitConverter.GetBytes(lSecs).CopyTo(yMsg, lPos) : lPos += 4
                        yMsg(lPos) = .TransFlags : lPos += 1
                        If (.TransFlags And Transport.elTransportFlags.eIdle) <> 0 Then
                            System.BitConverter.GetBytes(.LocationID).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.LocationTypeID).CopyTo(yMsg, lPos) : lPos += 2
                        Else
                            System.BitConverter.GetBytes(.DestinationID).CopyTo(yMsg, lPos) : lPos += 4
                            System.BitConverter.GetBytes(.DestinationTypeID).CopyTo(yMsg, lPos) : lPos += 2
                        End If
                        If .oUnitDef Is Nothing = False Then
                            System.BitConverter.GetBytes(.oUnitDef.ModelID).CopyTo(yMsg, lPos) : lPos += 2
                        Else
                            System.BitConverter.GetBytes(0S).CopyTo(yMsg, lPos) : lPos += 2
                        End If
                    End With
                End If
            End If
        Next X

        Return yMsg
    End Function
    Public Function RemoveTransport(ByVal lTransportID As Int32) As Boolean
        Dim lUB As Int32 = -1
        If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportUB, mlTransportIdx.GetUpperBound(0))
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lUB
            If mlTransportIdx(X) = lTransportID Then
                Dim oT As Transport = moTransports(X)
                If oT Is Nothing = False Then
                    If (oT.TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
                        Return False
                    Else
                        oT.DeleteMe()
                        mlTransportIdx(X) = -1
                        Return True
                    End If
                End If
            End If
        Next X
        Return False
    End Function

    Public Function GetTransport(ByVal lID As Int32) As Transport
        Dim lUB As Int32 = -1
        If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportUB, mlTransportIdx.GetUpperBound(0))
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lUB
            If mlTransportIdx(X) > -1 Then
                Dim oT As Transport = moTransports(X)
                If oT Is Nothing = False Then
                    If oT.TransportID = lID Then
                        Return oT
                    End If
                End If
            End If
        Next X
        Return Nothing
    End Function

    Public Sub AddTransport(ByRef oTrans As Transport)
        Dim lIdx As Int32 = mlTransportUB + 1
        ReDim Preserve moTransports(lIdx)
        ReDim Preserve mlTransportIdx(lIdx)
        moTransports(lIdx) = oTrans
        mlTransportIdx(lIdx) = oTrans.TransportID
        mlTransportUB += 1
    End Sub

    Public Function TransportCount() As Int32
        Dim lUB As Int32 = -1
        If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportUB, mlTransportIdx.GetUpperBound(0))
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lUB
            If mlTransportIdx(X) > -1 Then
                Dim oT As Transport = moTransports(X)
                If oT Is Nothing = False Then
                    lCnt += 1
                End If
            End If
        Next X
        Return lCnt
    End Function
#End Region

    '#Region "  Warpoint Management  "
    '    Public lLastWPUpkeep As Int32 = 0               'based on TotalTimePlayed
    '    Public lLastGuildShareUpkeep As Int32 = 0       'based on TIME
    '    Public lLastWPUpkeepTime As Int32 = 0           'based on TIME - 0 indicates it has never been set - when a player logs in for the first time/respawn, set this to now
    '    Public lTemporaryWarpointCheckFlag As elWarpointCheckFlag = elWarpointCheckFlag.eNoCheckRequired
    '    Public lTemporaryWarpointUpkeepCost As Int32 = 0
    '    Public blWarpoints As Int64 = 0
    '    Public blWarpointsAllTime As Int64 = 0
    '    Public lCurrentWarpointUpkeepCost As Int32 = 0
    '    Public dtLastWPUpkeep As DateTime
    '    Public Function CheckWPUpkeep(ByVal lNow As Int32, ByVal dtNow As DateTime) As elWarpointCheckFlag
    '        Dim lResult As elWarpointCheckFlag = elWarpointCheckFlag.eNoCheckRequired

    '        If lLastWPUpkeepTime < 1 Then
    '            If Me.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then Return lResult
    '            If Me.LastLogin.Year < 2010 Then Return lResult
    '            If Me.LastLogin.Month < 6 Then Return lResult
    '            If Me.LastLogin.Year = 2010 AndAlso Me.LastLogin.Month = 6 AndAlso Me.LastLogin.Day < 21 Then Return lResult

    '            lLastWPUpkeepTime = lNow
    '            dtLastWPUpkeep = GetDateFromNumber(lNow)
    '        End If

    '        If (TotalPlayTime - lLastWPUpkeep > 86400 AndAlso Me.ObjectID <> 3510 AndAlso Me.ObjectID <> 20611) OrElse dtNow.Subtract(dtLastWPUpkeep).Days > 30 Then
    '            '24 hour play period
    '            lLastWPUpkeep = TotalPlayTime
    '            lLastWPUpkeepTime = lNow
    '            dtLastWPUpkeep = GetDateFromNumber(lNow)
    '            'all units in the empire are required to pay WP upkeep
    '            '   Ships that are up for sale do not cost wp's to upkeep
    '            lResult = lResult Or elWarpointCheckFlag.eCheckAllNormalUnits
    '        End If

    '        'Units marked as guild share always calculate every 24 hours
    '        Dim lTemp As Int32 = lNow - lLastGuildShareUpkeep
    '        If lTemp >= 10000 AndAlso (lResult And elWarpointCheckFlag.eCheckAllNormalUnits) = 0 Then
    '            'Ok, first, check we're not rolling over years or months, which is any value over 100k
    '            If lTemp > 100000 AndAlso lLastGuildShareUpkeep > 0 Then
    '                'have to do this manually
    '                Dim dtLast As DateTime = GetDateFromNumber(lLastGuildShareUpkeep)
    '                Dim ts As TimeSpan = dtNow.Subtract(dtLast)
    '                If ts.TotalHours > 24 Then
    '                    '24 hour total period
    '                    lLastGuildShareUpkeep = lNow
    '                    'Any guild share units must be calculated now
    '                    lResult = lResult Or elWarpointCheckFlag.eCheckAllGuildShare
    '                End If
    '            Else
    '                '24 hour total period
    '                lLastGuildShareUpkeep = lNow
    '                'Any guild share units must be calculated now
    '                lResult = lResult Or elWarpointCheckFlag.eCheckAllGuildShare
    '            End If
    '        End If

    '        Return lResult
    '    End Function
    '    Public Sub ProcessWarpointReduction()
    '        If lTemporaryWarpointUpkeepCost > 250 OrElse ((lTemporaryWarpointCheckFlag And elWarpointCheckFlag.eCheckAllGuildShare) <> 0 AndAlso lTemporaryWarpointUpkeepCost > 0) Then
    '            Dim blCurrentWP As Int64 = blWarpoints
    '            blWarpoints -= lTemporaryWarpointUpkeepCost

    '            Dim oPC As PlayerComm = Nothing

    '            Dim sMsgHdr As String = "Warpoints were due for"
    '            If (Me.lTemporaryWarpointCheckFlag And elWarpointCheckFlag.eCheckAllGuildShare) <> 0 Then
    '                sMsgHdr &= " Guild Shared Units"
    '            End If
    '            If (Me.lTemporaryWarpointCheckFlag And elWarpointCheckFlag.eCheckAllNormalUnits) <> 0 Then
    '                If (Me.lTemporaryWarpointCheckFlag And elWarpointCheckFlag.eCheckAllGuildShare) <> 0 Then
    '                    sMsgHdr &= " and Normal Units"
    '                Else
    '                    sMsgHdr &= " Normal Units"
    '                End If
    '            End If
    '            sMsgHdr &= " upkeep. Total warpoint upkeep cost for this was " & lTemporaryWarpointUpkeepCost.ToString("#,##0") & ". At the time of processing warpoints, you had " & blCurrentWP.ToString("#,##0") & " total warpoints."

    '            If blWarpoints < 0 Then
    '                Dim blOriginal As Int64 = Math.Abs(blWarpoints)
    '                Dim blRemains As Int64 = blOriginal
    '                blWarpoints = 0

    '                'Now, the penalties
    '                ' all units are sorted by xp level and then "age" of unit
    '                ' then, for each unit while blremains greater than 0, the unit is reduced in xp level
    '                '   if the unit is green (xp level = 0) then the unit is destroyed (potentially recycled)

    '                Dim lUB As Int32 = -1
    '                If glUnitIdx Is Nothing = False Then lUB = Math.Min(glUnitIdx.GetUpperBound(0), glUnitUB)

    '                Dim lSortedIdx(-1) As Int32
    '                Dim ySortedXPLvl(-1) As Int32
    '                Dim lSortedID(-1) As Int32
    '                Dim lSortedUB As Int32 = -1

    '                For X As Int32 = 0 To lUB
    '                    If glUnitIdx(X) > 0 Then
    '                        Dim oUnit As Unit = goUnit(X)
    '                        If oUnit Is Nothing = False Then
    '                            If oUnit.Owner Is Nothing = False AndAlso oUnit.Owner.ObjectID = Me.ObjectID AndAlso oUnit.bUnitInSellOrder = False Then

    '                                Dim yLvl As Byte = oUnit.ExpLevel
    '                                Dim lIdx As Int32 = -1

    '                                For Y As Int32 = 0 To lSortedUB
    '                                    If ySortedXPLvl(Y) < yLvl Then
    '                                        lIdx = Y
    '                                        Exit For
    '                                    ElseIf ySortedXPLvl(Y) = yLvl Then
    '                                        If lSortedID(Y) < oUnit.ObjectID Then
    '                                            lIdx = Y
    '                                            Exit For
    '                                        End If
    '                                    End If
    '                                Next Y

    '                                lSortedUB += 1
    '                                If lSortedUB > lSortedIdx.GetUpperBound(0) Then
    '                                    ReDim Preserve lSortedIdx(lSortedUB + 1000)
    '                                    ReDim Preserve ySortedXPLvl(lSortedUB + 1000)
    '                                    ReDim Preserve lSortedID(lSortedUB + 1000)
    '                                End If
    '                                If lIdx = -1 Then
    '                                    lIdx = lSortedUB
    '                                Else
    '                                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
    '                                        lSortedIdx(Y) = lSortedIdx(Y - 1)
    '                                        ySortedXPLvl(Y) = ySortedXPLvl(Y - 1)
    '                                        lSortedID(Y) = lSortedID(Y - 1)
    '                                    Next Y
    '                                End If
    '                                lSortedIdx(lIdx) = X
    '                                ySortedXPLvl(lIdx) = yLvl
    '                                lSortedID(lIdx) = oUnit.ObjectID
    '                            End If
    '                        End If
    '                    End If
    '                Next X

    '                'Now, process our effects on the sorted list
    '                Dim oSB_DroppedRank As New System.Text.StringBuilder
    '                Dim oSB_Destroyed As New System.Text.StringBuilder

    '                For X As Int32 = 0 To lSortedUB
    '                    Dim lIdx As Int32 = lSortedIdx(X)
    '                    If glUnitIdx(lIdx) = lSortedID(X) Then
    '                        Dim oUnit As Unit = goUnit(lIdx)
    '                        If oUnit Is Nothing = False Then
    '                            If oUnit.ObjectID = lSortedID(X) Then
    '                                Dim lUpkeep As Int32 = oUnit.EntityDef.WarpointUpkeep

    '                                Dim sParent As String = ""
    '                                If oUnit.ParentObject Is Nothing = False Then
    '                                    With CType(oUnit.ParentObject, Epica_GUID)
    '                                        If .ObjTypeID = ObjectType.ePlanet Then
    '                                            sParent = BytesToString(CType(oUnit.ParentObject, Planet).PlanetName)
    '                                        ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
    '                                            sParent = BytesToString(CType(oUnit.ParentObject, SolarSystem).SystemName)
    '                                        ElseIf .ObjTypeID = ObjectType.eGalaxy Then
    '                                            sParent = "Between systems"
    '                                        ElseIf .ObjTypeID = ObjectType.eUnit OrElse .ObjTypeID = ObjectType.eFacility Then
    '                                            sParent = "Docked"
    '                                        End If
    '                                    End With
    '                                End If
    '                                If sParent <> "" Then sParent = BytesToString(oUnit.EntityName) & " (" & sParent & ")" Else sParent = BytesToString(oUnit.EntityName)

    '                                'Now, reduce its rank
    '                                Dim lLvl As Int32 = oUnit.ExpLevel
    '                                If lLvl < 25 Then
    '                                    'Destroy the unit
    '                                    oSB_Destroyed.AppendLine("   " & sParent)

    '                                    LogEvent(LogEventType.ExtensiveLogging, "Removing Unit due to Warpoint Upkeep dues: " & oUnit.ObjectID.ToString)
    '                                    DestroyEntity(CType(oUnit, Epica_Entity), True, oUnit.Owner.ObjectID, False, "WPUpkeepOverage")
    '                                Else
    '                                    'demote unit
    '                                    oSB_DroppedRank.AppendLine("   " & sParent)

    '                                    lLvl -= 25
    '                                    If lLvl < 0 Then lLvl = 0
    '                                    oUnit.ExpLevel = CByte(lLvl)

    '                                    If oUnit.ParentObject Is Nothing = False Then
    '                                        Dim iParentType As Int16 = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
    '                                        If iParentType = ObjectType.ePlanet OrElse iParentType = ObjectType.eSolarSystem Then
    '                                            Dim yMsg(8) As Byte
    '                                            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateUnitExpLevel).CopyTo(yMsg, 0)
    '                                            oUnit.GetGUIDAsString.CopyTo(yMsg, 2)
    '                                            yMsg(8) = oUnit.ExpLevel
    '                                            If iParentType = ObjectType.ePlanet Then
    '                                                CType(oUnit.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
    '                                            Else
    '                                                CType(oUnit.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
    '                                            End If
    '                                        End If
    '                                    End If
    '                                End If

    '                                blRemains -= lUpkeep
    '                                If blRemains <= 0 Then Exit For
    '                            End If
    '                        End If
    '                    End If
    '                Next X

    '                If lSortedUB > -1 Then
    '                    Dim oSB As New System.Text.StringBuilder
    '                    oSB.AppendLine(sMsgHdr)
    '                    oSB.AppendLine()
    '                    oSB.AppendLine("You were unable to pay your warpoint upkeep (" & blOriginal.ToString("#,##0") & " deficit). Due to this deficit, these penalties were imposed:")
    '                    oSB.AppendLine()

    '                    If oSB_DroppedRank.Length > 0 Then
    '                        oSB.AppendLine("The following units were dropped in experience rank:")
    '                        oSB.AppendLine(oSB_DroppedRank.ToString)
    '                    End If
    '                    If oSB_Destroyed.Length > 0 Then
    '                        If oSB_DroppedRank.Length > 0 Then oSB.AppendLine()
    '                        oSB.AppendLine("The following units were destroyed because they were already green:")
    '                        oSB.AppendLine(oSB_Destroyed.ToString)
    '                    End If

    '                    oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Warpoint Penalty", Me.ObjectID, GetDateAsNumber(Now), False, sPlayerNameProper, Nothing)
    '                End If
    '            Else
    '                Dim oSB As New System.Text.StringBuilder
    '                oSB.AppendLine(sMsgHdr)
    '                oSB.AppendLine()
    '                oSB.AppendLine("You were able to pay your warpoint upkeep and no penalties were imposed.")

    '                oPC = Me.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Warpoints Processed", Me.ObjectID, GetDateAsNumber(Now), False, sPlayerNameProper, Nothing)
    '            End If

    '            If oPC Is Nothing = False Then
    '                Me.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
    '            End If
    '        End If


    '        lTemporaryWarpointUpkeepCost = 0
    '    End Sub

    '    Public Sub RegenerateWPVValues()
    '        For X As Int32 = 0 To PlayerRelUB
    '            If mlPlayerRelIdx(X) > 0 AndAlso moPlayerRels(X) Is Nothing = False Then
    '                moPlayerRels(X).RegenerateWPV()
    '            End If
    '        Next X
    '    End Sub
    '#End Region
    


    Public Function GetObjAsString(Optional ByVal bExcludeUserNameAndPassword As Boolean = False) As Byte()
        'here we will return the entire object as a string
        Dim lPos As Int32

        'No longer cache this...
        'If mbStringReady = False Then
        If bExcludeUserNameAndPassword = True Then
            ReDim mySendString(186) '178
        Else
            ReDim mySendString(165) '161
        End If

        GetGUIDAsString.CopyTo(mySendString, 0)
        PlayerName.CopyTo(mySendString, 6)
        EmpireName.CopyTo(mySendString, 26)
        RaceName.CopyTo(mySendString, 46)
        lPos = 66

        If bExcludeUserNameAndPassword = False Then
            PlayerUserName.CopyTo(mySendString, lPos)
            lPos += 20
            PlayerPassword.CopyTo(mySendString, lPos)
            lPos += 20
        End If

        'If oSenate Is Nothing = False Then
        '	System.BitConverter.GetBytes(oSenate.ObjectID).CopyTo(mySendString, lPos)
        'Else 
        'System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos)
        'End If
        'lPos += 4
        mySendString(lPos) = yGender : lPos += 1

        lPos += 3   'unused 3 bytes here

        System.BitConverter.GetBytes(CommEncryptLevel).CopyTo(mySendString, lPos)
        lPos += 2
        mySendString(lPos) = EmpireTaxRate
        lPos += 1

        System.BitConverter.GetBytes(blCredits).CopyTo(mySendString, lPos)
        lPos += 8

        If bExcludeUserNameAndPassword = False Then
            System.BitConverter.GetBytes(lStartedEnvirID).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lStartLocX).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lStartLocZ).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(PirateStartLocX).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(PirateStartLocZ).CopyTo(mySendString, lPos) : lPos += 4
        Else
            mySendString(lPos) = yPlayerTitle : lPos += 1
            System.BitConverter.GetBytes(lPlayerIcon).CopyTo(mySendString, lPos) : lPos += 4

            With Me
                'If .lLastGlobalRequestTurnIn = 0 OrElse glCurrentCycle - .lLastGlobalRequestTurnIn > 9000 Then
                '    goMsgSys.SendRequestGlobalPlayerScores(.ObjectID, 64, .ObjectID)
                '    If .lLastGlobalRequestTurnIn = 0 Then
                .lLGTechScore = .TechnologyScore
                .lLGDiplomacyScore = .DiplomacyScore
                .lLGWealthScore = .WealthScore
                .lLGProductionScore = .ProductionScore
                .lLGPopulationScore = .PopulationScore
                .lLGMilitaryScore = .lMilitaryScore
                .lLGTotalScore = .TotalScore
                '    End If
                'End If
            End With

            System.BitConverter.GetBytes(lLGTechScore).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lLGDiplomacyScore).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(CInt(lLGMilitaryScore \ 50)).CopyTo(mySendString, lPos) : lPos += 4      'was 100
            System.BitConverter.GetBytes(lLGPopulationScore).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lLGProductionScore).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lLGWealthScore).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(lLGTotalScore).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(DeathBudgetBalance).CopyTo(mySendString, lPos) : lPos += 4
        End If
        System.BitConverter.GetBytes(lGuildID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lJoinedGuildOn).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(BadWarDecCPIncrease).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(BadWarDecMoralePenalty).CopyTo(mySendString, lPos) : lPos += 4

        mySendString(lPos) = Me.yPlayerPhase : lPos += 1
        System.BitConverter.GetBytes(lTutorialStep).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(AccountStatus).CopyTo(mySendString, lPos) : lPos += 4

        If bExcludeUserNameAndPassword = True Then
            System.BitConverter.GetBytes(lStatusFlags).CopyTo(mySendString, lPos) : lPos += 4
        End If

        Return mySendString
    End Function

    Public Function GetSaveObjectText(ByRef oAllDataSaver As AllDataSaver) As Boolean
        Dim bResult As Boolean = False

        If ObjectID = -1 Then
            Return SaveObject(True)
        End If

        Try
            Dim sSQL As String

            'UPDATE
            sSQL = "UPDATE tblPlayer SET PlayerName = '" & MakeDBStr(BytesToString(PlayerName)) & "', EmpireName = '" & MakeDBStr(BytesToString(EmpireName)) & _
              "', RaceName = '" & MakeDBStr(BytesToString(RaceName)) & "', PlayerUserName = '" & MakeDBStr(BytesToString(PlayerUserName)) & "', PlayerPassword = '" & _
              MakeDBStr(BytesToString(PlayerPassword)) & "', SenateID = "
            'If oSenate Is Nothing = False Then
            '	sSQL = sSQL & oSenate.ObjectID
            'Else
            sSQL = sSQL & "-1"
            'End If
            sSQL = sSQL & ", CommEncryptLevel = " & CommEncryptLevel & ", EmpireTaxRate = " & EmpireTaxRate & _
              ", Credits = " & blCredits & ", LastViewedID = " & lLastViewedEnvir & ", LastViewedTypeID = " & iLastViewedEnvirType & _
              ", BaseMorale = " & BaseMorale & ", StartedEnvirID = " & lStartedEnvirID & ", StartedEnvirTypeID = " & iStartedEnvirTypeID & _
              ", TotalPlayTime = " & TotalPlayTime & ", MaxTotalPop = " & blMaxPopulation & ", LastLogin = "
            If LastLogin <> Date.MinValue Then
                Dim lVal As Int32 = CInt(Val(LastLogin.ToString("yyMMddHHmm")))
                sSQL &= lVal.ToString
            Else : sSQL &= "0"
            End If
            Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
            If lTemp < 0 Then lTemp = -1
            Dim lDBValue As Int32 = -1
            Dim lDBFunds As Int32 = Me.DeathBudgetFundsRemaining
            If DeathBudgetEndTime <> Int32.MinValue AndAlso DeathBudgetEndTime > glCurrentCycle Then lDBValue = DeathBudgetEndTime - glCurrentCycle
            sSQL &= ", WarSentiment = " & lWarSentiment & ", CelebrationEnds = " & lTemp & ", StartLocX = " & lStartLocX & ", StartLocZ = " & lStartLocZ
            sSQL &= ", PirateStartLocX = " & PirateStartLocX & ", PirateStartLocZ = " & PirateStartLocZ & _
              ", GuildID = " & lGuildID & ", GuildRankID = " & lGuildRankID & ", JoinedGuildOn = " & lJoinedGuildOn & _
              ", PlayerTitle = " & yPlayerTitle & ", PlayerIcon = " & lPlayerIcon & ", PlayerGender = " & yGender & _
              ", DeathBudgetBalance = " & DeathBudgetBalance & ", AccountStatus = " & Me.AccountStatus & ", EmailAddress = '" & _
              MakeDBStr(BytesToString(ExternalEmailAddress)) & "', EmailSettings = " & iEmailSettings & _
              ", DeathBudgetCycles = " & lDBValue & ", DeathBudgetFunds = " & lDBFunds & ", BadWarDecCPIncrease = "
            If Me.BadWarDecCPIncreaseEndCycle - glCurrentCycle < 1 Then
                sSQL &= "0, BadWarDecCPIncreaseEndCycle = 0,"
            Else
                sSQL &= Me.BadWarDecCPIncrease & ", BadWarDecCPIncreaseEndCycle = " & (Me.BadWarDecCPIncreaseEndCycle - glCurrentCycle) & ", "
            End If
            If Me.BadWarDecMoralePenaltyEndCycle - glCurrentCycle < 1 Then
                sSQL &= "BadWarDecMoralePenalty = 0, BadWarDecMoralePenaltyEndCycle = 0 "
            Else
                sSQL &= "BadWarDecMoralePenalty = " & Me.BadWarDecMoralePenalty & ", BadWarDecMoralePenaltyEndCycle = " & (Me.BadWarDecMoralePenaltyEndCycle - glCurrentCycle) & " "
            End If
            For X As Int32 = 0 To 4
                sSQL &= ", Slot" & (X + 1) & "ID = " & lSlotID(X) & ", Slot" & (X + 1) & "State = " & ySlotState(X)
            Next X
            For X As Int32 = 0 To 2
                sSQL &= ", Faction" & (X + 1) & "ID = " & lFactionID(X)
            Next X
            sSQL &= ", IronCurtainPlanetID = " & lIronCurtainPlanet & ", PlayerPhase = " & CInt(Me.yPlayerPhase).ToString & _
              ", TutorialStep = " & Me.lTutorialStep.ToString & _
              ", PlayedTimeWhenFirstWave = " & PlayedTimeWhenFirstWave & ", PlayedTimeAtEndOfWaves = " & PlayedTimeAtEndOfWaves & _
              ", PlayedTimeInTutorialOne = " & PlayedTimeInTutorialOne & ", TutorialPhaseWaves = " & TutorialPhaseWaves & _
              ", InternalEmailSettings = " & iInternalEmailSettings & ", SpecTechCostMult = " & SpecTechCostMult.ToString & _
              ", GuaranteedSpecialTechID = " & GuaranteedSpecialTechID & ", SpawnSystemSetting = " & ySpawnSystemSetting & " WHERE PlayerID = " & ObjectID
            '", GuaranteedSpecialTechID = " & GuaranteedSpecialTechID & ", LastWPUpkeep = " & lLastWPUpkeep.ToString & _
            '", LastGuildShareUpkeep = " & lLastGuildShareUpkeep.ToString & ", Warpoints = " & blWarpoints.ToString & _
            '", WarpointsAllTime = " & blWarpointsAllTime.ToString & ", LastWPUpkeepTime = " & lLastWPUpkeepTime.ToString & " WHERE PlayerID = " & ObjectID

            If oAllDataSaver.AddCommandText(sSQL) = False Then Return False

            Dim lUB As Int32 = -1
            If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportIdx.GetUpperBound(0), mlTransportUB)
            For X As Int32 = 0 To lUB
                If mlTransportIdx(X) > -1 Then
                    Dim oTrans As Transport = moTransports(X)
                    If oTrans Is Nothing = False Then
                        sSQL = oTrans.GetSaveObjectText()
                        If sSQL <> "" Then
                            If oAllDataSaver.AddCommandText(sSQL) = False Then Return False
                        End If
                    End If
                End If
            Next X

            'For X As Int32 = 0 To Me.PlayerRelUB
            '    If mlPlayerRelIdx(X) <> -1 Then
            '        oAllDataSaver.AddCommandText(moPlayerRels(X).GetSaveObjectText())
            '    End If
            'Next X

            ''Save our Alias'
            'For X As Int32 = 0 To Me.lAliasUB
            '    Dim oComm As OleDb.OleDbCommand = Nothing
            '    Try
            '        If Me.lAliasIdx(X) <> -1 Then
            '            sSQL = "UPDATE tblAlias SET sAliasUserName = '" & MakeDBStr(BytesToString(uAliasLogin(X).yUserName)) & _
            '               "', sAliasPassword = '" & MakeDBStr(BytesToString(uAliasLogin(X).yPassword)) & "', lAliasRights = " & _
            '               uAliasLogin(X).lRights & " WHERE PlayerID = " & Me.ObjectID & " AND OtherPlayerID = " & Me.oAliases(X).ObjectID
            '            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '            If oComm.ExecuteNonQuery() = 0 Then
            '                oComm.Dispose()
            '                oComm = Nothing
            '                sSQL = "INSERT INTO tblAlias (PlayerID, OtherPlayerID, sAliasUserName, sAliasPassword, lAliasRights) VALUES (" & _
            '                 Me.ObjectID & ", " & Me.oAliases(X).ObjectID & ", '" & MakeDBStr(BytesToString(uAliasLogin(X).yUserName)) & _
            '                 "', '" & MakeDBStr(BytesToString(uAliasLogin(X).yPassword)) & "', " & uAliasLogin(X).lRights & ")"
            '                oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '                If oComm.ExecuteNonQuery() = 0 Then
            '                    Err.Raise(-1, "", "No records affected with Update/Insert!")
            '                End If
            '            End If
            '        End If
            '    Catch ex As Exception
            '        LogEvent(LogEventType.CriticalError, "Player.SaveObject.SaveAlias: " & ex.Message)
            '    Finally
            '        If oComm Is Nothing = False Then oComm.Dispose()
            '        oComm = Nothing
            '    End Try
            'Next X



            'sSQL = "DELETE FROM tblPlayerMineralProperty WHERE PlayerID = " & Me.ObjectID
            'oAllDataSaver.AddCommandText(sSQL)

            ''Save known properties
            'For X As Int32 = 0 To mlMinPropertyUB
            '    oAllDataSaver.AddCommandText(moMinProperties(X).GetSaveObjectText())
            'Next X

            ''Save player minerals
            'For X As Int32 = 0 To lPlayerMineralUB
            '    oAllDataSaver.AddCommandText(oPlayerMinerals(X).GetSaveObjectText(Me.ObjectID))
            'Next X

            'For X As Int32 = 0 To mlAlloyUB
            '    If mlAlloyIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moAlloy(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlArmorUB
            '    If mlArmorIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moArmor(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlEngineUB
            '    If mlEngineIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moEngine(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlHullUB
            '    If mlHullIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moHull(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlPrototypeUB
            '    If mlPrototypeIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moPrototype(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlRadarUB
            '    If mlRadarIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moRadar(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlShieldUB
            '    If mlShieldIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moShield(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlWeaponUB
            '    If mlWeaponIdx(X) <> -1 Then oAllDataSaver.AddCommandText(moWeapon(X).GetSaveObjectText())
            'Next X
            'For X As Int32 = 0 To mlSpecialTechUB
            '    If mlSpecialTechIdx(X) <> -1 Then moSpecialTech(X).SaveObject()
            'Next X

            ''Save our emails...
            'For X As Int32 = 0 To EmailFolderUB
            '    If EmailFolderIdx(X) <> -1 Then
            '        EmailFolders(X).PlayerID = Me.ObjectID
            '        EmailFolders(X).SaveFolder()
            '    End If
            'Next X

            ''Save our Player Intels
            'For X As Int32 = 0 To mlPlayerIntelUB
            '    If mlPlayerIntelIdx(X) <> -1 Then
            '        moPlayerIntel(X).SaveObject()
            '    End If
            'Next X

            ''our tech knowledges
            'For X As Int32 = 0 To mlPlayerTechKnowledgeUB
            '    If myPlayerTechKnowledgeUsed(X) <> 0 Then
            '        moPlayerTechKnowledge(X).SaveObject()
            '    End If
            'Next X

            'For X As Int32 = 0 To mlItemIntelUB
            '    If moItemIntel(X) Is Nothing = False Then moItemIntel(X).SaveObject()
            'Next X

            'sSQL = "DELETE FROM tblPlayerWormhole WHERE PlayerID = " & Me.ObjectID
            'oAllDataSaver.AddCommandText(sSQL)

            'For X As Int32 = 0 To lWormholeUB
            '    If oWormholes(X) Is Nothing = False Then
            '        sSQL = "INSERT INTO tblPlayerWormhole (PlayerID, WormholeID) VALUES (" & Me.ObjectID & ", " & oWormholes(X).ObjectID & ")"
            '        oAllDataSaver.AddCommandText(sSQL)
            '    End If
            'Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
            bResult = False
        End Try
        Return bResult
    End Function

    Public Function SaveObject(ByVal bPlayerObjectOnly As Boolean) As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!


        'Public PlayedTimeWhenTimerStarted As Int32 = Int32.MinValue			'for the tutorial phase 2, when the four hour timer started, what was the player's play time?
        'Public PlayedTimeWhenFirstWave As Int32 = Int32.MinValue
        'Public PlayedTimeAtEndOfWaves As Int32 = Int32.MinValue
        'Public PlayedTimeInTutorialOne As Int32 = Int32.MinValue
        'Public TutorialPhaseWaves As Int32 = 0

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblPlayer (PlayerName, EmpireName, RaceName, PlayerUserName, PlayerPassword, SenateID, CommEncryptLevel, " & _
                  "EmpireTaxRate, Credits, LastViewedID, LastViewedTypeID, BaseMorale, StartedEnvirID, StartedEnvirTypeID, TotalPlayTime, LastLogin, " & _
                  "WarSentiment, CelebrationEnds, StartLocX, StartLocZ, PirateStartLocX, PirateStartLocZ, GuildID, GuildRankID, MaxTotalPop, " & _
                  "JoinedGuildOn, PlayerTitle, PlayerIcon, PlayerGender, DeathBudgetBalance, DeathBudgetCycles, DeathBudgetFunds, AccountStatus, " & _
                  "EmailAddress, EmailSettings, BadWarDecCPIncrease, BadWarDecCPIncreaseEndCycle, BadWarDecMoralePenalty, BadWarDecMoralePenaltyEndCycle, " & _
                  "Slot1ID, Slot1State, Slot2ID, Slot2State, Slot3ID, Slot3State, Slot4ID, Slot4State, Slot5ID, Slot5State, Faction1ID, " & _
                  "Faction2ID, Faction3ID, IronCurtainPlanetID, PlayerPhase, TutorialStep, PlayedTimeWhenFirstWave, " & _
                  "PlayedTimeAtEndOfWaves, PlayedTimeInTutorialOne, TutorialPhaseWaves, InternalEmailSettings, ExTitleNew, ExTitleEnd, CustomTitlePermission, " & _
                  "CustomTitle, SpecTechCostMult, GuaranteedSpecialTechID, StatusFlags, DBTotalPop, LastGuildMembership, SpawnSystemSetting) VALUES ('" & _
                  MakeDBStr(BytesToString(PlayerName)) & "', '" & MakeDBStr(BytesToString(EmpireName)) & "', '" & MakeDBStr(BytesToString(RaceName)) & "', '" & _
                  MakeDBStr(BytesToString(PlayerUserName)) & "', '" & MakeDBStr(BytesToString(PlayerPassword)) & "', "
                '"CustomTitle, SpecTechCostMult, GuaranteedSpecialTechID, StatusFlags, DBTotalPop, LastGuildMembership, LastWPUpkeep, LastGuildShareUpkeep, Warpoints, WarpointsAllTime, LastWPUpkeepTime) VALUES ('" & _
                'MakeDBStr(BytesToString(PlayerName)) & "', '" & MakeDBStr(BytesToString(EmpireName)) & "', '" & MakeDBStr(BytesToString(RaceName)) & "', '" & _
                'MakeDBStr(BytesToString(PlayerUserName)) & "', '" & MakeDBStr(BytesToString(PlayerPassword)) & "', "

                'If oSenate Is Nothing = False Then
                '	sSQL = sSQL & oSenate.ObjectID & ", "
                'Else
                sSQL = sSQL & "-1, "
                'End If
                sSQL = sSQL & CommEncryptLevel & ", " & EmpireTaxRate & ", " & blCredits & ", " & lLastViewedEnvir & ", " & _
                  iLastViewedEnvirType & ", " & BaseMorale & ", " & lStartedEnvirID & ", " & iStartedEnvirTypeID & ", " & _
                  TotalPlayTime & ", "
                If LastLogin <> Date.MinValue Then
                    Dim lVal As Int32 = CInt(Val(LastLogin.ToString("yyMMddHHmm")))
                    sSQL &= lVal.ToString
                Else : sSQL &= "0"
                End If
                Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                If lTemp < 0 Then lTemp = -1
                Dim lDBValue As Int32 = -1
                Dim lDBFunds As Int32 = Me.DeathBudgetFundsRemaining
                If DeathBudgetEndTime <> Int32.MinValue AndAlso DeathBudgetEndTime > glCurrentCycle Then lDBValue = DeathBudgetEndTime - glCurrentCycle
                sSQL &= ", " & lWarSentiment & ", " & lTemp & ", " & lStartLocX & ", " & lStartLocZ & ", " & _
                  PirateStartLocX & ", " & PirateStartLocZ & ", " & lGuildID & ", " & lGuildRankID & ", " & blMaxPopulation & ", " & _
                  lJoinedGuildOn & ", " & yPlayerTitle & ", " & lPlayerIcon & ", " & yGender & ", " & DeathBudgetBalance & _
                  ", " & lDBValue & ", " & lDBFunds & ", " & Me.AccountStatus & ", '" & MakeDBStr(BytesToString(ExternalEmailAddress)) & _
                  "', " & iEmailSettings & ", "
                If glCurrentCycle - Me.BadWarDecCPIncreaseEndCycle < 1 Then sSQL &= "0, 0, " Else sSQL &= Me.BadWarDecCPIncrease & ", " & (glCurrentCycle - Me.BadWarDecCPIncreaseEndCycle) & ", "
                If glCurrentCycle - Me.BadWarDecMoralePenaltyEndCycle < 1 Then sSQL &= "0, 0" Else sSQL &= Me.BadWarDecMoralePenalty & ", " & (glCurrentCycle - Me.BadWarDecMoralePenaltyEndCycle)
                For X As Int32 = 0 To 4
                    sSQL &= ", " & lSlotID(X) & ", " & ySlotState(X)
                Next X
                For X As Int32 = 0 To 2
                    sSQL &= ", " & lFactionID(X)
                Next X
                sSQL &= ", " & lIronCurtainPlanet & ", " & CInt(Me.yPlayerPhase).ToString & ", " & Me.lTutorialStep.ToString & _
                  ", " & PlayedTimeWhenFirstWave & ", " & PlayedTimeAtEndOfWaves & ", " & _
                  PlayedTimeInTutorialOne & ", " & TutorialPhaseWaves & ", " & iInternalEmailSettings.ToString & ", " & yExTitleNew.ToString & _
                  ", " & GetDateAsNumber(dtExTitleEnd).ToString & ", " & lCustomTitlePermission.ToString & ", " & yCustomTitle & ", " & _
                  SpecTechCostMult.ToString() & ", " & GuaranteedSpecialTechID & ", " & lStatusFlags & ", " & blDBPopulation & ", " '& ")"
                If dtLastGuildMembership <> Date.MinValue Then
                    Dim lVal As Int32 = CInt(Val(dtLastGuildMembership.ToString("yyMMddHHmm")))
                    sSQL &= lVal.ToString
                Else
                    sSQL &= "0"
                End If
                'sSQL &= ", " & lLastWPUpkeep.ToString & ", " & lLastGuildShareUpkeep.ToString & ", " & blWarpoints.ToString
                'sSQL &= ", " & blWarpointsAllTime.ToString & ", " & lLastWPUpkeepTime.ToString & ")"
                sSQL &= ", " & ySpawnSystemSetting & ")"

            Else
                'UPDATE
                sSQL = "UPDATE tblPlayer SET PlayerName = '" & MakeDBStr(BytesToString(PlayerName)) & "', EmpireName = '" & MakeDBStr(BytesToString(EmpireName)) & _
                  "', RaceName = '" & MakeDBStr(BytesToString(RaceName)) & "', PlayerUserName = '" & MakeDBStr(BytesToString(PlayerUserName)) & "', PlayerPassword = '" & _
                  MakeDBStr(BytesToString(PlayerPassword)) & "', SenateID = "
                'If oSenate Is Nothing = False Then
                '	sSQL = sSQL & oSenate.ObjectID
                'Else
                sSQL = sSQL & "-1"
                'End If
                sSQL = sSQL & ", CommEncryptLevel = " & CommEncryptLevel & ", EmpireTaxRate = " & EmpireTaxRate & _
                  ", Credits = " & blCredits & ", LastViewedID = " & lLastViewedEnvir & ", LastViewedTypeID = " & iLastViewedEnvirType & _
                  ", BaseMorale = " & BaseMorale & ", StartedEnvirID = " & lStartedEnvirID & ", StartedEnvirTypeID = " & iStartedEnvirTypeID & _
                  ", TotalPlayTime = " & TotalPlayTime & ", MaxTotalPop = " & blMaxPopulation & ", LastLogin = "
                If LastLogin <> Date.MinValue Then
                    Dim lVal As Int32 = CInt(Val(LastLogin.ToString("yyMMddHHmm")))
                    sSQL &= lVal.ToString
                Else : sSQL &= "0"
                End If
                Dim lTemp As Int32 = lCelebrationEnds - glCurrentCycle
                If lTemp < 0 Then lTemp = -1
                Dim lDBValue As Int32 = -1
                Dim lDBFunds As Int32 = Me.DeathBudgetFundsRemaining
                If DeathBudgetEndTime <> Int32.MinValue AndAlso DeathBudgetEndTime > glCurrentCycle Then lDBValue = DeathBudgetEndTime - glCurrentCycle
                sSQL &= ", WarSentiment = " & lWarSentiment & ", CelebrationEnds = " & lTemp & ", StartLocX = " & lStartLocX & ", StartLocZ = " & lStartLocZ
                sSQL &= ", PirateStartLocX = " & PirateStartLocX & ", PirateStartLocZ = " & PirateStartLocZ & _
                  ", GuildID = " & lGuildID & ", GuildRankID = " & lGuildRankID & ", JoinedGuildOn = " & lJoinedGuildOn & _
                  ", PlayerTitle = " & yPlayerTitle & ", PlayerIcon = " & lPlayerIcon & ", PlayerGender = " & yGender & _
                  ", DeathBudgetBalance = " & DeathBudgetBalance & ", AccountStatus = " & Me.AccountStatus & ", EmailAddress = '" & _
                  MakeDBStr(BytesToString(ExternalEmailAddress)) & "', EmailSettings = " & iEmailSettings & _
                  ", DeathBudgetCycles = " & lDBValue & ", DeathBudgetFunds = " & lDBFunds & ", BadWarDecCPIncrease = "
                If Me.BadWarDecCPIncreaseEndCycle - glCurrentCycle < 1 Then
                    sSQL &= "0, BadWarDecCPIncreaseEndCycle = 0,"
                Else
                    sSQL &= Me.BadWarDecCPIncrease & ", BadWarDecCPIncreaseEndCycle = " & (Me.BadWarDecCPIncreaseEndCycle - glCurrentCycle) & ", "
                End If
                If Me.BadWarDecMoralePenaltyEndCycle - glCurrentCycle < 1 Then
                    sSQL &= "BadWarDecMoralePenalty = 0, BadWarDecMoralePenaltyEndCycle = 0 "
                Else
                    sSQL &= "BadWarDecMoralePenalty = " & Me.BadWarDecMoralePenalty & ", BadWarDecMoralePenaltyEndCycle = " & (Me.BadWarDecMoralePenaltyEndCycle - glCurrentCycle) & " "
                End If
                For X As Int32 = 0 To 4
                    sSQL &= ", Slot" & (X + 1) & "ID = " & lSlotID(X) & ", Slot" & (X + 1) & "State = " & ySlotState(X)
                Next X
                For X As Int32 = 0 To 2
                    sSQL &= ", Faction" & (X + 1) & "ID = " & lFactionID(X)
                Next X
                sSQL &= ", IronCurtainPlanetID = " & lIronCurtainPlanet & ", PlayerPhase = " & CInt(Me.yPlayerPhase).ToString & _
                  ", TutorialStep = " & Me.lTutorialStep.ToString & _
                  ", PlayedTimeWhenFirstWave = " & PlayedTimeWhenFirstWave & ", PlayedTimeAtEndOfWaves = " & PlayedTimeAtEndOfWaves & _
                  ", PlayedTimeInTutorialOne = " & PlayedTimeInTutorialOne & ", TutorialPhaseWaves = " & TutorialPhaseWaves & _
                  ", InternalEmailSettings = " & iInternalEmailSettings & ", ExTitleNew = " & yExTitleNew & ", ExTitleEnd = " & _
                  GetDateAsNumber(dtExTitleEnd).ToString & ", CustomTitlePermission = " & lCustomTitlePermission.ToString & ", CustomTitle = " & _
                  yCustomTitle & ", SpecTechCostMult = " & SpecTechCostMult.ToString & ", GuaranteedSpecialTechID = " & GuaranteedSpecialTechID & _
                  ", StatusFlags = " & lStatusFlags & ", DBTotalPop = " & blDBPopulation & ", LastGuildMembership = "
                If dtLastGuildMembership <> Date.MinValue Then
                    Dim lVal As Int32 = CInt(Val(dtLastGuildMembership.ToString("yyMMddHHmm")))
                    sSQL &= lVal.ToString
                Else
                    sSQL &= "0"
                End If
                'sSQL &= ", LastWPUpkeep = " & lLastWPUpkeep.ToString & ", LastGuildShareUpkeep = " & lLastGuildShareUpkeep.ToString & ", Warpoints = " & blWarpoints.ToString
                'sSQL &= ", WarpointsAllTime = " & blWarpointsAllTime.ToString & ", LastWPUpkeepTime = " & lLastWPUpkeepTime.ToString & " WHERE PlayerID = " & ObjectID
                sSQL &= ", SpawnSystemSetting = " & ySpawnSystemSetting & " WHERE PlayerID = " & ObjectID

            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PlayerID) FROM tblPlayer WHERE PlayerUserName = '" & MakeDBStr(BytesToString(PlayerUserName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            For X As Int32 = 0 To Me.PlayerRelUB
                If mlPlayerRelIdx(X) <> -1 AndAlso (moPlayerRels(X).bForceSaveMe = True OrElse (moPlayerRels(X).lCycleOfNextScoreUpdate > -1 AndAlso moPlayerRels(X).TargetScore <> moPlayerRels(X).WithThisScore)) Then
                    If moPlayerRels(X).SaveObject = False Then
                        Err.Raise(-1, "Player.PlayerRels.SaveObject()", "Unable to save Player Rel!")
                    End If
                End If
            Next X

            If bPlayerObjectOnly = False Then
                Dim lUB As Int32 = -1
                If mlTransportIdx Is Nothing = False Then lUB = Math.Min(mlTransportIdx.GetUpperBound(0), mlTransportUB)
                For X As Int32 = 0 To lUB
                    If mlTransportIdx(X) > -1 Then
                        Dim oTrans As Transport = moTransports(X)
                        If oTrans Is Nothing = False Then oTrans.SaveObject()
                    End If
                Next X
            End If

            ''Save our Alias'
            'For X As Int32 = 0 To Me.lAliasUB
            '    Try
            '        If Me.lAliasIdx(X) <> -1 Then
            '            sSQL = "UPDATE tblAlias SET sAliasUserName = '" & MakeDBStr(BytesToString(uAliasLogin(X).yUserName)) & _
            '               "', sAliasPassword = '" & MakeDBStr(BytesToString(uAliasLogin(X).yPassword)) & "', lAliasRights = " & _
            '               uAliasLogin(X).lRights & " WHERE PlayerID = " & Me.ObjectID & " AND OtherPlayerID = " & Me.oAliases(X).ObjectID
            '            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '            If oComm.ExecuteNonQuery() = 0 Then
            '                oComm.Dispose()
            '                oComm = Nothing
            '                sSQL = "INSERT INTO tblAlias (PlayerID, OtherPlayerID, sAliasUserName, sAliasPassword, lAliasRights) VALUES (" & _
            '                 Me.ObjectID & ", " & Me.oAliases(X).ObjectID & ", '" & MakeDBStr(BytesToString(uAliasLogin(X).yUserName)) & _
            '                 "', '" & MakeDBStr(BytesToString(uAliasLogin(X).yPassword)) & "', " & uAliasLogin(X).lRights & ")"
            '                oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '                If oComm.ExecuteNonQuery() = 0 Then
            '                    Err.Raise(-1, "", "No records affected with Update/Insert!")
            '                End If
            '            End If
            '        End If
            '    Catch ex As Exception
            '        LogEvent(LogEventType.CriticalError, "Player.SaveObject.SaveAlias: " & ex.Message)
            '    Finally
            '        If oComm Is Nothing = False Then oComm.Dispose()
            '        oComm = Nothing
            '    End Try
            'Next X

            'If bPlayerObjectOnly = False Then

            '    sSQL = "DELETE FROM tblPlayerMineralProperty WHERE PlayerID = " & Me.ObjectID
            '    oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '    oComm.ExecuteNonQuery()
            '    oComm.Dispose()
            '    oComm = Nothing

            '    'Save known properties
            '    For X As Int32 = 0 To mlMinPropertyUB
            '        moMinProperties(X).SaveObject()
            '    Next X

            '    'Save player minerals
            '    For X As Int32 = 0 To lPlayerMineralUB
            '        oPlayerMinerals(X).SaveObject(Me.ObjectID)
            '    Next X

            '    'For X As Int32 = 0 To mlAlloyUB
            '    '    If mlAlloyIdx(X) <> -1 Then moAlloy(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlArmorUB
            '    '    If mlArmorIdx(X) <> -1 Then moArmor(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlEngineUB
            '    '    If mlEngineIdx(X) <> -1 Then moEngine(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlHullUB
            '    '    If mlHullIdx(X) <> -1 Then moHull(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlPrototypeUB
            '    '    If mlPrototypeIdx(X) <> -1 Then moPrototype(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlRadarUB
            '    '    If mlRadarIdx(X) <> -1 Then moRadar(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlShieldUB
            '    '    If mlShieldIdx(X) <> -1 Then moShield(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlWeaponUB
            '    '    If mlWeaponIdx(X) <> -1 Then moWeapon(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlSpecialTechUB
            '    '    If mlSpecialTechIdx(X) <> -1 Then moSpecialTech(X).SaveObject()
            '    'Next

            '    'Save our emails...
            '    For X As Int32 = 0 To EmailFolderUB
            '        If EmailFolderIdx(X) <> -1 Then
            '            EmailFolders(X).PlayerID = Me.ObjectID
            '            EmailFolders(X).SaveFolder()
            '        End If
            '    Next X

            '    'Save our Player Intels
            '    For X As Int32 = 0 To mlPlayerIntelUB
            '        If mlPlayerIntelIdx(X) <> -1 Then
            '            moPlayerIntel(X).SaveObject()
            '        End If
            '    Next X

            '    'our tech knowledges
            '    For X As Int32 = 0 To mlPlayerTechKnowledgeUB
            '        If myPlayerTechKnowledgeUsed(X) <> 0 Then
            '            moPlayerTechKnowledge(X).SaveObject()
            '        End If
            '    Next X

            '    For X As Int32 = 0 To mlItemIntelUB
            '        If moItemIntel(X) Is Nothing = False Then moItemIntel(X).SaveObject()
            '    Next X

            '    sSQL = "DELETE FROM tblPlayerWormhole WHERE PlayerID = " & Me.ObjectID
            '    oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '    oComm.ExecuteNonQuery()
            '    oComm.Dispose()
            '    oComm = Nothing

            '    For X As Int32 = 0 To lWormholeUB
            '        If oWormholes(X) Is Nothing = False Then
            '            sSQL = "INSERT INTO tblPlayerWormhole (PlayerID, WormholeID) VALUES (" & Me.ObjectID & ", " & oWormholes(X).ObjectID & ")"
            '            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '            If oComm.ExecuteNonQuery() = 0 Then
            '                LogEvent(LogEventType.CriticalError, "Unable to save Player Wormhole knowledge: No records affected!")
            '            End If
            '            oComm.Dispose()
            '            oComm = Nothing
            '        End If
            '    Next X
            'Else
            '    'For X As Int32 = 0 To mlAlloyUB
            '    '    If mlAlloyIdx(X) <> -1 AndAlso moAlloy(X).bNeedsSaved = True Then moAlloy(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlArmorUB
            '    '    If mlArmorIdx(X) <> -1 AndAlso moArmor(X).bNeedsSaved = True Then moArmor(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlEngineUB
            '    '    If mlEngineIdx(X) <> -1 AndAlso moEngine(X).bNeedsSaved = True Then moEngine(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlHullUB
            '    '    If mlHullIdx(X) <> -1 AndAlso moHull(X).bNeedsSaved = True Then moHull(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlPrototypeUB
            '    '    If mlPrototypeIdx(X) <> -1 AndAlso moPrototype(X).bNeedsSaved = True Then moPrototype(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlRadarUB
            '    '    If mlRadarIdx(X) <> -1 AndAlso moRadar(X).bNeedsSaved = True Then moRadar(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlShieldUB
            '    '    If mlShieldIdx(X) <> -1 AndAlso moShield(X).bNeedsSaved = True Then moShield(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlWeaponUB
            '    '    If mlWeaponIdx(X) <> -1 AndAlso moWeapon(X).bNeedsSaved = True Then moWeapon(X).SaveObject()
            '    'Next X
            '    'For X As Int32 = 0 To mlSpecialTechUB
            '    '    If mlSpecialTechIdx(X) <> -1 AndAlso moSpecialTech(X).bNeedsSaved = True Then moSpecialTech(X).SaveObject()
            '    'Next X

            '    If lWormholeUB <> -1 Then
            '        Dim bWHSaved(lWormholeUB) As Boolean
            '        For X As Int32 = 0 To lWormholeUB
            '            bWHSaved(X) = False
            '        Next X
            '        sSQL = "SELECT * FROM tblPlayerWormhole WHERE PlayerID = " & Me.ObjectID
            '        oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '        Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()
            '        While oData.Read
            '            Dim lWormholeID As Int32 = CInt(oData("WormholeID"))
            '            For X As Int32 = 0 To lWormholeUB
            '                If oWormholes(X).ObjectID = lWormholeID Then
            '                    bWHSaved(X) = True
            '                    Exit For
            '                End If
            '            Next X
            '        End While
            '        oData.Close()
            '        oData = Nothing
            '        oComm.Dispose()
            '        oComm = Nothing

            '        For X As Int32 = 0 To lWormholeUB
            '            If oWormholes(X) Is Nothing = False AndAlso bWHSaved(X) = False Then
            '                Try
            '                    sSQL = "INSERT INTO tblPlayerWormhole (PlayerID, WormholeID) VALUES (" & Me.ObjectID & ", " & oWormholes(X).ObjectID & ")"
            '                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
            '                    If oComm.ExecuteNonQuery() = 0 Then
            '                        Err.Raise(-1, "Player.SaveObject (async)", "No records affected!")
            '                    End If
            '                Catch ex As Exception
            '                    LogEvent(LogEventType.CriticalError, "Player.SaveObject: " & ex.Message)
            '                Finally
            '                    oComm.Dispose()
            '                    oComm = Nothing
            '                End Try
            '            End If
            '        Next X
            '    End If
            'End If

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function
    Private Function SavePlayerWormhole(ByVal lIdx As Int32) As Boolean
        Dim bResult As Boolean = False

        If lIdx > -1 AndAlso Me.lWormholeUB >= lIdx Then
            Dim oComm As OleDb.OleDbCommand = Nothing
            Try
                Dim sSQL As String = "INSERT INTO tblPlayerWormhole (PlayerID, WormholeID) VALUES (" & Me.ObjectID & " , " & oWormholes(lIdx).ObjectID & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "Player.SavePlayerWormhole", "No records affected!")
                End If
                bResult = True
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Player.SavePlayerWormhole: " & ex.Message)
            Finally
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End Try
        End If

        Return bResult

    End Function
    Private Function SaveAliasConfig(ByVal lIdx As Int32) As Boolean
        'Save our Alias'
        Dim bResult As Boolean = False

        If lIdx > -1 AndAlso Me.lAliasUB >= lIdx Then
            Dim oComm As OleDb.OleDbCommand = Nothing
            Try
                If Me.lAliasIdx(lIdx) <> -1 Then
                    Dim sSQL As String = "UPDATE tblAlias SET sAliasUserName = '" & MakeDBStr(BytesToString(uAliasLogin(lIdx).yUserName)) & _
                       "', sAliasPassword = '" & MakeDBStr(BytesToString(uAliasLogin(lIdx).yPassword)) & "', lAliasRights = " & _
                       uAliasLogin(lIdx).lRights & " WHERE PlayerID = " & Me.ObjectID & " AND OtherPlayerID = " & Me.oAliases(lIdx).ObjectID
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        oComm.Dispose()
                        oComm = Nothing
                        sSQL = "INSERT INTO tblAlias (PlayerID, OtherPlayerID, sAliasUserName, sAliasPassword, lAliasRights) VALUES (" & _
                         Me.ObjectID & ", " & Me.oAliases(lIdx).ObjectID & ", '" & MakeDBStr(BytesToString(uAliasLogin(lIdx).yUserName)) & _
                         "', '" & MakeDBStr(BytesToString(uAliasLogin(lIdx).yPassword)) & "', " & uAliasLogin(lIdx).lRights & ")"
                        oComm = New OleDb.OleDbCommand(sSQL, goCN)
                        If oComm.ExecuteNonQuery() = 0 Then
                            Err.Raise(-1, "", "No records affected with Update/Insert!")
                        End If
                    End If
                    bResult = True
                End If
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Player.SaveAliasConfig: " & ex.Message)
            Finally
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End Try
        End If

        Return bResult
    End Function

    Public Sub New()
        If oBudget Is Nothing Then oBudget = New Budget
        oBudget.oPlayer = Me
        If oSpecials Is Nothing Then oSpecials = New Player_Specials
        oSpecials.oPlayer = Me
        If oSecurity Is Nothing Then oSecurity = New PlayerSecurity
        oSecurity.oParentPlayer = Me

        'add the default email boxes...
        AddFolder(PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF, "Deleted")
        AddFolder(PlayerCommFolder.ePCF_ID_HARDCODES.eDrafts_PCF, "Drafts")
        AddFolder(PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, "Inbox")
        AddFolder(PlayerCommFolder.ePCF_ID_HARDCODES.eOutbox_PCF, "Outbox")
    End Sub

    Public Sub CreateAndSendPlayerSpecialAttributeSetting(ByVal iSetting As ePlayerSpecialAttributeSetting, ByVal lValue As Int32)
        If Me.lConnectedPrimaryID > -1 Then Return
        Dim yOutMsg(11) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yOutMsg, 0)
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yOutMsg, 2)
        System.BitConverter.GetBytes(iSetting).CopyTo(yOutMsg, 6)
        System.BitConverter.GetBytes(lValue).CopyTo(yOutMsg, 8)
        'oSocket.SendData(yOutMsg)
        Me.CrossPrimarySafeSendMsg(yOutMsg)
    End Sub

    Public Sub HandleComponentListEspionage(ByVal bResearchedOnly As Boolean, ByRef oAttacker As Player)
        For X As Int32 = 0 To mlAlloyUB
            If mlAlloyIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moAlloy(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moAlloy(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moAlloy(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
        For X As Int32 = 0 To mlArmorUB
            If mlArmorIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moArmor(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moArmor(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moArmor(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
        For X As Int32 = 0 To mlEngineUB
            If mlEngineIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moEngine(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moEngine(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moEngine(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
        For X As Int32 = 0 To mlHullUB
            If mlHullIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moHull(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moHull(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moHull(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
        For X As Int32 = 0 To mlPrototypeUB
            If mlPrototypeIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moPrototype(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moPrototype(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moPrototype(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
        For X As Int32 = 0 To mlRadarUB
            If mlRadarIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moRadar(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moRadar(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moRadar(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
        For X As Int32 = 0 To mlShieldUB
            If mlShieldIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moShield(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moShield(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moShield(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
        For X As Int32 = 0 To Me.mlWeaponUB
            If mlWeaponIdx(X) <> -1 AndAlso ((bResearchedOnly = False AndAlso moWeapon(X).ComponentDevelopmentPhase < Epica_Tech.eComponentDevelopmentPhase.eResearched) OrElse bResearchedOnly = True AndAlso moWeapon(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched) Then
                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oAttacker, CType(moWeapon(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
            End If
        Next X
    End Sub

    'Pass 0 for no setting, 1 fr cargo, 2 for hangar
    Public Function GetRandomUnitOrFacility(ByVal yType As Byte, ByVal yCargoHangarSetting As Byte) As Epica_Entity
        Try
            Dim lUnitCnt As Int32 = 0
            Dim lFacCnt As Int32 = 0
            If yType = 0 OrElse yType <> 1 Then
                lUnitCnt = Me.oBudget.GetTotalUnitCount()
            End If
            If yType = 1 OrElse yType <> 0 Then
                'FACILITY or UNIT AND FACILITY
                Dim lCurUB As Int32 = Math.Min(glFacilityUB, glFacilityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If glFacilityIdx(X) <> -1 Then
                        Dim oFac As Facility = goFacility(X)
                        If oFac Is Nothing = False AndAlso oFac.Owner.ObjectID = Me.ObjectID Then

                            If yCargoHangarSetting = 1 Then
                                If oFac.EntityDef.Cargo_Cap > 0 Then lFacCnt += 1
                            ElseIf yCargoHangarSetting = 2 Then
                                If oFac.EntityDef.Hangar_Cap > 0 Then lFacCnt += 1
                            Else : lFacCnt += 1
                            End If

                        End If
                    End If
                Next X
            End If

            Dim lTotalCnt As Int32 = lUnitCnt + lFacCnt
            Dim lValue As Int32 = CInt(Rnd() * lTotalCnt)

            'ok, now, what to return...
            If lTotalCnt > lUnitCnt Then
                'a facility
                lTotalCnt -= lUnitCnt
                Dim lCurUB As Int32 = Math.Min(glFacilityUB, glFacilityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If glFacilityIdx(X) <> -1 Then
                        Dim oFac As Facility = goFacility(X)
                        If oFac Is Nothing = False AndAlso oFac.Owner.ObjectID = Me.ObjectID Then
                            If yCargoHangarSetting = 1 Then
                                If oFac.EntityDef.Cargo_Cap > 0 Then lTotalCnt -= 1
                            ElseIf yCargoHangarSetting = 2 Then
                                If oFac.EntityDef.Hangar_Cap > 0 Then lTotalCnt -= 1
                            Else : lTotalCnt -= 1
                            End If
                            If lTotalCnt <= lValue Then Return oFac
                        End If
                    End If
                Next X
            Else
                'a unit
                Dim lCurUB As Int32 = Math.Min(glUnitUB, glUnitIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If glUnitIdx(X) <> -1 Then
                        Dim oUnit As Unit = goUnit(X)
                        If oUnit Is Nothing = False AndAlso oUnit.Owner.ObjectID = Me.ObjectID Then
                            lTotalCnt -= 1
                            If lTotalCnt <= lValue Then Return oUnit
                        End If
                    End If
                Next X
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Public Sub InitiateFullLockdown()

        Me.lStatusFlags = Me.lStatusFlags Or elPlayerStatusFlag.FullInvulnerabilityRaised

        If HasOnlineAliases(0) = True OrElse (oSocket Is Nothing = False AndAlso oSocket.IsConnected = True) OrElse lConnectedPrimaryID > -1 Then
            Return
        End If

        '* Pauses all production and research
        Dim lCurUB As Int32 = -1
        If glFacilityIdx Is Nothing = False Then lCurUB = Math.Min(glFacilityUB, glFacilityIdx.GetUpperBound(0))
        'Force all facilities to recalculate production
        For X As Int32 = 0 To lCurUB
            If glFacilityIdx(X) <> -1 Then
                Dim oFac As Facility = goFacility(X)
                If oFac Is Nothing = False AndAlso oFac.Owner Is Nothing = False AndAlso oFac.Owner.ObjectID = Me.ObjectID Then
                    If oFac.bProducing = True AndAlso oFac.CurrentProduction Is Nothing = False Then
                        oFac.RecalcProduction()
                    End If
                End If
            End If
        Next X
        lCurUB = -1
        'Force all units to recalculate production
        If glUnitIdx Is Nothing = False Then lCurUB = Math.Min(glUnitUB, glUnitIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glUnitIdx(X) <> -1 Then
                Dim oUnit As Unit = goUnit(X)
                If oUnit Is Nothing = False AndAlso oUnit.Owner Is Nothing = False AndAlso oUnit.Owner.ObjectID = Me.ObjectID Then
                    If oUnit.bProducing = True Then
                        oUnit.RecalcProduction()
                    End If
                End If
            End If
        Next X

        '* Stops all unit movement except system to system movement - handled on region server
        '* Pauses income change (cashflow) - handled in budget and other places

        If Me.lIronCurtainPlanet < 1 Then Me.lIronCurtainPlanet = Me.lStartedEnvirID
        'ok, now, make our msg
        Dim yMsg(10) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eSetIronCurtain).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(Me.lIronCurtainPlanet).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = 255 : lPos += 1

        goMsgSys.BroadcastToDomains(yMsg)


        Me.bInFullLockDown = True

    End Sub

    Public Sub WipePlayer()
        'remove all of my tech designs
        For X As Int32 = 0 To mlAlloyUB
            If mlAlloyIdx(X) <> -1 Then
                If moAlloy(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Alloy in WipePlayer did not complete. Player: " & Me.ObjectID & ", AlloyID: " & mlAlloyIdx(X))
                End If
            End If
        Next X
        mlAlloyUB = -1
        ReDim mlAlloyIdx(-1)
        ReDim moAlloy(-1)
        For X As Int32 = 0 To mlArmorUB
            If mlArmorIdx(X) <> -1 Then
                If moArmor(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Armor in WipePlayer did not complete. Player: " & Me.ObjectID & ", ArmorID: " & mlArmorIdx(X))
                End If
            End If
        Next X
        mlArmorUB = -1
        ReDim mlArmorIdx(-1)
        ReDim moArmor(-1)
        For X As Int32 = 0 To mlEngineUB
            If mlEngineIdx(X) <> -1 Then
                If moEngine(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Engine in WipePlayer did not complete. Player: " & Me.ObjectID & ", EngineID: " & mlEngineIdx(X))
                End If
            End If
        Next X
        mlEngineUB = -1
        ReDim mlEngineIdx(-1)
        ReDim moEngine(-1)
        For X As Int32 = 0 To mlHullUB
            If mlHullIdx(X) <> -1 Then
                If moHull(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Hull in WipePlayer did not complete. Player: " & Me.ObjectID & ", HullID: " & mlHullIdx(X))
                End If
            End If
        Next X
        mlHullUB = -1
        ReDim mlHullIdx(-1)
        ReDim moHull(-1)
        For X As Int32 = 0 To mlPrototypeUB
            If mlPrototypeIdx(X) <> -1 Then
                If moPrototype(X).ResultingDef Is Nothing = False Then
                    'delete the resulting def
                    moPrototype(X).ResultingDef.DeleteMe()
                End If

                If moPrototype(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Prototype in WipePlayer did not complete. Player:" & Me.ObjectID & ", PrototypeID: " & mlPrototypeIdx(X))
                End If
            End If
        Next X
        mlPrototypeUB = -1
        ReDim mlPrototypeIdx(-1)
        ReDim moPrototype(-1)
        For X As Int32 = 0 To mlRadarUB
            If mlRadarIdx(X) <> -1 Then
                If moRadar(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Radar in WipePlayer did not complete. Player: " & Me.ObjectID & ", RadarID: " & mlRadarIdx(X))
                End If
            End If
        Next X
        mlRadarUB = -1
        ReDim mlRadarIdx(-1)
        ReDim moRadar(-1)
        For X As Int32 = 0 To mlShieldUB
            If mlShieldIdx(X) <> -1 Then
                If moShield(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Shield in WipePlayer did not complete. Player: " & Me.ObjectID & ", ShieldID: " & mlShieldIdx(X))
                End If
            End If
        Next X
        mlShieldUB = -1
        ReDim mlShieldIdx(-1)
        ReDim moShield(-1)
        For X As Int32 = 0 To mlWeaponUB
            If mlWeaponIdx(X) <> -1 Then
                If moWeapon(X).DeleteMe() = False Then
                    LogEvent(LogEventType.Warning, "Delete Weapon in WipePlayer did not complete. Player: " & Me.ObjectID & ", WeaponID: " & mlWeaponIdx(X))
                End If
            End If
        Next X
        mlWeaponUB = -1
        ReDim mlWeaponIdx(-1)
        ReDim moWeapon(-1)

        'remove all of my mineral knowledge
        For X As Int32 = 0 To Me.lPlayerMineralUB
            If Me.oPlayerMinerals(X).DeleteMe() = False Then
                LogEvent(LogEventType.Warning, "Delete PlayerMineral in WipePlayer did not complete. Player: " & Me.ObjectID & ".")
            End If
        Next X
        lPlayerMineralUB = -1
        ReDim oPlayerMinerals(-1)

        'remove all of my unitdefs and facilitydefs
        For X As Int32 = 0 To glUnitDefUB
            If glUnitDefIdx(X) > -1 Then
                If goUnitDef(X).OwnerID = Me.ObjectID Then
                    goUnitDef(X).DeleteMe()
                    glUnitDefIdx(X) = -1
                End If
            End If
        Next X
        For X As Int32 = 0 To glFacilityDefUB
            If glFacilityDefIdx(X) > -1 Then
                If goFacilityDef(X).OwnerID = Me.ObjectID Then
                    goFacilityDef(X).DeleteMe()
                    glFacilityDefIdx(X) = -1
                End If
            End If
        Next X

        'remove all of my scores
        lMilitaryScore = 0

        'remove all player intels
        If PlayerIntel.DeleteAllPlayerIntels(Me.ObjectID) = False Then
            LogEvent(LogEventType.Warning, "Delete PlayerIntel in WipePlayer did not complete. Player: " & Me.ObjectID & ".")
        End If
        mlPlayerIntelUB = -1
        ReDim mlPlayerIntelIdx(-1)
        ReDim moPlayerIntel(-1)

        'remove all player item intel
        If PlayerItemIntel.DeleteAllPlayerItemIntel(Me.ObjectID) = False Then
            LogEvent(LogEventType.Warning, "Delete PlayerItemIntel in WipePlayer did not complete. Player: " & Me.ObjectID & ".")
        End If
        mlItemIntelUB = -1
        ReDim moItemIntel(-1)

        'remove all player missions
        For X As Int32 = 0 To glPlayerMissionUB
            If glPlayerMissionIdx(X) > -1 AndAlso goPlayerMission(X).oPlayer.ObjectID = Me.ObjectID Then
                goPlayerMission(X).DeleteMe()
                glPlayerMissionIdx(X) = -1
            End If
        Next X

        'remove all player tech knowledge
        If PlayerTechKnowledge.DeleteAllPlayerTechKnowledge(Me.ObjectID) = False Then
            LogEvent(LogEventType.Warning, "Delete PlayerTechKnowledge in WipePlayer did not complete. Player: " & Me.ObjectID & ".")
        End If
        mlPlayerTechKnowledgeUB = -1
        ReDim myPlayerTechKnowledgeUsed(-1)
        ReDim moPlayerTechKnowledge(-1)

        ClearFactionSlots()

        'remove all of my tally systems
        blMaxPopulation = 0
        TotalPlayTime = 0
        Me.RemovePlayerRel(-1, False)
        BadWarDecMoralePenalty = 0
        ClearCPPenaltyList()
        BadWarDecMoralePenaltyEndCycle = 0
        BadWarDecCPIncrease = 0
        BadWarDecCPIncreaseEndCycle = 0
    End Sub


    Public Sub ForceKillMe()
        For X As Int32 = 0 To mlColonyUB
            If mlColonyIdx(X) <> -1 AndAlso glColonyIdx(mlColonyIdx(X)) = mlColonyID(X) Then
                Dim oColony As Colony = goColony(mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    'set our pops to -20
                    oColony.Population = -20
                End If
            End If
        Next X
    End Sub
End Class

