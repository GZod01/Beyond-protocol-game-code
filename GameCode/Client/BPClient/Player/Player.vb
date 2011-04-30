Option Strict On

Public Enum ePlayerSpecialAttributeSetting As Short
    eCelebrationPeriod = 0
    eInPirateSpawn = 1
    eNegativeCashflow = 2
	eDoctrineOfLeadership = 3
	eBadWarDecMoralePenalty = 4
	eBadWarDecCPPenalty = 5
    'TODO: add more here
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

Public Structure PlayerAlias
    Public sPlayerName As String
    Public sUserName As String
    Public sPassword As String
    Public lRights As Int32
End Structure

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

    Magistrate = 2097152
    Governor = 4194304
    Overseer = 8388608
    Duke = 16777216
    Baron = 33554432
    King = 67108864
    Emperor = 134217728

End Enum

Public Enum AccountStatusType As Integer
    eInactiveAccount = 0
    eActiveAccount
    eBannedAccount
    eSuspendedAccount
    eOpenBetaAccount
    eTrialAccount = 99
    eMondelisInactive = 100
    eMondelisActive = 101
End Enum

Public Enum elPlayerStatusFlag As int32
    FullInvulnerabilityRaised = 1
    AlwaysRaiseFullInvul = 2
End Enum

'This covers the data of the current player and any player the current player may know about
Public Class Player
    Inherits Base_GUID

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

    Public PlayerName As String
    Public EmpireName As String
    Public RaceName As String

	'Public lGuildID As Int32 = -1
	'Public lJoinedGuildOn As Int32 = -1
	Public oGuild As Guild = Nothing

    Public lSenateID As Int32 = -1

    Public lStartID As Int32 = -1
    Public iStartTypeID As Int16 = -1

    Public CommEncryptLevel As Int16
    Public EmpireTaxRate As Byte

    Public blCredits As Int64
    Public blCashFlow As Int64 = 0
    'Public blWarpoints As Int64 = 0
    'Public blWarpointsSession As Int64 = -1
    'Public lCurrentWarpointUpkeepCost As Int32 = 0

    Public muAliases() As PlayerAlias
    Public mlAliasUB As Int32 = -1
    Public muAllowances() As PlayerAlias
    Public mlAllowanceUB As Int32 = -1

    Private moPlayerRels() As PlayerRel
    Private mlPlayerRelIdx() As Int32
    Public PlayerRelUB As Int32 = -1

    Public moTechs() As Base_Tech
    Public mlTechUB As Int32 = -1

    Public moUnitGroups() As UnitGroup
    Public mlUnitGroupIdx() As Int32
    Public mlUnitGroupUB As Int32 = -1

    Public moEmailFolders() As PlayerCommFolder
    Public mlEmailFolderIdx() As Int32
    Public mlEmailFolderUB As Int32 = -1

    Public moColonies() As Colony
    Public mlColonyIdx() As Int32
    Public mlColonyUB As Int32 = -1

    Public moTrades() As Trade
    Public mlTradeIdx() As Int32
    Public mlTradeUB As Int32 = -1

    Public moTradeHistory() As TradeHistory
    Public mlTradeHistoryUB As Int32 = -1

    Public oBudget As Budget

    Public yPlayerTitle As Byte
    Public yCustomTitle As Byte
    Public lCustomTitlePermission As Int32
    Public bIsMale As Boolean = True
    Public lPlayerIcon As Int32

    Public DeathBudgetBalance As Int32 = 0

    Public PlayerMissions() As PlayerMission
    Public PlayerMissionIdx() As Int32
    Public PlayerMissionUB As Int32 = -1
    Public Agents() As Agent
    Public AgentIdx() As Int32
	Public AgentUB As Int32 = -1
	Public CapturedAgents() As CapturedAgent
	Public CapturedAgentIdx() As Int32
	Public CapturedAgentUB As Int32 = -1

    'These are the scores of the player...
    Public lTechnologyScore As Int32
    Public lDiplomacyScore As Int32
    Public lMilitaryScore As Int32
    Public lPopulationScore As Int32
    Public lProductionScore As Int32
    Public lWealthScore As Int32
    Public lTotalScore As Int32
    Public lControlPlanets() As Int32       'ID's of planets that the player has a stake interest in
    Public yControlPlanetAmt() As Byte      'percentage stored as byte indicating majority ownership share

    Public lCelebrationPeriodEnd As Int32 = 0
    Public bInNegativeCashFlow As Boolean = False
    Public bInPirateSpawn As Boolean = False
    Public yDoctrineOfLeadership As Byte = eDoctrineOfLeadershipSetting.eNoSetting

    Private mlSpecialTechProperty() As Int32
    Private mlSpecialTechPropertyValue() As Int32
	Private mlSuperSpecials As Int32 = 0

	Public BadWarDecCPIncrease As Int32 = 0
	Public BadWarDecMoralePenalty As Int32 = 0

	Public lFormationIdx() As Int32
	Public lFormationUB As Int32 = -1
	Public oFormations() As FormationDef

	Public yPlayerPhase As Byte
    Public lTutorialStep As Int32 = 1
    Public lAccountStatus As Int32 = 0

    Public lStatusFlags As Int32 = 0

#Region "  Socketed Faction Management  "
	Public lSlotID(4) As Int32
	Public ySlotState(4) As Byte
	Public lFactionID(2) As Int32
	Public yFactionState(2) As Byte
	'Public yFactionBonus As Byte = 0
	'Public yOtherFactionBonus As Byte = 0
	'Public yCurrentResearchBonus As Byte = 0

	Public Sub HandleUpdateSlotMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 6	'msgcode and playerid

		ReDim lSlotID(4)
		ReDim ySlotState(4)
		ReDim lFactionID(2)
		ReDim yFactionState(2)

		lSlotID(0) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ySlotState(0) = yData(lPos) : lPos += 1
		lSlotID(1) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ySlotState(1) = yData(lPos) : lPos += 1
		lSlotID(2) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ySlotState(2) = yData(lPos) : lPos += 1
		lSlotID(3) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ySlotState(3) = yData(lPos) : lPos += 1
		lSlotID(4) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ySlotState(4) = yData(lPos) : lPos += 1

		lFactionID(0) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yFactionState(0) = yData(lPos) : lPos += 1
		lFactionID(1) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yFactionState(1) = yData(lPos) : lPos += 1
		lFactionID(2) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yFactionState(2) = yData(lPos) : lPos += 1

		'yFactionBonus = yData(lPos) : lPos += 1
		'yOtherFactionBonus = yData(lPos) : lPos += 1
		'yCurrentResearchBonus = yData(lPos) : lPos += 1
	End Sub
#End Region

    Public Sub HandleUpdatePlayerTechValue(ByRef yData() As Byte)
        Try
            Dim lPos As Int32 = 6       'msgcode and playerid
            mlSuperSpecials = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            ReDim mlSpecialTechProperty(lCnt - 1)
            ReDim mlSpecialTechPropertyValue(lCnt - 1)
            For X As Int32 = 0 To lCnt - 1
                mlSpecialTechProperty(X) = X
                mlSpecialTechPropertyValue(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Next X
        Catch
            'Do nothing?
            'TODO: Possible cheat?
        End Try
    End Sub

    Public Shared Function GetPlayerNameWithTitle(ByVal pyTitle As Byte, ByVal sName As String, ByVal pbIsMale As Boolean) As String

        Dim bTrade As Boolean = (pyTitle And eyCustomRank.TraderShift) <> 0
        Dim bDiplomacy As Boolean = (pyTitle And eyCustomRank.DiplomacyShift) <> 0
        Dim bResearch As Boolean = (pyTitle And eyCustomRank.ResearcherShift) <> 0

        Dim sTemp As String = ""
        Dim bIsExRank As Boolean = (pyTitle And Player.PlayerRank.ExRankShift) <> 0
        If bIsExRank = True Then pyTitle = pyTitle Xor Player.PlayerRank.ExRankShift
        If bTrade = True Then pyTitle = pyTitle Xor eyCustomRank.TraderShift
        If bDiplomacy = True Then pyTitle = pyTitle Xor eyCustomRank.DiplomacyShift
        If bResearch = True Then pyTitle = pyTitle Xor eyCustomRank.ResearcherShift


        If bDiplomacy = True Then
            Select Case pyTitle
                Case PlayerRank.Baron
                    sTemp = "High Senator "
                Case PlayerRank.Duke
                    sTemp = "Senator "
                Case PlayerRank.Emperor
                    sTemp = "Supreme Chancellor "
                Case PlayerRank.Governor
                    If pbIsMale = True Then sTemp = "Arbiter " Else sTemp = "Arbitress "
                Case PlayerRank.King
                    sTemp = "Chancellor "
                Case PlayerRank.Magistrate
                    sTemp = "Diplomat "
                Case PlayerRank.Overseer
                    sTemp = "Counselor "
            End Select
        ElseIf bResearch = True Then
            Select Case pyTitle
                Case PlayerRank.Baron
                    sTemp = "Chief Scientist "
                Case PlayerRank.Duke
                    sTemp = "Inquisitor "
                Case PlayerRank.Emperor
                    sTemp = "Transcendent "
                Case PlayerRank.Governor
                    sTemp = "Scientist "
                Case PlayerRank.King
                    sTemp = "Preeminence "
                Case PlayerRank.Magistrate
                    sTemp = "Explorer "
                Case PlayerRank.Overseer
                    sTemp = "Master Scientist "
            End Select
        ElseIf bTrade = True Then
            Select Case pyTitle
                Case PlayerRank.Baron
                    sTemp = "Chief Broker "
                Case PlayerRank.Duke
                    If pbIsMale = True Then sTemp = "Trade Lord " Else sTemp = "Trade Lady "
                Case PlayerRank.Emperor
                    If pbIsMale = True Then sTemp = "Commerce Czar " Else sTemp = "Commerce Czarina "
                Case PlayerRank.Governor
                    sTemp = "Merchant "
                Case PlayerRank.King
                    sTemp = "Master Merchant "
                Case PlayerRank.Magistrate
                    sTemp = "Trader "
                Case PlayerRank.Overseer
                    sTemp = "Broker "
            End Select
        Else
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

            If bIsExRank = True Then sTemp = "Ex-" & sTemp
        End If

        Return sTemp & sName
    End Function
    Public Shared Function GetPlayerTitle(ByVal pyTitle As Byte, ByVal pbIsMale As Boolean) As String
        Dim bTrade As Boolean = (pyTitle And eyCustomRank.TraderShift) <> 0
        Dim bDiplomacy As Boolean = (pyTitle And eyCustomRank.DiplomacyShift) <> 0
        Dim bResearch As Boolean = (pyTitle And eyCustomRank.ResearcherShift) <> 0

        Dim sTemp As String = ""
        Dim bIsExRank As Boolean = (pyTitle And Player.PlayerRank.ExRankShift) <> 0
        If bIsExRank = True Then pyTitle = pyTitle Xor Player.PlayerRank.ExRankShift
        If bTrade = True Then pyTitle = pyTitle Xor eyCustomRank.TraderShift
        If bDiplomacy = True Then pyTitle = pyTitle Xor eyCustomRank.DiplomacyShift
        If bResearch = True Then pyTitle = pyTitle Xor eyCustomRank.ResearcherShift


        If bDiplomacy = True Then
            Select Case pyTitle
                Case PlayerRank.Baron
                    sTemp = "High Senator "
                Case PlayerRank.Duke
                    sTemp = "Senator "
                Case PlayerRank.Emperor
                    sTemp = "Supreme Chancellor "
                Case PlayerRank.Governor
                    If pbIsMale = True Then sTemp = "Arbiter " Else sTemp = "Arbitress "
                Case PlayerRank.King
                    sTemp = "Chancellor "
                Case PlayerRank.Magistrate
                    sTemp = "Diplomat "
                Case PlayerRank.Overseer
                    sTemp = "Counselor "
            End Select
        ElseIf bResearch = True Then
            Select Case pyTitle
                Case PlayerRank.Baron
                    sTemp = "Chief Scientist "
                Case PlayerRank.Duke
                    sTemp = "Inquisitor "
                Case PlayerRank.Emperor
                    sTemp = "Transcendent "
                Case PlayerRank.Governor
                    sTemp = "Scientist "
                Case PlayerRank.King
                    sTemp = "Preeminence "
                Case PlayerRank.Magistrate
                    sTemp = "Explorer "
                Case PlayerRank.Overseer
                    sTemp = "Master Scientist "
            End Select
        ElseIf bTrade = True Then
            Select Case pyTitle
                Case PlayerRank.Baron
                    sTemp = "Chief Broker "
                Case PlayerRank.Duke
                    If pbIsMale = True Then sTemp = "Trade Lord " Else sTemp = "Trade Lady "
                Case PlayerRank.Emperor
                    If pbIsMale = True Then sTemp = "Commerce Czar " Else sTemp = "Commerce Czarina "
                Case PlayerRank.Governor
                    sTemp = "Merchant "
                Case PlayerRank.King
                    sTemp = "Master Merchant "
                Case PlayerRank.Magistrate
                    sTemp = "Trader "
                Case PlayerRank.Overseer
                    sTemp = "Broker "
            End Select
        Else
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

            'If bIsExRank = True Then sTemp = "Ex-" & sTemp
        End If

        Return sTemp '& sName
    End Function

    Public Function GetOrAddTrade(ByVal lTradeID As Int32) As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1
        For X As Int32 = 0 To mlTradeUB
            If mlTradeIdx(X) = lTradeID Then
                lIdx = X
                Exit For
            ElseIf lFirstIdx = -1 AndAlso mlTradeIdx(X) = -1 Then
                lFirstIdx = X
            End If
        Next X

        If lIdx = -1 Then
            If lFirstIdx = -1 Then
                ReDim Preserve moTrades(mlTradeUB + 1)
                ReDim Preserve mlTradeIdx(mlTradeUB + 1)
                mlTradeIdx(mlTradeUB + 1) = -1
                mlTradeUB += 1
                lIdx = mlTradeUB
            Else : lIdx = lFirstIdx
            End If
            moTrades(lIdx) = New Trade()
            mlTradeIdx(lIdx) = lTradeID
        End If

        Return lIdx

    End Function

    Public Sub SetPlayerRel(ByVal lTargetPlayerID As Int32, ByVal yScore As Byte)
        Dim X As Int32
        Dim lIdx As Int32 = -1

        For X = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) = lTargetPlayerID Then
                moPlayerRels(X).WithThisScore = yScore
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

        moPlayerRels(lIdx) = New PlayerRel
        With moPlayerRels(lIdx)
            .lPlayerRegards = Me.ObjectID
            .lThisPlayer = lTargetPlayerID
            .WithThisScore = yScore
        End With
        mlPlayerRelIdx(lIdx) = lTargetPlayerID
    End Sub

    Public Function GetPlayerRelScore(ByVal lPlayerID As Int32) As Byte
        If lPlayerID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return elRelTypes.eWar

        Try
            For X As Int32 = 0 To PlayerRelUB
                If mlPlayerRelIdx(X) = lPlayerID Then
                    Return moPlayerRels(X).WithThisScore
                End If
            Next X
        Catch
            Return 0
        End Try
        Return elRelTypes.eNeutral
    End Function

    Public Function GetPlayerRelByIndex(ByVal lIndex As Int32) As PlayerRel
        If lIndex < 0 OrElse lIndex > PlayerRelUB Then Return Nothing
        Return moPlayerRels(lIndex)
    End Function

    Public Function GetPlayerRel(ByVal lPlayerID As Int32) As PlayerRel
        For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) = lPlayerID Then
                Return moPlayerRels(X)
            End If
        Next X
        Return Nothing
	End Function

	Public Sub RemovePlayerRel(ByVal lWithPlayerID As Int32)
		If lWithPlayerID = -1 Then
			PlayerRelUB = -1
		Else
			Dim lShiftIdx As Int32 = -1
			For X As Int32 = 0 To PlayerRelUB
				If mlPlayerRelIdx(X) = lWithPlayerID Then
					lShiftIdx = X
					If goUILib Is Nothing = False Then
						goUILib.AddNotification("We have lost contact with " & GetCacheObjectValue(lWithPlayerID, ObjectType.ePlayer) & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					End If
					Exit For
				End If
            Next X
            If lShiftIdx = -1 Then Return
			For X As Int32 = lShiftIdx To PlayerRelUB - 1
				moPlayerRels(X) = moPlayerRels(X + 1)
				mlPlayerRelIdx(X) = mlPlayerRelIdx(X + 1)
			Next X
			PlayerRelUB -= 1
		End If
	End Sub

    Public Sub AddTech(ByRef oTech As Object)
        Dim lIdx As Int32 = -1

        With CType(oTech, Base_Tech)

            For X As Int32 = 0 To mlTechUB
                If moTechs(X).ObjectID = .ObjectID AndAlso moTechs(X).ObjTypeID = .ObjTypeID Then
                    If moTechs(X).ComponentDevelopmentPhase < .ComponentDevelopmentPhase AndAlso moTechs(X).ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
                        If goUILib Is Nothing = False Then
                            goUILib.AddNotification("Research completed on " & .GetComponentName & "!", Color.Teal, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        End If
                    End If
                    moTechs(X) = CType(oTech, Base_Tech)
                    lIdx = X
                    Exit For
                End If
            Next X

            'Sort the addition
            If lIdx = -1 Then
                'Make room for the result...
                mlTechUB += 1
                ReDim Preserve moTechs(mlTechUB)


                'Ok, we'll need to sort here...
                Dim sCompName As String = .GetComponentName
                For X As Int32 = 0 To mlTechUB - 1
                    If moTechs(X).GetComponentName.ToLower > sCompName.ToLower Then
                        'Ok, here is where we put it... push all other elements back
                        For Y As Int32 = mlTechUB To X + 1 Step -1
                            moTechs(Y) = moTechs(Y - 1)
                        Next Y

                        lIdx = X
                        Exit For
                    End If
                Next

                If lIdx = -1 Then
                    'Ok, add it to the end
                    lIdx = mlTechUB
                End If

                moTechs(lIdx) = CType(oTech, Base_Tech)
            End If
        End With
    End Sub

    Public Function GetTech(ByVal lID As Int32, ByVal iTypeID As Int16) As Base_Tech
        For X As Int32 = 0 To mlTechUB
            If moTechs(X) Is Nothing = False AndAlso moTechs(X).ObjectID = lID AndAlso moTechs(X).ObjTypeID = iTypeID Then
                Return moTechs(X)
            End If
        Next X

        For X As Int32 = 0 To glPlayerTechKnowledgeUB
            If glPlayerTechKnowledgeIdx(X) = lID AndAlso goPlayerTechKnowledge(X).oTech.ObjTypeID = iTypeID Then
                Return goPlayerTechKnowledge(X).oTech
            End If
        Next X

        Return Nothing
    End Function

    Public Sub RemoveTech(ByVal lID As Int32, ByVal iTypeID As Int16)
        For X As Int32 = 0 To mlTechUB
            If moTechs(X) Is Nothing = False AndAlso moTechs(X).ObjectID = lID AndAlso moTechs(X).ObjTypeID = iTypeID Then
                moTechs(X) = Nothing
                'Now, shift all other techs up
                For Y As Int32 = X To mlTechUB - 1
                    moTechs(Y) = moTechs(Y + 1)
                Next Y
                mlTechUB -= 1
                Exit For
            End If
        Next X
    End Sub

    'NOTE: A hard-code in the game
    Public Function PlayerKnowsProperty(ByVal sPropName As String) As Boolean
        sPropName = sPropName.ToUpper.Trim()

        For X As Int32 = 0 To glMineralPropertyUB
            If glMineralPropertyIdx(X) <> -1 AndAlso goMineralProperty(X).MineralPropertyName.ToUpper.Trim() = sPropName Then
                Return goMineralProperty(X).yKnowledgeLevel > 0
            End If
        Next X
        Return False
    End Function

    'Public Function SpecialTechResearched(ByVal lProgCtrlID As Int32) As Boolean
    '    Dim oTech As Epica_Tech = GetTech(lProgCtrlID, ObjectType.eSpecialTech)
    '    If oTech Is Nothing = False AndAlso oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then Return True
    '    Return False
    'End Function  
    Public Function GetSpecialTechPropertyValue(ByVal lProgCtrlID As Int32) As Int32
        If lProgCtrlID = PlayerSpecialAttributeID.eSuperSpecials Then Return mlSuperSpecials

        If mlSpecialTechProperty Is Nothing = False Then
            For X As Int32 = 0 To mlSpecialTechProperty.GetUpperBound(0)
                If mlSpecialTechProperty(X) = lProgCtrlID Then
                    Return mlSpecialTechPropertyValue(X)
                End If
            Next X
        End If
        Return 0
    End Function
 
    Public ReadOnly Property ShowSpaceDots() As Boolean
        Get 
            Return (mlSuperSpecials And 65536) <> 0
        End Get
    End Property
    Public ReadOnly Property StationPlacementCloserToPlanet() As Boolean
        Get
            Return (mlSuperSpecials And 8) <> 0
        End Get
    End Property

    Public Sub AddUnitGroupFromMsg(ByRef yData() As Byte, ByVal lObjID As Int32)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlUnitGroupUB
            If mlUnitGroupIdx(X) = lObjID Then
                moUnitGroups(X).FillFromMsg(yData)
                Return
            ElseIf lIdx = -1 AndAlso mlUnitGroupIdx(X) = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx = -1 Then
            mlUnitGroupUB += 1
            ReDim Preserve moUnitGroups(mlUnitGroupUB)
            ReDim Preserve mlUnitGroupIdx(mlUnitGroupUB)
            lIdx = mlUnitGroupUB
        End If

        mlUnitGroupIdx(lIdx) = lObjID
        moUnitGroups(lIdx) = New UnitGroup()
        moUnitGroups(lIdx).FillFromMsg(yData)
    End Sub

    Public Function AddEmailFolder(ByVal lPCF_ID As Int32, ByVal sName As String) As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1


        For X As Int32 = 0 To mlEmailFolderUB
            If mlEmailFolderIdx(X) = lPCF_ID Then
                lIdx = X
                Exit For
            ElseIf lFirstIdx = -1 AndAlso mlEmailFolderIdx(X) = -1 Then
                lFirstIdx = X
            End If
        Next X

        If lIdx = -1 Then
            If lFirstIdx = -1 Then
                mlEmailFolderUB += 1
                ReDim Preserve mlEmailFolderIdx(mlEmailFolderUB)
                ReDim Preserve moEmailFolders(mlEmailFolderUB)
                lIdx = mlEmailFolderUB
            Else : lIdx = lFirstIdx
            End If
            moEmailFolders(lIdx) = New PlayerCommFolder()
            mlEmailFolderIdx(lIdx) = lPCF_ID
        End If

        With moEmailFolders(lIdx)
            .FolderName = sName
            .PCF_ID = lPCF_ID
            .PlayerID = Me.ObjectID
        End With

        Dim oNewFolder As PlayerCommFolder = moEmailFolders(lIdx)
        If lPCF_ID > 0 AndAlso mlEmailFolderUB > 4 Then
            'Resort
            For Y As Int32 = 4 To mlEmailFolderUB - 1
                If moEmailFolders(Y).FolderName.ToUpper > sName.ToUpper Then
                    For Z As Int32 = mlEmailFolderUB To Y Step -1
                        moEmailFolders(Z) = moEmailFolders(Z - 1)
                        mlEmailFolderIdx(Z) = mlEmailFolderIdx(Z - 1)
                    Next Z
                    mlEmailFolderIdx(Y) = lPCF_ID
                    moEmailFolders(Y) = oNewFolder
                    Exit For
                End If
            Next
        End If
        oNewFolder = Nothing

        Return lIdx
    End Function

    Public Sub AddPlayerComm(ByRef oComm As PlayerComm)
        'Ok... figure out where to put it
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlEmailFolderUB
            If mlEmailFolderIdx(X) = oComm.PCF_ID Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            'Did not find a mailbox... determine what type of mailbox it would have been put into
            Dim sName As String = "Unknown"
            Select Case oComm.PCF_ID
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF
                    sName = "Deleted"
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eDrafts_PCF
                    sName = "Drafts"
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF
                    sName = "Inbox"
                Case PlayerCommFolder.ePCF_ID_HARDCODES.eOutbox_PCF
                    sName = "Outbox"
            End Select

            lIdx = AddEmailFolder(oComm.PCF_ID, sName)
        End If

        If lIdx <> -1 Then
            Dim lMsgIdx As Int32 = -1
            With moEmailFolders(lIdx)
                For X As Int32 = 0 To .PlayerMsgUB
                    If .PlayerMsgsIdx(X) = oComm.ObjectID Then
						'Msg is already there
						.PlayerMsgs(X) = oComm
                        Return
                    ElseIf .PlayerMsgsIdx(X) = -1 Then
                        lMsgIdx = X
                    End If
                Next X

                If lMsgIdx = -1 Then
                    .PlayerMsgUB += 1
                    ReDim Preserve .PlayerMsgs(.PlayerMsgUB)
                    ReDim Preserve .PlayerMsgsIdx(.PlayerMsgUB)
                    lMsgIdx = .PlayerMsgUB
                End If

                .PlayerMsgs(lMsgIdx) = oComm
                .PlayerMsgsIdx(lMsgIdx) = oComm.ObjectID

                .bHasBeenSorted = False
            End With
        End If

    End Sub

    Public Function GetPlayerTradeCost() As Single
        Dim lVal As Int32 = GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eTradeCosts)
        Return 1.0F + (lVal / 100.0F)
    End Function
    Public Function GetPlayerSellSlots() As Byte 
        Dim lVal As Int32 = GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBuyAndSellOrderSlots)
        Select Case lVal
            Case 1
				Return 15 '10
            Case 2
				Return 20 '15
            Case 3
				Return 25 '20
            Case 4
				Return 30 '25
            Case Else
				Return 10 '5
        End Select
    End Function
    Public Function GetPlayerBuySlots() As Byte
        Dim lVal As Int32 = GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBuyAndSellOrderSlots)
        Select Case lVal
            Case 2
				Return 2
            Case 3
                Return 3
            Case 4
                Return 5
            Case Else
				Return 1
        End Select
    End Function

    Public Sub SetPlayerSpecialAttributeSetting(ByVal iSetting As ePlayerSpecialAttributeSetting, ByVal lValue As Int32)
        Select Case iSetting
            Case ePlayerSpecialAttributeSetting.eCelebrationPeriod
                Me.lCelebrationPeriodEnd = glCurrentCycle + lValue
			Case ePlayerSpecialAttributeSetting.eInPirateSpawn
                If Me.bInPirateSpawn = True AndAlso lValue = 0 Then

                    If NewTutorialManager.TutorialOn = True Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePirateBaseKilled, -1, -1, -1, "")
                    End If
                    'If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
                    '    If goTutorial.EventTriggered(TutorialManager.TutorialTriggerType.PiratesEliminated) = False Then
                    '        goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.PiratesEliminated)
                    '    End If
                    'End If
                End If
				Me.bInPirateSpawn = (lValue <> 0)
			Case ePlayerSpecialAttributeSetting.eNegativeCashflow
				Me.bInNegativeCashFlow = (lValue <> 0)
			Case ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty
				Me.BadWarDecCPIncrease = lValue
			Case ePlayerSpecialAttributeSetting.eBadWarDecMoralePenalty
				Me.BadWarDecMoralePenalty = lValue
		End Select
    End Sub

End Class

Public Structure PlayerIconManager
    Private myInstanceMember As Byte

    Public Enum PlayerIconColor As Byte
        eBlack = 0
        eWhite
        eDarkGray       '64,64,64
        eSilver         '192,192,192
        eRed            '255, 0, 0
        eDarkRed        '128, 0, 0
        eGold           '192,192,0
        eGreen          '0, 255, 0
        eDarkGreen      '0, 128, 0
        eBlue           '64,64,255
        eDarkBlue       '0, 0, 128
        eTeal           '0, 255, 255
        ePurple         '255, 0, 255
        eDarkPurple     '128, 0, 128
        eOrange         '255, 128, 0
        eBrown          '127, 64, 0
    End Enum

	Public Shared Function CreatePlayerIconNumber(ByVal yBackImg As Byte, ByVal yBackClr As Byte, ByVal yFore1Img As Byte, ByVal yFore1Clr As Byte, ByVal yFore2Img As Byte, ByVal yFore2Clr As Byte) As Int32
		Try
			Dim sHex(7) As String

			sHex(0) = Hex(yBackImg)
			sHex(1) = Hex(yBackClr)
			Dim sTemp As String = Hex(yFore1Img)
			If sTemp.Length > 1 Then
				sHex(2) = sTemp.Substring(0, 1)
				sHex(3) = sTemp.Substring(1, 1)
			Else
				sHex(2) = "0"
				sHex(3) = sTemp
			End If

			sTemp = Hex(yFore2Img)
			If sTemp.Length > 1 Then
				sHex(4) = sTemp.Substring(0, 1)
				sHex(5) = sTemp.Substring(1, 1)
			Else
				sHex(4) = "0"
				sHex(5) = sTemp
			End If

			sHex(6) = Hex(yFore1Clr)
			sHex(7) = Hex(yFore2Clr)

			sTemp = "&H" & Join(sHex)

			Return CInt(Val(sTemp.Replace(" ", "")))
		Catch
			If goUILib Is Nothing = False Then goUILib.AddNotification("Invalid icon, please select a valid icon.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End Try
		Return 0
	End Function

    Public Shared Sub FillIconValues(ByVal lIcon As Int32, ByRef yBackImg As Byte, ByRef yBackClr As Byte, ByRef yFore1Img As Byte, ByRef yFore1Clr As Byte, ByRef yFore2Img As Byte, ByRef yFore2Clr As Byte)
        Dim yVals() As Byte = System.BitConverter.GetBytes(lIcon)
        Dim lVal As Int32 = yVals(0) And 240

        yFore1Clr = CByte(DownShiftValue(lVal))
        lVal = yVals(0) And 15
        yFore2Clr = CByte(lVal)

        yFore2Img = yVals(1)
        yFore1Img = yVals(2)

        lVal = yVals(3) And 240
        yBackImg = CByte(DownShiftValue(lVal))
        lVal = yVals(3) And 15
        yBackClr = CByte(lVal)
    End Sub

    Private Shared Function DownShiftValue(ByVal lVal As Int32) As Int32
        Dim lResult As Int32 = 0

        If (lVal And 128) <> 0 Then lResult += 8
        If (lVal And 64) <> 0 Then lResult += 4
        If (lVal And 32) <> 0 Then lResult += 2
        If (lVal And 16) <> 0 Then lResult += 1

        Return lResult
    End Function

    Public Shared Function ReturnImageRectangle(ByVal yValue As Byte, ByVal bIsBackImage As Boolean) As Rectangle
        Dim lDivisor As Int32 = 8
        If bIsBackImage = True Then lDivisor = 4

        Dim lY As Int32 = (yValue \ lDivisor)
        Dim lX As Int32 = (yValue - (lY * lDivisor))
        lX *= 64
        lY *= 64

        Return Rectangle.FromLTRB(lX, lY, lX + 64, lY + 64)
    End Function

    Public Shared Function GetColorValue(ByVal yValue As Byte) As System.Drawing.Color
        Select Case yValue
            Case PlayerIconColor.eBlack
                Return Color.FromArgb(255, 0, 0, 0)
            Case PlayerIconColor.eBlue
                Return Color.FromArgb(255, 64, 64, 255)
            Case PlayerIconColor.eBrown
                Return Color.FromArgb(255, 127, 64, 0)
            Case PlayerIconColor.eDarkBlue
                Return Color.FromArgb(255, 0, 0, 128)
            Case PlayerIconColor.eDarkGray
                Return Color.FromArgb(255, 64, 64, 64)
            Case PlayerIconColor.eDarkGreen
                Return Color.FromArgb(255, 0, 128, 0)
            Case PlayerIconColor.eDarkPurple
                Return Color.FromArgb(255, 128, 0, 128)
            Case PlayerIconColor.eDarkRed
                Return Color.FromArgb(255, 128, 0, 0)
            Case PlayerIconColor.eGold
                Return Color.FromArgb(255, 192, 192, 0)
            Case PlayerIconColor.eGreen
                Return Color.FromArgb(255, 0, 255, 0)
            Case PlayerIconColor.eOrange
                Return Color.FromArgb(255, 255, 128, 0)
            Case PlayerIconColor.ePurple
                Return Color.FromArgb(255, 255, 0, 255)
            Case PlayerIconColor.eRed
                Return Color.FromArgb(255, 255, 0, 0)
            Case PlayerIconColor.eSilver
                Return Color.FromArgb(255, 192, 192, 192)
            Case PlayerIconColor.eTeal
                Return Color.FromArgb(255, 0, 255, 255)
            Case PlayerIconColor.eWhite
                Return Color.FromArgb(255, 255, 255, 255)
        End Select
        Return Color.White
    End Function
End Structure