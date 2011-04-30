Option Strict On

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

Public Enum ePlayerSpecialAttributeSetting As Short
	eCelebrationPeriod = 0
	eInPirateSpawn = 1
	eNegativeCashflow = 2
	eDoctrineOfLeadership = 3
	eBadWarDecMoralePenalty = 4
	eBadWarDecCPPenalty = 5
End Enum

Public Class Player
    Inherits Epica_GUID

    Public oSocket As NetSock       'the socket object this player is using (pointer)
    Public oAliases() As Player
    Public lAliasIdx() As Int32
    Public uAliasLogin() As AliasLogin
    Public lAliasUB As Int32 = -1

	Public lAliasingPlayerID As Int32 = -1
    Public PlayerName(19) As Byte
    Public EmpireName(19) As Byte
    Public RaceName(19) As Byte
    Public PlayerUserName(19) As Byte
	Public PlayerPassword(19) As Byte

	Public BadWarDecCPIncrease As Int32 = 0

    Public AccountStatus As Int32 = 0

    Public SenateID As Int32        'not used yet, but we need to define an object for it

    Public lEnvirIdx As Int32       'index of the environment the player is a part of

    Public CriticalHitMax As Byte = 2 'indicates the highest number a critical hit can occur on as determined by the race

    Public CommEncryptLevel As Int16
    Public EmpireTaxRate As Byte

    Public lStartEnvirID As Int32
    Public lStartLocX As Int32
    Public lStartLocZ As Int32

    Public lPirateStartLocX As Int32
    Public lPirateStartLocZ As Int32

    Private moPlayerRels() As PlayerRel
    Private mlPlayerRelIdx() As Int32
    Public PlayerRelUB As Int32 = -1

    Public lCPLimit As Int32 = 300

    Public lGuildID As Int32 = -1
    Public lJoinedGuildOn As Int32 = -1

    Public yDoctrineOfLeadership As Byte = eDoctrineOfLeadershipSetting.eNoSetting

    Public moKnownWormholes() As Wormhole
	Public mlKnownWormholeUB As Int32 = -1

    Public lIronCurtainPlanetID As Int32 = -1       'cleared for all
    Public iIronCurtainStatus As Byte = 0

	Public yPlayerPhase As Byte = 0

    Public Sub HandleCheckFirstContactWithWormhole(ByRef oWormhole As Wormhole, ByVal lSystemID As Int32)
        For X As Int32 = 0 To mlKnownWormholeUB
            If moKnownWormholes(X).ObjectID = oWormhole.ObjectID Then Return
        Next X
        mlKnownWormholeUB += 1
        ReDim Preserve moKnownWormholes(mlKnownWormholeUB)
        moKnownWormholes(mlKnownWormholeUB) = oWormhole

        Dim yMsg(13) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.ePlayerDiscoversWormhole).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
		System.BitConverter.GetBytes(oWormhole.ObjectID).CopyTo(yMsg, 6)
		System.BitConverter.GetBytes(lSystemID).CopyTo(yMsg, 10)
		goMsgSys.SendToPrimary(yMsg)
	End Sub
	Public Sub AddWormholeKnowledge(ByRef oWormhole As Wormhole)
		For X As Int32 = 0 To mlKnownWormholeUB
			If moKnownWormholes(X).ObjectID = oWormhole.ObjectID Then Return
		Next X
		mlKnownWormholeUB += 1
		ReDim Preserve moKnownWormholes(mlKnownWormholeUB)
		moKnownWormholes(mlKnownWormholeUB) = oWormhole
	End Sub

#Region " Environment Alerts Functionality "
	Private mlEnvirAlertIdx() As Int32			'EnvirID
	Private muEnvirAlert() As EnvirAlert
	Private mlEnvirAlertUB As Int32 = -1

	Private mlLastAlertCheck As Int32			'this is to avoid tests on a per-cycle basis

    Public Sub CheckForAlert(ByVal yAlertType As PlayerAlertType, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lEnemyID As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32)
        If glCurrentCycle <> mlLastAlertCheck Then
            mlLastAlertCheck = glCurrentCycle

            'Now, test... find that environment index...
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To mlEnvirAlertUB
                If mlEnvirAlertIdx(X) = lEnvirID AndAlso muEnvirAlert(X).iEnvirTypeID = iEnvirTypeID AndAlso muEnvirAlert(X).yAlertType = yAlertType Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                mlEnvirAlertUB += 1
                ReDim Preserve mlEnvirAlertIdx(mlEnvirAlertUB)
                ReDim Preserve muEnvirAlert(mlEnvirAlertUB)

                lIdx = mlEnvirAlertUB
                With muEnvirAlert(lIdx)
                    .lEnvirID = lEnvirID
                    .iEnvirTypeID = iEnvirTypeID
                    .yAlertType = yAlertType
                End With
                mlEnvirAlertIdx(mlEnvirAlertUB) = lEnvirID
            End If

            'Always update our location because the Primary server may ask us for the most recent location
            muEnvirAlert(lIdx).LocX = lLocX
            muEnvirAlert(lIdx).LocZ = lLocZ

            If muEnvirAlert(lIdx).CheckForAlert() = True Then
                'Ok, let's do it
                Dim yMsg(50) As Byte

                'PlayerAlertCode - 2 bytes
                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAlert).CopyTo(yMsg, 0)
                'PlayerAlertType - 1 byte
                yMsg(2) = yAlertType
                'EntityGUID - 6
                System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg, 3)
                System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, 7)
                'EnvirGUID - 6
                System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, 9)
                System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, 13)
                'PlayerID - 4
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 15)
                'EnemyID - 4
                System.BitConverter.GetBytes(lEnemyID).CopyTo(yMsg, 19)
                'LocX -4
                System.BitConverter.GetBytes(lLocX).CopyTo(yMsg, 23)
                'LocZ - 4
                System.BitConverter.GetBytes(lLocZ).CopyTo(yMsg, 27)

                goMsgSys.SendToPrimary(yMsg)
            End If

        End If
    End Sub

	Public Function GetAlertLoc(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal yType As PlayerAlertType) As Point
		For X As Int32 = 0 To mlEnvirAlertUB
			If mlEnvirAlertIdx(X) = lEnvirID AndAlso muEnvirAlert(X).iEnvirTypeID = iEnvirTypeID AndAlso muEnvirAlert(X).yAlertType = yType Then
				Return New Point(muEnvirAlert(X).LocX, muEnvirAlert(X).LocZ)
			End If
		Next X
		Return Point.Empty
	End Function
#End Region

	Public Sub SetPlayerRel(ByVal lTargetPlayerID As Int32, ByRef oPlayerRel As PlayerRel)
		Dim X As Int32
		Dim lIdx As Int32 = -1

		For X = 0 To PlayerRelUB
			If mlPlayerRelIdx(X) = lTargetPlayerID Then
				moPlayerRels(X) = oPlayerRel
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
	End Sub

	Public Function HasPlayerRelationship(ByVal lPlayerID As Int32) As Boolean
		For X As Int32 = 0 To PlayerRelUB
			If mlPlayerRelIdx(X) = lPlayerID Then
				Return True
			End If
		Next X
		Return False
	End Function

	Public Function GetPlayerRelScore(ByVal lPlayerID As Int32, ByVal bAddIfNotFound As Boolean, ByVal lFirstContactEntityServerIdx As Int32) As Byte

		If Me.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID OrElse lPlayerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
			If Me.ObjectID = lPlayerID Then Return elRelTypes.eNeutral Else Return elRelTypes.eWar
		End If

		For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) = lPlayerID Then

                'If moPlayerRels(X).WithThisScore > elRelTypes.eWar Then
                '    Dim yRel As Byte = moPlayerRels(X).oThisPlayer.GetPlayerRelScore(Me.ObjectID, False, -1)
                '    If yRel <= elRelTypes.eWar Then Return yRel
                'End If

                Return Math.Min(moPlayerRels(X).WithThisScore, moPlayerRels(X).OtherTowardsMe)
            End If
		Next X

		'Ok, if we are here, then no relationship exists... we create one automatically
		If bAddIfNotFound = True Then
			Dim oPlayer As Player = Nothing
			For X As Int32 = 0 To glPlayerUB
				If glPlayerIdx(X) = lPlayerID Then
					oPlayer = goPlayers(X)
					Exit For
				End If
			Next X
			If oPlayer Is Nothing = False Then
				'Create the relationship object for both parties (default is neutral)... we do this to ensure this test only
				'  occurs once on this server...
				Dim oRel As PlayerRel = New PlayerRel()
				oRel.oPlayerRegards = Me
				oRel.oThisPlayer = oPlayer
                oRel.WithThisScore = elRelTypes.eNeutral
                oRel.OtherTowardsMe = elRelTypes.eNeutral
				Me.SetPlayerRel(lPlayerID, oRel)
				oRel = Nothing

				oRel = New PlayerRel
				oRel.oPlayerRegards = oPlayer
				oRel.oThisPlayer = Me
                oRel.WithThisScore = elRelTypes.eNeutral
                oRel.OtherTowardsMe = elRelTypes.eNeutral
				oPlayer.SetPlayerRel(Me.ObjectID, oRel)
				oRel = Nothing

				'Now, send a message to the primary server that the new relationship is established, it will handle the rest
				Dim yMsg(30) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yMsg, 0)
				System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
				System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, 6)
				yMsg(10) = elRelTypes.eNeutral

				If lFirstContactEntityServerIdx <> -1 Then
					Try
						With goEntity(lFirstContactEntityServerIdx)
							.ParentEnvir.GetGUIDAsString.CopyTo(yMsg, 11)
							.GetGUIDAsString.CopyTo(yMsg, 17)
							System.BitConverter.GetBytes(.LocX).CopyTo(yMsg, 23)
							System.BitConverter.GetBytes(.LocZ).CopyTo(yMsg, 27)
						End With
					Catch
						'do nothing
					End Try
				End If
				goMsgSys.SendToPrimary(yMsg)
			End If
		End If
		Return elRelTypes.eNeutral
	End Function

    Public Function GetPlayerRel(ByVal lPlayerID As Int32) As PlayerRel
        For X As Int32 = 0 To PlayerRelUB
            If mlPlayerRelIdx(X) = lPlayerID Then
                Return moPlayerRels(X)
            End If
        Next X
        Return Nothing
    End Function

    'Public Function PlayerHasWar() As Boolean
    '    For X As Int32 = 0 To PlayerRelUB
    '        If mlPlayerRelIdx(X) <> -1 AndAlso moPlayerRels(X).WithThisScore <= elRelTypes.eWar Then Return True
    '    Next X
    'End Function

	Public Sub RemovePlayerRel(ByVal lWithPlayerID As Int32, ByVal bNoRecursion As Boolean)
		Try
			Dim lShiftIdx As Int32 = -1
			'Remove all relationships
			For X As Int32 = 0 To PlayerRelUB
				If mlPlayerRelIdx(X) <> -1 AndAlso (lWithPlayerID = -1 OrElse lWithPlayerID = mlPlayerRelIdx(X)) Then
					If moPlayerRels(X).oThisPlayer Is Nothing = False AndAlso moPlayerRels(X).oPlayerRegards Is Nothing = False Then
						If moPlayerRels(X).oThisPlayer.ObjectID = Me.ObjectID Then
							'ok, wrong direction
							If bNoRecursion = False Then moPlayerRels(X).oPlayerRegards.RemovePlayerRel(Me.ObjectID, True)

							Dim lOtherPlayerID As Int32 = moPlayerRels(X).oPlayerRegards.ObjectID

							For Y As Int32 = 0 To glEnvirUB
								For Z As Int32 = 0 To goEnvirs(Y).lPlayersWhoHaveUnitsHereUB
									If goEnvirs(Y).lPlayersWhoHaveUnitsHereIdx(Z) = Me.ObjectID OrElse goEnvirs(Y).lPlayersWhoHaveUnitsHereIdx(Z) = lOtherPlayerID Then
										goEnvirs(Y).TestForAggression()
										Exit For
									End If
								Next Z
							Next Y
						Else
							If bNoRecursion = False Then moPlayerRels(X).oThisPlayer.RemovePlayerRel(Me.ObjectID, True)

							Dim lOtherPlayerID As Int32 = moPlayerRels(X).oThisPlayer.ObjectID

							For Y As Int32 = 0 To glEnvirUB
								For Z As Int32 = 0 To goEnvirs(Y).lPlayersWhoHaveUnitsHereUB
									If goEnvirs(Y).lPlayersWhoHaveUnitsHereIdx(Z) = Me.ObjectID OrElse goEnvirs(Y).lPlayersWhoHaveUnitsHereIdx(Z) = lOtherPlayerID Then
										goEnvirs(Y).TestForAggression()
										Exit For
									End If
								Next Z
							Next Y
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
			Else
				For X As Int32 = lShiftIdx To PlayerRelUB - 1
					moPlayerRels(X) = moPlayerRels(X + 1)
					mlPlayerRelIdx(X) = mlPlayerRelIdx(X + 1)
				Next X
				PlayerRelUB -= 1
			End If
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("RemovePlayerRel.Err: " & ex.Message)
		End Try
	End Sub
End Class

Public Class PlayerRel
    Public oPlayerRegards As Player
    Public oThisPlayer As Player
    Public WithThisScore As Byte
    Public OtherTowardsMe As Byte
End Class

Public Structure EnvirAlert
    'This is the minimum time between alerting the PRIMARY of an under attack alert in this environment
    Private Const ml_UNDER_ATTACK_ALERT_THRESHOLD As Int32 = 3600       '2 minutes... 
    'This is the minimum time between alerting the PRIMARY of an engage alert in this environment
    Private Const ml_ENGAGE_ALERT_THRESHOLD As Int32 = 3600             '2 minutes...

    Public lEnvirID As Int32
    Public iEnvirTypeID As Int16
    Public yAlertType As Byte
    Public lLastAlertUpdate As Int32

    Public LocX As Int32
    Public LocZ As Int32

    Public Function CheckForAlert() As Boolean
        If glCurrentCycle - lLastAlertUpdate > ml_UNDER_ATTACK_ALERT_THRESHOLD Then
            lLastAlertUpdate = glCurrentCycle
            Return True
        End If
        'Regardless, we reset the alert time... this causes a very intelligent approach to the alert system
        lLastAlertUpdate = glCurrentCycle
        Return False
    End Function
End Structure