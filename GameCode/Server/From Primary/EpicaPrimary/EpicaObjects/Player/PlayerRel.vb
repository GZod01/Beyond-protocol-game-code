Public Class PlayerRel
    Public oPlayerRegards As Player
    Public oThisPlayer As Player
    Public WithThisScore As Byte

    Public TargetScore As Byte
    Public lCycleOfNextScoreUpdate As Int32 = -1

    Public blTotalWarpointsGained As Int64 = 0      'the total number of warpoints gained from oPlayerRegards killing oThisPlayer
    Public lPlayersWPV As Int32 = 10000             'the Warpoint Vitality for items that oPlayerRegards kills of oThisPlayer
    Public bForceSaveMe As Boolean = False

    Private mlLastWPVRegeneration As Int32 = Int32.MinValue        'cycle from the last time the WPV was regenerated... this value does not persist

    Public Sub RegenerateWPV()
        If mlLastWPVRegeneration <> Int32.MinValue AndAlso lPlayersWPV <> 10000 Then
            'Get number of cycles passed
            Dim lCyclesPassed As Int32 = glCurrentCycle - mlLastWPVRegeneration
            'Now, there are 30 cycles in a second... 60 seconds in a minute
            lCyclesPassed = lCyclesPassed \ 1800
            lPlayersWPV += lCyclesPassed
            If lPlayersWPV > 10000 Then lPlayersWPV = 10000
            bForceSaveMe = True
        End If
        mlLastWPVRegeneration = glCurrentCycle
    End Sub

	Public Function GetObjAsString() As Byte()
		'here we will return the entire object as a string
        Dim yMsg(13) As Byte           '0 to 8 = 9 bytes
		System.BitConverter.GetBytes(oPlayerRegards.ObjectID).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(oThisPlayer.ObjectID).CopyTo(yMsg, 4)
		yMsg(8) = WithThisScore

        yMsg(9) = TargetScore
        Dim lCycles As Int32 = -1
        If lCycleOfNextScoreUpdate > glCurrentCycle Then lCycles = lCycleOfNextScoreUpdate - glCurrentCycle
        System.BitConverter.GetBytes(lCycles).CopyTo(yMsg, 10)

		Return yMsg
	End Function

	Public Function GetSaveObjectText() As String
        Dim sSQL As String = "DELETE FROM tblPlayerToPlayerRel WHERE Player1ID = " & oPlayerRegards.ObjectID & " AND Player2ID = " & oThisPlayer.ObjectID & vbCrLf

        Dim lCycles As Int32 = -1
        If lCycleOfNextScoreUpdate > glCurrentCycle Then
            lCycles = lCycleOfNextScoreUpdate - glCurrentCycle
        End If

        sSQL &= "INSERT INTO tblPlayerToPlayerRel (Player1ID, Player2ID, RelTypeID, TargetScore, CyclesToNextUpdate, TotalWPGain, WPV) VALUES (" & _
            oPlayerRegards.ObjectID & ", " & oThisPlayer.ObjectID & ", " & WithThisScore & ", " & TargetScore & ", " & lCycles & ", " & blTotalWarpointsGained & ", " & lPlayersWPV & ")"
		Return sSQL
	End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            bForceSaveMe = False
            sSQL = "DELETE FROM tblPlayerToPlayerRel WHERE Player1ID = " & oPlayerRegards.ObjectID & " AND Player2ID = " & oThisPlayer.ObjectID
            oComm = Nothing
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm = Nothing

            Dim lCycles As Int32 = -1
            If lCycleOfNextScoreUpdate > glCurrentCycle Then
                lCycles = lCycleOfNextScoreUpdate - glCurrentCycle
            End If

            sSQL = "INSERT INTO tblPlayerToPlayerRel (Player1ID, Player2ID, RelTypeID, TargetScore, CyclesToNextUpdate, TotalWPGain, WPV) VALUES (" & _
              oPlayerRegards.ObjectID & ", " & oThisPlayer.ObjectID & ", " & WithThisScore & ", " & TargetScore & ", " & lCycles & ", " & blTotalWarpointsGained & ", " & lPlayersWPV & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object Player to Player Rel. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub StorePlayerRelHistory()
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand


        Try
            Dim lDTStamp As Int32 = GetDateAsNumber(Now)

            sSQL = "UPDATE tblP2PRelHist SET RelTypeID = " & WithThisScore & " WHERE Player1ID = " & oPlayerRegards.ObjectID & " AND Player2ID = " & _
                oThisPlayer.ObjectID & " and DateTimeStamp = " & lDTStamp
            oComm = Nothing
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                sSQL = "INSERT INTO tblP2PRelHist (Player1ID, Player2ID, DateTimeStamp, RelTypeID) VALUES (" & oPlayerRegards.ObjectID & ", " & _
                    oThisPlayer.ObjectID & ", " & lDTStamp & ", " & WithThisScore & ")"
                oComm = Nothing
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then Throw New Exception("No records affected when saving StorePlayerRelHistory!")
            End If
            oComm = Nothing
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to StorePlayerRelHistory. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try

    End Sub

End Class

Public Class CascadeWarDec
    Private Structure uPlayerSide
        Public oPlayer As Player        'this player
		Public ySide As Byte			'0 indicates undecided, 1 indicates Side A, 2 indicates Side B
		Public oSupportingPlayer As Player	'who this player supports
		Public yHighestScore As Byte
    End Structure
    Private muSides() As uPlayerSide
    Private mlSideUB As Int32 = -1

	Private Sub AddPlayerSides(ByRef oPlayer As Player, ByVal lTier As Int32)
		For X As Int32 = 0 To oPlayer.PlayerRelUB
			Dim oTmpRel As PlayerRel = oPlayer.GetPlayerRelByIndex(X)
			Dim bFound As Boolean = False
			If oTmpRel Is Nothing = False Then
				If oTmpRel.oThisPlayer.GetPlayerRelScore(oTmpRel.oPlayerRegards.ObjectID) >= elRelTypes.ePeace Then
					For Y As Int32 = 0 To mlSideUB
						If muSides(Y).oPlayer.ObjectID = oTmpRel.oThisPlayer.ObjectID Then
							bFound = True
							Exit For
						End If
					Next Y
					If bFound = False Then
						mlSideUB += 1
						ReDim Preserve muSides(mlSideUB)
						With muSides(mlSideUB)
							.oPlayer = oTmpRel.oThisPlayer
							.ySide = 0
						End With
						If lTier < 3 Then AddPlayerSides(oTmpRel.oThisPlayer, lTier + 1)
					End If
				End If
			End If
		Next X
	End Sub

    Public Sub HandleCascadeWarDec(ByRef oInitiator As Player, ByRef oTarget As Player, ByVal yThisScore As Byte)
        mlSideUB = -1
        ReDim muSides(-1)

        'Ok, add all of our PlayerSides

        mlSideUB += 1
        ReDim Preserve muSides(mlSideUB)
        With muSides(mlSideUB)
            .oPlayer = oInitiator
            .ySide = 1
        End With
        mlSideUB += 1
        ReDim Preserve muSides(mlSideUB)
        With muSides(mlSideUB)
            .oPlayer = oTarget
            .ySide = 2
        End With

		'AddPlayerSides(oInitiator, 1)
		'AddPlayerSides(oTarget, 1)

		'      'Ok, now, loop until we got everyone
		'      Dim bDone As Boolean = False
		'      Dim bNoChange As Boolean = False
		'      Dim lNoChangeCnt As Int32 = 0

		'      Dim bForceDecisionPass1 As Boolean = False      'If a player has a relationship with a player involved in the initial war dec, then they will choose that side (if they have both players then it chooses the highest between the two. If they are both the same, it is an indecisive)
		'      Dim bForceDecisionPass2 As Boolean = False      'brute force decision making

		'      Dim lCntA As Int32 = 0
		'      Dim lCntB As Int32 = 0
		'      Dim bEarlyExit As Boolean = False

		'      While bDone = False
		'          bDone = True
		'          bNoChange = True
		'          bEarlyExit = False

		'          For X As Int32 = 0 To mlSideUB
		'              'Do we know this one yet?
		'              If muSides(X).ySide = 0 Then
		'                  'No, ok, bdone = false
		'                  bDone = False

		'			Dim yPreferredSide As Byte = 0
		'			Dim yHighestScore As Byte = elRelTypes.ePeace
		'			'Now, loop through this player's relationship list until you find a side
		'                  For Y As Int32 = 0 To muSides(X).oPlayer.PlayerRelUB
		'                      If X = Y Then Continue For

		'				'get the relationship by index...
		'				Dim oTmpRel As PlayerRel = muSides(X).oPlayer.GetPlayerRelByIndex(Y)
		'				'Is this player allied?
		'                      If oTmpRel.WithThisScore >= yHighestScore Then
		'					'Ok, now, find the playerside for that player
		'					Dim bUpdateSupporting As Boolean = False
		'                          For Z As Int32 = 0 To mlSideUB
		'                              If muSides(Z).oPlayer.ObjectID = oTmpRel.oThisPlayer.ObjectID Then
		'                                  'ok, has this player made a choice?
		'                                  If muSides(Z).ySide <> 0 Then
		'                                      'Yes, they have... ok, check for a special situation... do I have a preferred side already?
		'                                      'and is the relscore the same as the current highest score?
		'                                      If yPreferredSide <> 0 AndAlso yPreferredSide <> muSides(Z).ySide AndAlso oTmpRel.WithThisScore = yHighestScore Then
		'                                          'Ok, does this player have a relationship with the initiator or target?
		'                                          Dim yRelA As Byte = muSides(X).oPlayer.GetPlayerRelScore(oInitiator.ObjectID)
		'                                          Dim yRelB As Byte = muSides(X).oPlayer.GetPlayerRelScore(oTarget.ObjectID)

		'                                          If yRelA > yRelB Then
		'                                              'Ok, player will prefer to be on Side A's side
		'                                              yPreferredSide = 1
		'                                          ElseIf yRelB > yRelA Then
		'                                              'Ok, player will prefer to be on Side B's side
		'                                              yPreferredSide = 2
		'									End If
		'									bUpdateSupporting = False
		'                                      Else
		'                                          'Alright, just use that side's chosen side
		'									yPreferredSide = muSides(Z).ySide
		'									bUpdateSupporting = True
		'                                      End If
		'                                  ElseIf oTmpRel.WithThisScore >= yHighestScore Then
		'                                      If bForceDecisionPass1 = True Then
		'                                          'Ok, in this pass, we check if the current player has a relationship with one or both sides
		'                                          Dim yRelA As Byte = muSides(Z).oPlayer.GetPlayerRelScore(oInitiator.ObjectID)
		'                                          Dim yRelB As Byte = muSides(Z).oPlayer.GetPlayerRelScore(oTarget.ObjectID)

		'                                          If yRelA > yRelB Then
		'                                              'Ok, player will prefer to be on Side A's side
		'                                              yPreferredSide = 1
		'                                          ElseIf yRelB > yRelA Then
		'                                              'Ok, player will prefer to be on Side B's side
		'                                              yPreferredSide = 2
		'									End If
		'									bUpdateSupporting = False
		'                                          bEarlyExit = True       'restart our while loop when we are done
		'                                      ElseIf bForceDecisionPass2 = True Then
		'                                          'ok, in this side, we do a quick count
		'                                          If lCntB > lCntA Then
		'                                              'side B is already overloaded, so push this player to side A
		'                                              yPreferredSide = 1
		'                                          Else : yPreferredSide = 2
		'									End If
		'									bUpdateSupporting = True
		'                                          bEarlyExit = True
		'                                      Else
		'									yPreferredSide = 0
		'									bUpdateSupporting = False
		'                                      End If
		'                                  End If
		'                                  Exit For
		'                              End If
		'                          Next Z
		'					yHighestScore = oTmpRel.WithThisScore
		'					If bUpdateSupporting = True Then
		'						muSides(X).oSupportingPlayer = oTmpRel.oThisPlayer
		'						muSides(X).yHighestScore = yHighestScore
		'					End If
		'				End If
		'			Next Y
		'                  'Now, what is preferredside?
		'                  If yPreferredSide <> 0 Then
		'                      bNoChange = False
		'                      muSides(X).ySide = yPreferredSide
		'                      If yPreferredSide = 1 Then lCntA += 1 Else lCntB += 1
		'                      lNoChangeCnt = 0

		'                      If bEarlyExit = True Then Exit For
		'                  End If
		'              End If
		'          Next X

		'          bForceDecisionPass1 = False
		'          bForceDecisionPass2 = False
		'          If bNoChange = True Then
		'              lNoChangeCnt += 1
		'              If lNoChangeCnt = 3 Then
		'                  bForceDecisionPass1 = True
		'              ElseIf lNoChangeCnt = 4 Then
		'                  bForceDecisionPass2 = True
		'              ElseIf lNoChangeCnt = 5 Then
		'                  bDone = True
		'              End If
		'          End If
		'      End While

		'      'Ok, here, we check for a final time if there were any undecided's and force them to choose a side
		'      For X As Int32 = 0 To mlSideUB
		'          If muSides(X).ySide = 0 Then
		'              If lCntB > lCntA Then
		'                  muSides(X).ySide = 1
		'                  lCntA += 1
		'              Else
		'                  muSides(X).ySide = 2
		'                  lCntB += 1
		'              End If
		'          End If
		'      Next X

		'create our email body msg
		Dim oSB As New System.Text.StringBuilder
		oSB.AppendLine(oInitiator.sPlayerNameProper & " initiated war on " & oTarget.sPlayerNameProper & "!")
		oSB.AppendLine()

        'Now, do all of the alerts 
        Dim yData(10) As Byte
		For X As Int32 = 0 To mlSideUB
			If muSides(X).oPlayer.ObjectID <> oInitiator.ObjectID AndAlso muSides(X).oPlayer.ObjectID <> oTarget.ObjectID Then
				If muSides(X).ySide = 1 Then
					If muSides(X).oSupportingPlayer Is Nothing = False Then
						oSB.AppendLine(muSides(X).oPlayer.sPlayerNameProper & " joins the war on the side of " & oInitiator.sPlayerNameProper & " to honor their alliance with " & muSides(X).oSupportingPlayer.sPlayerNameProper & " (" & muSides(X).yHighestScore & ")")
					Else : oSB.AppendLine(muSides(X).oPlayer.sPlayerNameProper & " joins on the side of " & oInitiator.sPlayerNameProper)
					End If
				Else
					If muSides(X).oSupportingPlayer Is Nothing = False Then
						oSB.AppendLine(muSides(X).oPlayer.sPlayerNameProper & " joins the war on the side of " & oTarget.sPlayerNameProper & " to honor their alliance with " & muSides(X).oSupportingPlayer.sPlayerNameProper & " (" & muSides(X).yHighestScore & ")")
					Else : oSB.AppendLine(muSides(X).oPlayer.sPlayerNameProper & " joins on the side of " & oTarget.sPlayerNameProper)
					End If
				End If
			End If

			For Y As Int32 = 0 To mlSideUB
				If X <> Y Then
                    If muSides(X).ySide <> muSides(Y).ySide Then
                        Dim yResp(27) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yResp, 0)
                        muSides(Y).oPlayer.GetGUIDAsString.CopyTo(yResp, 2)
                        muSides(Y).oPlayer.PlayerName.CopyTo(yResp, 8)
                        muSides(X).oPlayer.SendPlayerMessage(yResp, False, 0)
                        muSides(X).oPlayer.GetGUIDAsString.CopyTo(yResp, 2)
                        muSides(X).oPlayer.PlayerName.CopyTo(yResp, 8)
                        muSides(Y).oPlayer.SendPlayerMessage(yResp, False, 0)

                        Dim oXYRel As New PlayerRel()
                        oXYRel.oPlayerRegards = muSides(X).oPlayer
                        oXYRel.oThisPlayer = muSides(Y).oPlayer
                        oXYRel.WithThisScore = yThisScore
                        muSides(X).oPlayer.SetPlayerRel(muSides(Y).oPlayer.ObjectID, oXYRel, True)

                        System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(oXYRel.oPlayerRegards.ObjectID).CopyTo(yData, 2)
                        System.BitConverter.GetBytes(oXYRel.oThisPlayer.ObjectID).CopyTo(yData, 6)
                        yData(10) = oXYRel.WithThisScore
                        muSides(Y).oPlayer.SendPlayerMessage(yData, False, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)
                        goMsgSys.BroadcastToDomains(yData)

                        Dim oYXRel As New PlayerRel()
                        oYXRel.oPlayerRegards = muSides(Y).oPlayer
                        oYXRel.oThisPlayer = muSides(X).oPlayer
                        oYXRel.WithThisScore = yThisScore
                        muSides(Y).oPlayer.SetPlayerRel(muSides(X).oPlayer.ObjectID, oYXRel, True)

                        System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yData, 0)
                        System.BitConverter.GetBytes(oYXRel.oPlayerRegards.ObjectID).CopyTo(yData, 2)
                        System.BitConverter.GetBytes(oYXRel.oThisPlayer.ObjectID).CopyTo(yData, 6)
                        yData(10) = oYXRel.WithThisScore
                        muSides(X).oPlayer.SendPlayerMessage(yData, False, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)
                        goMsgSys.BroadcastToDomains(yData)
                    End If
				End If
			Next Y
		Next X

		Dim sText As String = oSB.ToString
		Dim oPC As PlayerComm = Nothing
        If (oInitiator.iInternalEmailSettings And eEmailSettings.ePlayerRels) <> 0 Then oPC = oInitiator.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sText, "War Declaration", oInitiator.ObjectID, GetDateAsNumber(Now), False, oInitiator.sPlayerNameProper, Nothing)
		If oPC Is Nothing = False Then
			oInitiator.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
		End If
		oPC = Nothing
        If (oTarget.iInternalEmailSettings And eEmailSettings.ePlayerRels) <> 0 Then oPC = oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sText, "War Declaration", oTarget.ObjectID, GetDateAsNumber(Now), False, oTarget.sPlayerNameProper, Nothing)
		If oPC Is Nothing = False Then
			oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
		End If

		For X As Int32 = 0 To mlSideUB
			If muSides(X).oPlayer.ObjectID <> oInitiator.ObjectID AndAlso muSides(X).oPlayer.ObjectID <> oTarget.ObjectID Then
                oPC = muSides(X).oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sText, "War Declaration", muSides(X).oPlayer.ObjectID, GetDateAsNumber(Now), False, muSides(X).oPlayer.sPlayerNameProper, Nothing)
				If oPC Is Nothing = False Then
					muSides(X).oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
				End If
			End If
		Next X

		'Dim bBadWarDec As Boolean = True
		'If oTarget.oSocket Is Nothing = False Then
		'	If oTarget.oSocket.IsConnected = True Then
		'		bBadWarDec = False
		'	Else
		'		bBadWarDec = Not oTarget.HasOnlineAliases(AliasingRights.eMoveUnits)
		'	End If
		'Else
		'	If oTarget.HasOnlineAliases(AliasingRights.eMoveUnits) = True Then
		'		bBadWarDec = False
		'	End If
		'End If
		Dim bUpdateCP As Boolean = False
		'Dim bUpdateMorale As Boolean = False
		'If bBadWarDec = True Then
		'	'K, the initiator receives 
		'	'-70 morale to all colonies for the next 48 hours
		'	'+5 CP costs for every unit for 24 hours
		'	bUpdateCP = True
		'	bUpdateMorale = True

		'	'both values are cumulative and will reset the time for each value each time it happens
		'	oInitiator.BadWarDecMoralePenalty -= 70
		'	oInitiator.BadWarDecMoralePenaltyEndCycle = glCurrentCycle + 5184000		'48 hours
		'	oInitiator.BadWarDecCPIncrease += 5
		'	oInitiator.BadWarDecCPIncreaseEndCycle = glCurrentCycle + 2592000			'24 hours

		'	'Finally, inform GNS...
		'	Dim yMsg(28) As Byte
		'	System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yMsg, 0)
		'	yMsg(2) = NewsItemType.eOfflineWarDec
		'	goGalaxy(0).GetGUIDAsString.CopyTo(yMsg, 3)
		'	oInitiator.PlayerName.CopyTo(yMsg, 9)
		'	goMsgSys.SendToEmailSrvr(yMsg)
		'End If

        'Ok, now, determine if the player being declared war on is less rank
        Dim yTargetTestTitle As Byte = oTarget.yPlayerTitle
        Dim yPlayerTestTitle As Byte = oInitiator.yPlayerTitle

        If oInitiator.oGuild Is Nothing = False Then
            yPlayerTestTitle = Math.Max(yPlayerTestTitle, oInitiator.oGuild.GetHighestPlayerTitle())
        End If
        If oTarget.oGuild Is Nothing = False Then
            yTargetTestTitle = Math.Max(yTargetTestTitle, oTarget.oGuild.GetHighestPlayerTitle())
        End If

        If (yTargetTestTitle And Player.PlayerRank.ExRankShift) <> 0 Then yTargetTestTitle = yTargetTestTitle Xor Player.PlayerRank.ExRankShift
        If (yPlayerTestTitle And Player.PlayerRank.ExRankShift) <> 0 Then yPlayerTestTitle = yPlayerTestTitle Xor Player.PlayerRank.ExRankShift
        If yTargetTestTitle < yPlayerTestTitle OrElse (oInitiator.yPlayerPhase = eyPlayerPhase.eSecondPhase AndAlso oInitiator.TotalPlayTime > 86400) Then
            'ok, it is...
            Dim lCPDiff As Int32 = CInt(yPlayerTestTitle) - CInt(yTargetTestTitle)
            If oInitiator.yPlayerPhase = eyPlayerPhase.eSecondPhase Then
                If lCPDiff < 0 Then lCPDiff = 0
                lCPDiff += (oInitiator.TotalPlayTime \ 86400)
                If lCPDiff > 20 Then lCPDiff = 20 'just to be safe
            End If

            If lCPDiff <> 0 Then
                Dim oNewCP As New PlayerCPPenalty()
                With oNewCP
                    .oPlayer = oInitiator
                    .oDecPlayer = oTarget
                    .yPenalty = CByte(Math.Abs(lCPDiff))
                    .SaveObject()
                End With
                oInitiator.AddCPPenalty(oNewCP)
                oInitiator.SendCPPenaltyList()
            End If

            oInitiator.BadWarDecCPIncrease += lCPDiff
            If gb_IS_TEST_SERVER = True Then
                oInitiator.BadWarDecCPIncreaseEndCycle = glCurrentCycle + 9000
            Else
                oInitiator.BadWarDecCPIncreaseEndCycle = glCurrentCycle + 2592000           '24 hours
            End If


            oPC = oInitiator.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "You have incurred a Command Point penalty of " & lCPDiff & " for declaring war on a player/guild with a rank less than your rank or the highest rank of player in your guild. The duration of this penalty is 24 hours. If you already had a penalty in effect, that penalty is also increased to the new duration of 24 hours.", "Command Point Penalty", oInitiator.ObjectID, GetDateAsNumber(Now), False, oInitiator.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then
                oInitiator.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If

            bUpdateCP = True
        End If

        If bUpdateCP = True Then
            'Now, update all region servers with this data...
            Dim yMsg(11) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(oInitiator.ObjectID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty).CopyTo(yMsg, 6)
            System.BitConverter.GetBytes(oInitiator.BadWarDecCPIncrease).CopyTo(yMsg, 8)
            'goMsgSys.BroadcastToDomains(yMsg)
            goMsgSys.SendMsgToOperator(yMsg)
            'oInitiator.SendPlayerMessage(yMsg, False, AliasingRights.eViewUnitsAndFacilities)
        End If

        'bUpdateCP = False
        'If (oTarget.yPlayerPhase = eyPlayerPhase.eSecondPhase AndAlso oTarget.TotalPlayTime > 86400) Then
        '    'ok, it is...
        '    Dim lCPDiff As Int32 = (oTarget.yPlayerPhase \ 86400)
        '    If lCPDiff > 20 Then lCPDiff = 20 'just to be safe

        '    If lCPDiff <> 0 Then
        '        Dim oNewCP As New PlayerCPPenalty()
        '        With oNewCP
        '            .oPlayer = oTarget
        '            .oDecPlayer = oInitiator
        '            .yPenalty = CByte(Math.Abs(lCPDiff))
        '            .SaveObject()
        '        End With
        '        oTarget.AddCPPenalty(oNewCP)
        '        oTarget.SendCPPenaltyList()
        '    End If

        '    oTarget.BadWarDecCPIncrease += lCPDiff
        '    If gb_IS_TEST_SERVER = True Then
        '        oTarget.BadWarDecCPIncreaseEndCycle = glCurrentCycle + 9000
        '    Else
        '        oTarget.BadWarDecCPIncreaseEndCycle = glCurrentCycle + 2592000           '24 hours
        '    End If


        '    oPC = oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "You have incurred a Command Point penalty of " & lCPDiff & " for declaring war on a player/guild with a rank less than your rank or the highest rank of player in your guild. The duration of this penalty is 24 hours. If you already had a penalty in effect, that penalty is also increased to the new duration of 24 hours.", "Command Point Penalty", oTarget.ObjectID, GetDateAsNumber(Now), False, oTarget.sPlayerNameProper, Nothing)
        '    If oPC Is Nothing = False Then
        '        oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
        '    End If

        '    bUpdateCP = True
        'End If

        'If bUpdateCP = True Then
        '    'Now, update all region servers with this data...
        '    Dim yMsg(11) As Byte
        '    System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
        '    System.BitConverter.GetBytes(oTarget.ObjectID).CopyTo(yMsg, 2)
        '    System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty).CopyTo(yMsg, 6)
        '    System.BitConverter.GetBytes(oTarget.BadWarDecCPIncrease).CopyTo(yMsg, 8)
        '    'goMsgSys.BroadcastToDomains(yMsg)
        '    goMsgSys.SendMsgToOperator(yMsg)
        '    'oInitiator.SendPlayerMessage(yMsg, False, AliasingRights.eViewUnitsAndFacilities)
        'End If

        'If bUpdateMorale = True Then
        '    Dim yMsg(11) As Byte
        '    System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
        '    System.BitConverter.GetBytes(oInitiator.ObjectID).CopyTo(yMsg, 2)
        '    System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecMoralePenalty).CopyTo(yMsg, 6)
        '    System.BitConverter.GetBytes(oInitiator.BadWarDecMoralePenalty).CopyTo(yMsg, 8)
        '    oInitiator.SendPlayerMessage(yMsg, False, AliasingRights.eViewColonyStats)
        'End If

	End Sub

End Class