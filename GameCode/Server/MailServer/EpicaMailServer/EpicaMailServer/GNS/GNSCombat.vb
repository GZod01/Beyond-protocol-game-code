Public Class GNSCombat
	Public lStoryEnvirID As Int32
	Public iStoryEnvirTypeID As Int16

	Public lCombatLocationID As Int32
	Public iCombatLocationTypeID As Int16

	Private muData() As GNSCombatData
	Private mlDataPlayer() As Int32
	Private mlDataUB As Int32 = -1

	Private myLatestType As Byte

	Public Sub AddGNSCombatData(ByVal uData As GNSCombatData, ByVal yType As Byte)
		myLatestType = yType
		For X As Int32 = 0 To mlDataUB
			If mlDataPlayer(X) = uData.lPlayerID Then
				muData(X) = uData
				Return
			End If
		Next X
		SyncLock Me
			mlDataUB += 1
			ReDim Preserve mlDataPlayer(mlDataUB)
			ReDim Preserve muData(mlDataUB)
			mlDataPlayer(mlDataUB) = uData.lPlayerID
			muData(mlDataUB) = uData
		End SyncLock

		CheckForNewStory()
	End Sub

	'ok, now, we have a storage of data....
	'each data has a playerID that represents that player's side of the story
	'  each of those has a Killed By and a Killed list
	'  based on those lists, I can determine combat/war...
	'for example, if A was killed by B, C, D and Killed B, C, D, E then
	'  we can determine that player A is at war with forces belonging to BCDE

	'Now, we receive a new item... F was killed by B,C,D and Killed B,C,D
	'  we can determine that F and A are allied (not strictly, simply against the same forces)
	'
	'We get E and it says, E was killed by A and B and killed A and B
	'  because both sides are at conflict with each other, we can determine that E is a separate story.

	'However, we can make a separate story type... Pandemonium or something and make it work by allowing it to state simply
	'  all forces involved in fighting without naming sides

	Private Sub CheckForNewStory()
		Dim bDone As Boolean = False
		While bDone = False
			bDone = True

			For X As Int32 = 0 To mlDataUB
				If mlDataPlayer(X) <> -1 Then
					Dim bGood As Boolean = True

					Try
						'ok, now, see if my killed by's have put in a data object
						For Y As Int32 = 0 To muData(X).lKilledByUB
							Dim bFound As Boolean = False
							For Z As Int32 = 0 To mlDataUB
								If mlDataPlayer(Z) = muData(X).uKilledBy(Y).lPlayerID Then
									bFound = True

									For R As Int32 = 0 To muData(Z).lKilledByUB
										Dim bNewFound As Boolean = False
										For Q As Int32 = 0 To mlDataUB
											If mlDataPlayer(Q) = muData(Z).uKilledBy(R).lPlayerID Then
												bNewFound = True
												Exit For
											End If
										Next Q
										If bNewFound = False Then
											bFound = False
											Exit For
										End If
									Next R
									If bFound = False Then Exit For
									For R As Int32 = 0 To muData(Z).lKilledUB
										Dim bNewFound As Boolean = False
										For Q As Int32 = 0 To mlDataUB
											If mlDataPlayer(Q) = muData(Z).uKilled(R).lPlayerID Then
												bNewFound = True
												Exit For
											End If
										Next Q
										If bNewFound = False Then
											bFound = False
											Exit For
										End If
									Next R

									Exit For
								End If
							Next Z
							If bFound = False Then
								bGood = False
								Exit For
							End If
						Next Y
						If bGood = False Then Continue For
					Catch
						Continue For
					End Try

					Try
						'Ok, now, see if my killed have checkde in
						For Y As Int32 = 0 To muData(X).lKilledUB
							Dim bFound As Boolean = False
							For Z As Int32 = 0 To mlDataUB
								If mlDataPlayer(Z) = muData(X).uKilled(Y).lPlayerID Then
									bFound = True

									For R As Int32 = 0 To muData(Z).lKilledByUB
										Dim bNewFound As Boolean = False
										For Q As Int32 = 0 To mlDataUB
											If mlDataPlayer(Q) = muData(Z).uKilledBy(R).lPlayerID Then
												bNewFound = True
												Exit For
											End If
										Next Q
										If bNewFound = False Then
											bFound = False
											Exit For
										End If
									Next R
									If bFound = False Then Exit For
									For R As Int32 = 0 To muData(Z).lKilledUB
										Dim bNewFound As Boolean = False
										For Q As Int32 = 0 To mlDataUB
											If mlDataPlayer(Q) = muData(Z).uKilled(R).lPlayerID Then
												bNewFound = True
												Exit For
											End If
										Next Q
										If bNewFound = False Then
											bFound = False
											Exit For
										End If
									Next R
									Exit For
								End If
							Next Z
							If bFound = False Then
								bGood = False
								Exit For
							End If
						Next Y
						If bGood = False Then Continue For
					Catch
						Continue For
					End Try

					'Ok, all players have checked in... let's generate our news story
					'now, generate and send our story... this will cause mlDataPlayer() variables to clear so we set our
					'bdone = false so we can go back through the list (if needed)
					GenerateAndSendStory(X)
					bDone = False		'do this because we will need to go through the loop again when we are done...
					Exit For
				End If
			Next X
		End While
	End Sub

	Private Sub GenerateAndSendStory(ByVal lDataIdx As Int32)
		'Ok, lDataIdx is the muData() array index for the item that we are gonna generate news for...
		'  we have to look at all values associated to the item...
		Dim uData As GNSCombatData = muData(lDataIdx)

		'Ok, here we go...
		Dim oData As New GNSData()
		With oData
			.LocationID = uData.lFightLocID
			.LocationTypeID = uData.iFightLocTypeID
			.EndOfBattle = uData.lEnd
			.StartOfBattle = uData.lStart

			If .LocationTypeID = ObjectType.ePlanet Then
				.PlanetName = uData.sLocName
			Else : .SystemName = uData.sLocName
			End If
 

			'Ok, now, side A... get uData's list... and allies?
			'Now, side B... get uData's Killed list... and allies?

			'Add me to player Side A
			'  add those I killed to player side B
			Dim lPlayerID(-1) As Int32
			Dim ySide(-1) As Byte
			Dim lPS_UB As Int32 = -1
			Dim bConflict As Boolean = False

			.PlayerSideListA = ""
			.PlayerSideListB = ""

			DetermineSide(uData.lPlayerID, 1, lPlayerID, ySide, lPS_UB, bConflict)

			If bConflict = True Then
				For X As Int32 = 0 To lPS_UB
					Dim sPlayer As String = GetPlayerName(lPlayerID(X))
					If sPlayer <> "" Then
						If .PlayerSideListA <> "" Then .PlayerSideListA &= ", "
						.PlayerSideListA &= sPlayer
					End If

					For Y As Int32 = 0 To mlDataUB
						If mlDataPlayer(Y) = lPlayerID(X) Then
							mlDataPlayer(Y) = -1
							.TotalFacilitiesLostSideA += muData(Y).lFacilitiesLost
							.TotalJobsLostSideA += muData(Y).lLostJobs
							.TotalLostSideA += muData(Y).lLivesLost
							.TotalResidenceLostSideA += muData(Y).lLostResidence
							.TotalUnitLostSideA += muData(Y).lUnitsLost
						End If
					Next Y
				Next X
			Else
				For X As Int32 = 0 To lPS_UB
					Dim sPlayer As String = GetPlayerName(lPlayerID(X))
					If sPlayer <> "" Then
						If ySide(X) = 1 Then
							If .PlayerSideListA <> "" Then .PlayerSideListA &= ", "
							.PlayerSideListA &= sPlayer
						Else
							If .PlayerSideListB <> "" Then .PlayerSideListB &= ", "
							.PlayerSideListB &= sPlayer
						End If
					End If
				Next X

				'Ok, now, find the last item in each list
				Dim lTmpIdx As Int32 = .PlayerSideListA.LastIndexOf(","c)
				If lTmpIdx <> -1 Then
					Dim sTemp As String = .PlayerSideListA.Substring(0, lTmpIdx)
					sTemp &= " and " & .PlayerSideListA.Substring(lTmpIdx + 1).Trim
					.PlayerSideListA = sTemp
				End If
				lTmpIdx = .PlayerSideListB.LastIndexOf(","c)
				If lTmpIdx <> -1 Then
					Dim sTemp As String = .PlayerSideListB.Substring(0, lTmpIdx)
					sTemp &= " and " & .PlayerSideListB.Substring(lTmpIdx + 1).Trim
					.PlayerSideListB = sTemp
				End If

				'Now, get our side totals...
				For X As Int32 = 0 To lPS_UB
					For Y As Int32 = 0 To mlDataUB
						If mlDataPlayer(Y) = lPlayerID(X) Then
							.EndOfBattle = Math.Max(.EndOfBattle, muData(Y).lEnd)
							.StartOfBattle = Math.Min(.StartOfBattle, muData(Y).lStart)

							If ySide(X) = 1 Then
								.TotalFacilitiesLostSideA += muData(Y).lFacilitiesLost
								.TotalJobsLostSideA += muData(Y).lLostJobs
								.TotalLostSideA += muData(Y).lLivesLost
								.TotalResidenceLostSideA += muData(Y).lLostResidence
								.TotalUnitLostSideA += muData(Y).lUnitsLost
							Else
								.TotalFacilitiesLostSideB += muData(Y).lFacilitiesLost
								.TotalJobsLostSideB += muData(Y).lLostJobs
								.TotalLostSideB += muData(Y).lLivesLost
								.TotalResidenceLostSideB += muData(Y).lLostResidence
								.TotalUnitLostSideB += muData(Y).lUnitsLost
							End If

							Exit For
						End If
					Next Y
					.TotalInvolved += 1

					For Y As Int32 = 0 To mlDataUB
						If mlDataPlayer(Y) = lPlayerID(X) Then
							mlDataPlayer(Y) = -1
						End If
					Next Y
				Next X
			End If

		End With

		'Ok, now...
		If oData Is Nothing = False AndAlso oData.PlayerSideListB Is Nothing = False AndAlso oData.PlayerSideListB <> "" AndAlso oData.PlayerSideListA Is Nothing = False AndAlso oData.PlayerSideListA <> "" Then
			'ok, Now get our template
			Dim oTemplate As GNSTemplate = GNSMgr.GetGNSMgr().GetTemplate(CType(myLatestType, NewsItemType))
			If oTemplate Is Nothing = False Then
				Dim sText As String = oTemplate.GetNewsStoryText(oData)
				If sText Is Nothing Then Return

				oData.TitleText = oTemplate.GetNewsTitleText(oData)
				oData.SectionID = oTemplate.lSectionID
				oData.BodyText = sText
				oData.sAssociated = GNSData.GetAssociated(oTemplate.sAssociated, oTemplate.lSectionID, oData.LocationTypeID)
				oData.SummaryText = oTemplate.GetSummaryText(oData)
				oData.sInformantText = oTemplate.GetInformant(oData.GetTagData("ENVIRNAME"))

				'here, we save the data to the database
				If oData.PostNewsItem() = False Then
					LogEvent("GNS Post News Item Failed")
				End If
			End If
		End If

	End Sub

	Private Function GetPlayerName(ByVal lPlayerID As Int32) As String
		Dim sResult As String = ""
		For X As Int32 = 0 To glPlayerUB
			If glPlayerIdx(X) = lPlayerID Then
				sResult = goPlayer(X).sPlayerName
				Exit For
			End If
		Next X
		If sResult <> "" AndAlso sResult.Length > 2 Then
			sResult = sResult.Substring(0, 1).ToUpper & sResult.Substring(1)
		End If
		Return sResult
	End Function

	Private Sub DetermineSide(ByVal lPlayerID As Int32, ByVal ySide As Byte, ByRef lPlayerIDList() As Int32, ByRef ySideList() As Byte, ByRef lUB As Int32, ByRef bConflict As Boolean)
		'ok, add me to the side passed in...
		' now, look through my Killed list and my KilledBy list....
		' based on the current side settings, determine where each item in the list goes...
		' in the event of a conflict... discard this result or do pandemonium...

		lUB += 1
		ReDim Preserve lPlayerIDList(lUB)
		ReDim Preserve ySideList(lUB)
		lPlayerIDList(lUB) = lPlayerID
		ySideList(lUB) = ySide

		Dim lDIdx As Int32 = -1
		For X As Int32 = 0 To mlDataUB
			If mlDataPlayer(X) = lPlayerID Then
				lDIdx = X
				Exit For
			End If
		Next X
		If lDIdx <> -1 Then
			Dim uData As GNSCombatData = muData(lDIdx)

			For X As Int32 = 0 To uData.lKilledByUB
				If uData.uKilledBy(X).lPlayerID > 0 Then
					'check if it is already in the list
					Dim bFound As Boolean = False
					For Y As Int32 = 0 To lUB
						If lPlayerIDList(Y) = uData.uKilledBy(X).lPlayerID Then
							'ok, it is
							bFound = True
							'I was killed by this player, are the sides ok?
							If ySideList(Y) = ySide Then
								'no...
								'conflict!
								bConflict = True
							End If
							Exit For
						End If
					Next Y
					If bFound = False Then
						If ySide = 1 Then
							DetermineSide(uData.uKilledBy(X).lPlayerID, 2, lPlayerIDList, ySideList, lUB, bConflict)
						Else : DetermineSide(uData.uKilledBy(X).lPlayerID, 1, lPlayerIDList, ySideList, lUB, bConflict)
						End If
					End If
				End If
				'If bConflict = True Then Return
			Next X

			'Now, who did I kill?
			For X As Int32 = 0 To uData.lKilledUB
				If uData.uKilled(X).lPlayerID > 0 Then
					'check if it is already in the list
					Dim bFound As Boolean = False
					For Y As Int32 = 0 To lUB
						If lPlayerIDList(Y) = uData.uKilled(X).lPlayerID Then
							'ok, it is
							bFound = True
							'I was killed by this player, are the sides ok?
							If ySideList(Y) = ySide Then
								'no...
								'conflict!
								bConflict = True
							End If
							Exit For
						End If
					Next Y
					If bFound = False Then
						If ySide = 1 Then
							DetermineSide(uData.uKilled(X).lPlayerID, 2, lPlayerIDList, ySideList, lUB, bConflict)
						Else : DetermineSide(uData.uKilled(X).lPlayerID, 1, lPlayerIDList, ySideList, lUB, bConflict)
						End If
					End If
				End If
				'If bConflict = True Then Return
			Next X
		End If
	End Sub

End Class

Public Structure GNSCombatData
	Public lFightLocID As Int32
	Public iFightLocTypeID As Int16
	Public sLocName As String
	Public lPlayerID As Int32
	Public lStart As Int32
	Public lEnd As Int32
	Public lLostResidence As Int32
	Public lLostJobs As Int32
	Public lLivesLost As Int32
	Public lUnitsLost As Int32
	Public lFacilitiesLost As Int32
	Public lUnitsKilled As Int32
	Public lFacilitiesKilled As Int32

	Public sPlayerName As String
	Public yPlayerGender As Byte
	Public sEmpireName As String


	'Public lKilledBy() As Int32
	'Public lKilled() As Int32

	Public Structure KillListItem
		Public lPlayerID As Int32
		Public sPlayerName As String
		Public yGender As Byte
	End Structure

	Public lKilledByUB As Int32
	Public uKilledBy() As KillListItem
	Public lKilledUB As Int32
	Public uKilled() As KillListItem

	Public Sub SetFromGNSMsg(ByRef yData() As Byte, ByVal lPos As Int32)
		lFightLocID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		iFightLocTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		sLocName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lStart = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lEnd = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lLostResidence = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lLostJobs = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lLivesLost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lUnitsLost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lFacilitiesLost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lUnitsKilled = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lFacilitiesKilled = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		sPlayerName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		yPlayerGender = yData(lPos) : lPos += 1
		sEmpireName = GetStringFromBytes(yData, lPos, 20) : lPos += 20

		lKilledByUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
		ReDim uKilledBy(lKilledByUB)
		For X As Int32 = 0 To lKilledByUB
			With uKilledBy(X)
				.lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.sPlayerName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
				.yGender = yData(lPos) : lPos += 1
			End With
		Next X
		lKilledUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
		ReDim uKilled(lKilledUB)
		For X As Int32 = 0 To lKilledUB
			With uKilled(X)
				.lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.sPlayerName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
				.yGender = yData(lPos) : lPos += 1
			End With
		Next X

		'Now, verify our data...
		Dim bDone As Boolean = False
		While bDone = False
			bDone = True
			For X As Int32 = 0 To lKilledByUB
				If uKilledBy(X).lPlayerID < 1 Then
					bDone = False
					For Y As Int32 = X To lKilledByUB - 1
						uKilledBy(Y) = uKilledBy(Y + 1)
					Next Y
					lKilledByUB -= 1
					Exit For
				End If
			Next X
		End While
		bDone = False
		While bDone = False
			bDone = True
			For X As Int32 = 0 To lKilledUB
				If uKilled(X).lPlayerID < 1 Then
					bDone = False
					For Y As Int32 = X To lKilledUB - 1
						uKilled(Y) = uKilled(Y + 1)
					Next Y
					lKilledUB -= 1
					Exit For
				End If
			Next X
		End While

		'lKilledByUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
		'ReDim lKilledBy(lKilledByUB)
		'For X As Int32 = 0 To lKilledByUB
		'	lKilledBy(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		'Next X
		'lKilledUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
		'ReDim lKilled(lKilledUB)
		'For X As Int32 = 0 To lKilledUB
		'	lKilled(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		'Next X
	End Sub
End Structure

