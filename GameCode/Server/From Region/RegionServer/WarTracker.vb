Option Strict On

Public Class WarTracker
	Public lStart As Int32 = -1				'date as number
	Private mlLastUpdateCycle As Int32 = 0

	Public lPlayerID As Int32

	Public lLostResidence As Int32 = 0
	Public lLostJobs As Int32 = 0
	Public lLivesLost As Int32 = 0

	Public lUnitsLost As Int32 = 0
	Public lFacilitiesLost As Int32 = 0

	Public lUnitsKilled As Int32 = 0
	Public lFacilitiesKilled As Int32 = 0

	Public lKilledByID() As Int32
	Public lKilledByUB As Int32 = -1

	Public lKilledID() As Int32
	Public lKilledUB As Int32 = -1

	Public bExpired As Boolean = False
	Public lPreviousMsgSend As Int32 = 0
	Private mbFirstMsgSent As Boolean = False
	Public Const ml_WAR_OVER_DELAY As Int32 = 27000	'15 minutes

	Private Sub AddKilledBy(ByVal lPlayerID As Int32)
		For X As Int32 = 0 To lKilledByUB
			If lKilledByID(X) = lPlayerID Then Return
		Next X
		lKilledByUB += 1
		ReDim Preserve lKilledByID(lKilledByUB)
		lKilledByID(lKilledByUB) = lPlayerID
	End Sub
	Private Sub AddKilled(ByVal lPlayerID As Int32)
		For X As Int32 = 0 To lKilledUB
			If lKilledID(X) = lPlayerID Then Return
		Next X
		lKilledUB += 1
		ReDim Preserve lKilledID(lKilledUB)
		lKilledID(lKilledUB) = lPlayerID
	End Sub

	Public Sub AddEntityKilled(ByRef oDef As Epica_Entity_Def, ByVal lKilledBy As Int32)
		If lKilledBy = lPlayerID Then Return
		mlLastUpdateCycle = glCurrentCycle
		lLivesLost += oDef.MaxCrew
		If oDef.ObjTypeID = ObjectType.eFacilityDef Then
			lLostJobs += oDef.WorkerFactor
			If oDef.ProductionTypeID = ProductionType.eColonists Then lLostResidence += oDef.ProdFactor
			lFacilitiesLost += 1
		Else : lUnitsLost += 1
		End If
		AddKilledBy(lKilledBy)
	End Sub
	Public Sub AddKill(ByVal lOwnerID As Int32, ByVal iKilledTypeID As Int16)
		If lOwnerID = lPlayerID Then Return
		mlLastUpdateCycle = glCurrentCycle
		If iKilledTypeID = ObjectType.eFacility Then
			lFacilitiesKilled += 1
		Else : lUnitsKilled += 1
		End If
		AddKilled(lOwnerID)
	End Sub

	Public Function GetNewsItemMsg(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Byte()
		Dim yMsg(82 + ((lKilledByUB + lKilledUB + 2) * 4)) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yMsg, lPos) : lPos += 2

		If glCurrentCycle - mlLastUpdateCycle > ml_WAR_OVER_DELAY Then
			If iEnvirTypeID = ObjectType.ePlanet Then
				yMsg(lPos) = NewsItemType.ePlanetCombatEnd : lPos += 1
			Else : yMsg(lPos) = NewsItemType.eSpaceCombatEnd : lPos += 1
			End If
		ElseIf mbFirstMsgSent = False Then
			If iEnvirTypeID = ObjectType.ePlanet Then
				yMsg(lPos) = NewsItemType.ePlanetCombat : lPos += 1
			Else : yMsg(lPos) = NewsItemType.eSpaceCombat : lPos += 1
			End If
		Else
			If iEnvirTypeID = ObjectType.ePlanet Then
				yMsg(lPos) = NewsItemType.ePlanetCombatUpdate : lPos += 1
			Else : yMsg(lPos) = NewsItemType.eSpaceCombatUpdate : lPos += 1
			End If
		End If
		mbFirstMsgSent = True
		lPreviousMsgSend = glCurrentCycle

		System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2

		'System.Text.ASCIIEncoding.ASCII.GetBytes("some location").CopyTo(yMsg, lPos) : lPos += 20
		lPos += 20

		System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lStart).CopyTo(yMsg, lPos) : lPos += 4

		If glCurrentCycle - mlLastUpdateCycle > ml_WAR_OVER_DELAY Then
            Dim lSeconds As Int32 = (glCurrentCycle - mlLastUpdateCycle) \ 30
			Dim lEnd As Int32 = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lSeconds)))
			System.BitConverter.GetBytes(lEnd).CopyTo(yMsg, lPos) : lPos += 4
			bExpired = True
		Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
		End If

		System.BitConverter.GetBytes(lLostResidence).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lLostJobs).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lLivesLost).CopyTo(yMsg, lPos) : lPos += 4

		System.BitConverter.GetBytes(lUnitsLost).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lFacilitiesLost).CopyTo(yMsg, lPos) : lPos += 4

		System.BitConverter.GetBytes(lUnitsKilled).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lFacilitiesKilled).CopyTo(yMsg, lPos) : lPos += 4

		System.BitConverter.GetBytes(lKilledByUB + 1).CopyTo(yMsg, lPos) : lPos += 4
		For X As Int32 = 0 To lKilledByUB
			System.BitConverter.GetBytes(lKilledByID(X)).CopyTo(yMsg, lPos) : lPos += 4
		Next X
		System.BitConverter.GetBytes(lKilledUB + 1).CopyTo(yMsg, lPos) : lPos += 4
		For X As Int32 = 0 To lKilledUB
			System.BitConverter.GetBytes(lKilledID(X)).CopyTo(yMsg, lPos) : lPos += 4
		Next X

		Return yMsg
	End Function
End Class
