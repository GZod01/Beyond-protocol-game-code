Option Strict On

Public Enum CriteriaType As Byte
	eHullSize = 0
	eWeaponSlots
	eMaxRadarRange
	eMostFrontArmorHP
	eMostShieldHP
	eManeuver
	eSpeed
	eCombatRating
	eCargoBayCap
	eHangarCap
	eEntityName
End Enum

'An instance of a formation
Public Class Formation
	Public Structure FormationMoveItem
		Public lObjID As Int32
		Public iObjTypeID As Int16
		Public lLocX As Int32
		Public lLocZ As Int32
		Public iLocA As Int16
        'Public iModelID As Int16
        Private miModelID As Int16
        Public Property iModelID() As Int16
            Get
                Return miModelID
            End Get
            Set(ByVal value As Int16)
                'MSC: the pathfinding server only cares about the first byte of the model id - the model
                miModelID = (value And 255S)
            End Set
        End Property
		Public yCritType As CriteriaType
		Public lCritValue As Int32

		Public lSlotLocIdx As Int32

		Public ptCenterPtRel As Point

		Public bItemInPlace As Boolean

		Public lMoverIdx As Int32		'index within the global glMoverIdx array, verify the mover versus the GUID
	End Structure

	Public ServerIndex As Int32 = -1			'index within the global formation array
	Public DomainServerIndex As Int32 = -1		'domain server to interact with

	Public oFormationDef As FormationDefinition
	Public iFormationDirection As Int16
	Public ptCenterPoint As Point
	Public muItems() As FormationMoveItem
	Public mlItemUB As Int32 = -1

	Public ptFinalDest As Point
	Public lDestID As Int32
	Public iDestTypeID As Int16
	Public bSetFinalDest As Boolean = False

	Public Sub DetachItem(ByVal lObjID As Int32, ByVal iObjTypeID As Int16)
		For X As Int32 = 0 To mlItemUB
			If muItems(X).lObjID = lObjID AndAlso muItems(X).iObjTypeID = iObjTypeID Then
				muItems(X).lMoverIdx = -1
				muItems(X).lObjID = -1
				muItems(X).iObjTypeID = -1
			End If
		Next X

		If CheckForDeletion() = True Then Return

		CheckForArrival()
	End Sub

	Public Function CheckForDeletion() As Boolean
		Dim lCnt As Int32 = 0
		For X As Int32 = 0 To mlItemUB
			If muItems(X).lObjID <> -1 AndAlso muItems(X).iObjTypeID <> -1 AndAlso muItems(X).lMoverIdx <> -1 Then
				If glMoverIdx(muItems(X).lMoverIdx) = muItems(X).lObjID Then
					lCnt += 1
					If lCnt > 1 Then Return False
				End If
			End If
		Next X

		If ServerIndex <> -1 Then
			glFormationIdx(ServerIndex) = -1
			goFormation(ServerIndex) = Nothing
		End If

		Return True
	End Function

	Public Sub ItemInPlace(ByVal lObjID As Int32, ByVal iObjTypeID As Int16)
		For X As Int32 = 0 To mlItemUB
			If muItems(X).lObjID = lObjID AndAlso muItems(X).iObjTypeID = iObjTypeID Then
				muItems(X).bItemInPlace = True
				Exit For
			End If
		Next X
	End Sub

	Public Sub CheckForArrival()
		Dim lCnt As Int32 = 0
		'Now, check if all elements are in place
		For X As Int32 = 0 To mlItemUB
			If muItems(X).lObjID <> -1 AndAlso muItems(X).iObjTypeID <> -1 AndAlso muItems(X).lMoverIdx <> -1 Then
				If muItems(X).bItemInPlace = False Then Return
				lCnt += 1
			End If
		Next X

		'Ok, we are here, now, mark all items as integrity broken
		For X As Int32 = 0 To mlItemUB
			If muItems(X).lObjID <> -1 AndAlso muItems(X).iObjTypeID <> -1 AndAlso muItems(X).lMoverIdx <> -1 Then muItems(X).bItemInPlace = False
		Next X

		'Now, send the MoveFormation message to the domain server
		If bSetFinalDest = False Then
			bSetFinalDest = True

			Dim yMsg(15 + (lCnt * 14)) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eMoveFormation).CopyTo(yMsg, lPos) : lPos += 2
			lPos += 2	'for maxspeed/maneuver placeholder
			System.BitConverter.GetBytes(lDestID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(iDestTypeID).CopyTo(yMsg, lPos) : lPos += 2

			Dim iAdjAngle As Int16 = iFormationDirection + 900S
			If iAdjAngle > 3600 Then iAdjAngle -= 3600S

			System.BitConverter.GetBytes(iAdjAngle).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

			For X As Int32 = 0 To mlItemUB
				If muItems(X).lObjID <> -1 AndAlso muItems(X).iObjTypeID <> -1 AndAlso muItems(X).lMoverIdx <> -1 Then
					System.BitConverter.GetBytes(muItems(X).lObjID).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(muItems(X).iObjTypeID).CopyTo(yMsg, lPos) : lPos += 2

					Dim lActDestX As Int32 = muItems(X).ptCenterPtRel.X + ptFinalDest.X
					Dim lActDestY As Int32 = muItems(X).ptCenterPtRel.Y + ptFinalDest.Y
					System.BitConverter.GetBytes(lActDestX).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(lActDestY).CopyTo(yMsg, lPos) : lPos += 4
				End If
			Next X

			goMsgSys.SendMsgToDomain(yMsg, Me.DomainServerIndex)
		End If
	End Sub

End Class


