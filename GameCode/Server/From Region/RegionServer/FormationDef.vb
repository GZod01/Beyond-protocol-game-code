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

Public Class FormationDef
	Private Const ml_GRID_SIZE_WH As Int32 = 25
	Public Structure FormationPoint
		Public ptRel As Point	'relative +/- coordinate from the centerpoint
		Public iCriteriaSortOrder As Int16
	End Structure

	Public lFormationID As Int32

	Public yFormationName(19) As Byte		'Claw, Sentinel, Sphere, whatever

	Public lPlayerID As Int32
	Public yDefault As Byte
	Public yCriteria As CriteriaType
	Public lCellSize As Int32

	Public uLocs() As FormationPoint
	Public lLocUB As Int32 = -1

	Public Sub FillFromMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2	'for msgcode

		lFormationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yDefault = yData(lPos) : lPos += 1
		yCriteria = CType(yData(lPos), CriteriaType) : lPos += 1
		'name, not needed here...
		lPos += 20
		lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lCellSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		lLocUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
		ReDim uLocs(lLocUB)
		For lIdx As Int32 = 0 To lLocUB
			Dim lSlotIdx As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim iSlotVal As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

			Dim lSlotY As Int32 = lSlotIdx \ ml_GRID_SIZE_WH
			Dim lSlotX As Int32 = lSlotIdx - (lSlotY * ml_GRID_SIZE_WH)

			If lSlotX < ml_GRID_SIZE_WH AndAlso lSlotY < ml_GRID_SIZE_WH AndAlso lSlotX > -1 AndAlso lSlotY > -1 Then
				uLocs(lIdx).iCriteriaSortOrder = iSlotVal
				uLocs(lIdx).ptRel.X = lSlotX
				uLocs(lIdx).ptRel.Y = lSlotY
			End If
		Next lIdx
	End Sub
End Class
