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
	Public FormationID As Int32
	Public mlSlots(,) As Int32
	Public sName As String
	Public yDefault As Byte
	Public yCriteria As CriteriaType
	Public lOwnerID As Int32
	Public lCellSize As Int32

	Public ptLocs() As Point
	Public lLocUB As Int32 = -1

	Public Sub FillFromMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2	'for msgcode

		FormationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim bIsDefault As Boolean = yDefault <> 0
		yDefault = yData(lPos) : lPos += 1
		yCriteria = CType(yData(lPos), CriteriaType) : lPos += 1
		sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		lOwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lCellSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'If bIsDefault = True AndAlso yDefault = 0 Then
        '	yDefault = 1
        'ElseIf bIsDefault = False AndAlso yDefault <> 0 Then
        '	Try
        '		For X As Int32 = 0 To goCurrentPlayer.lFormationUB
        '			If goCurrentPlayer.lFormationIdx(X) <> -1 Then
        '				If goCurrentPlayer.oFormations(X).yDefault <> 0 AndAlso goCurrentPlayer.oFormations(X).FormationID <> Me.FormationID Then goCurrentPlayer.oFormations(X).yDefault = 0
        '			End If
        '		Next X
        '	Catch
        '	End Try
        'End If

		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		lLocUB = lCnt - 1
		ReDim ptLocs(lLocUB)

		ReDim mlSlots(UIFormation.ml_GRID_SIZE_WH - 1, UIFormation.ml_GRID_SIZE_WH - 1)

		For lIdx As Int32 = 0 To lCnt - 1
			Dim lSlotIdx As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim iSlotVal As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

			Dim lSlotY As Int32 = lSlotIdx \ UIFormation.ml_GRID_SIZE_WH
			Dim lSlotX As Int32 = lSlotIdx - (lSlotY * UIFormation.ml_GRID_SIZE_WH)

			If lSlotX < UIFormation.ml_GRID_SIZE_WH AndAlso lSlotY < UIFormation.ml_GRID_SIZE_WH AndAlso lSlotX > -1 AndAlso lSlotY > -1 Then
				ptLocs(lIdx).X = lSlotX
				ptLocs(lIdx).Y = lSlotY
				mlSlots(lSlotX, lSlotY) = iSlotVal
			End If
		Next lIdx
	End Sub
 
End Class
