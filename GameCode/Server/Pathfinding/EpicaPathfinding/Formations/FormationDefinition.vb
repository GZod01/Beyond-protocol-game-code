Option Strict On

Public Class FormationDefinition
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

	Public Sub HandleMoveFormation(ByRef yData() As Byte, ByVal lDomainServerIndex As Int32)
		Dim lPos As Int32 = 6	'for msgcode and formation ID

		Dim oFormation As Formation = New Formation()
		oFormation.oFormationDef = Me

        'Now, get our data from the region... first is dest
        Dim bMoveThenForm As Boolean = False
        With oFormation
            bMoveThenForm = yData(lPos) <> 0 : lPos += 1

            .ptFinalDest.X = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ptFinalDest.Y = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .iFormationDirection = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .lDestID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .iDestTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .DomainServerIndex = lDomainServerIndex
        End With

		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		'For determining the first center point
		Dim lMinSpeedID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iMinSpeedTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lFirstPtX As Int32
		Dim lFirstPtZ As Int32

		'Initialize our items array
		oFormation.mlItemUB = lCnt - 1
		ReDim oFormation.muItems(oFormation.mlItemUB)

		'ok... now, this has our values
		For X As Int32 = 0 To lCnt - 1
			With oFormation.muItems(X)
				.lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				.lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.iLocA = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .iModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lCritValue = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lMoverIdx = -1

                If .lObjID = lMinSpeedID AndAlso .iObjTypeID = iMinSpeedTypeID Then
                    lFirstPtX = .lLocX
                    lFirstPtZ = .lLocZ
                End If

                'Dim oMover As Mover = GetOrAddMover(.lObjID, .iObjTypeID, .lLocX, .lLocZ, .iLocA, .iModelID, oFormation.lDestID, oFormation.iDestTypeID)
                Dim lMoverIdx As Int32 = GetOrAddMover(.lObjID, .iObjTypeID, .lLocX, .lLocZ, .iLocA, .iModelID, oFormation.lDestID, oFormation.iDestTypeID)
                If lMoverIdx > -1 Then ' Is Nothing = False Then
                    .lMoverIdx = lMoverIdx
                    If goMovers(lMoverIdx).oFormation Is Nothing = False Then goMovers(lMoverIdx).oFormation.DetachItem(.lObjID, .iObjTypeID)
                    goMovers(lMoverIdx).oFormation = oFormation
                End If
				'For Y As Int32 = 0 To glMoverUB
				'	If glMoverIdx(Y) = .lObjID AndAlso goMovers(Y).ObjTypeID = .iObjTypeID Then
				'		.lMoverIdx = Y

				'		'Now, detach the mover from their current formation
				'		If goMovers(Y).oFormation Is Nothing = False Then
				'			goMovers(Y).oFormation.DetachItem(.lObjID, .iObjTypeID)
				'		End If
				'		'add attach this formation
				'		goMovers(Y).oFormation = oFormation

				'		Exit For
				'	End If
				'Next Y
			End With
		Next X

        'Do a presort on the ID...
        With oFormation
            Dim uPreSort(.mlItemUB) As Formation.FormationMoveItem
            Dim lPreSortUB As Int32 = -1
            For X As Int32 = 0 To .mlItemUB
                Dim lIdx As Int32 = -1
                For Y As Int32 = 0 To lPreSortUB
                    If uPreSort(Y).lObjID > .muItems(X).lObjID Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lPreSortUB += 1
                If lIdx = -1 Then
                    uPreSort(lPreSortUB) = .muItems(X)
                Else
                    For Y As Int32 = lPreSortUB To lIdx + 1 Step -1
                        uPreSort(Y) = uPreSort(Y - 1)
                    Next
                    uPreSort(lIdx) = .muItems(X)
                End If
            Next X

            .muItems = uPreSort
        End With


		'Now, let's assign slots... we do a bubble sort on our items?
		Dim lValue As Int32 = Int32.MinValue
		Dim bDone As Boolean = False
		Dim lCurrentIdx As Int32 = -1
		Dim lHalfSize As Int32 = ml_GRID_SIZE_WH \ 2
		Dim lTrueCellSize As Int32 = Me.lCellSize * gl_FINAL_GRID_SQUARE_SIZE
		Dim lLeftXMod As Int32
		Dim lRightXMod As Int32
		Dim lFrontZMod As Int32
		Dim lBackZMod As Int32
		Dim lPlaceID As Int32 = 3	'0 = front, 1 = left, 2 = back, 3 = right
		Dim bFirst As Boolean = True
        Dim fAngle As Single = oFormation.iFormationDirection / 10.0F

        Dim lSpaceMod As Int32 = 1
        Dim lSpaceModX As Int32 = 0
        Dim lSpaceModZ As Int32 = 0

		With oFormation
			While bDone = False
				Dim lNextValue As Int32 = Int32.MaxValue
				For X As Int32 = 0 To .mlItemUB
					If .muItems(X).lCritValue = lValue Then
						lCurrentIdx += 1
						.muItems(X).lSlotLocIdx = lCurrentIdx
					ElseIf .muItems(X).lCritValue < lNextValue AndAlso .muItems(X).lCritValue > lValue Then
						lNextValue = .muItems(X).lCritValue
					End If
				Next X
				If lNextValue = Int32.MaxValue Then bDone = True
				lValue = lNextValue
			End While

			'Now, we get assign our points...
			For X As Int32 = 0 To .mlItemUB
				lCurrentIdx = .muItems(X).lSlotLocIdx
				Dim bDefault As Boolean = False
				If lCurrentIdx > lLocUB Then
					lCurrentIdx = lLocUB
					bDefault = True
				End If
				If lCurrentIdx < 0 Then lCurrentIdx = 0

				If bDefault = False Then
					.muItems(X).ptCenterPtRel.X = (uLocs(lCurrentIdx).ptRel.X - lHalfSize) * lTrueCellSize
					.muItems(X).ptCenterPtRel.Y = (uLocs(lCurrentIdx).ptRel.Y - lHalfSize) * lTrueCellSize
                Else
                    Select Case yDefault
                        Case 1
                            'HORIZONTAL FORMATION
                            '====================
                            'Increment our space usage
                            If bFirst = True Then
                                bFirst = False
                                lFrontZMod = 1
                                lLeftXMod = -1
                                lBackZMod = -1
                                lRightXMod = 1
                            Else
                                Select Case lPlaceID
                                    Case 0 : lLeftXMod -= 1
                                    Case 1 : lRightXMod += 1
                                End Select
                            End If

                            lPlaceID += 1
                            If lPlaceID > 1 Then lPlaceID = 0

                            'Get our next coordinate
                            Dim lTX As Int32 = 0
                            Dim lTZ As Int32 = 0
                            Select Case lPlaceID
                                Case 0 : lTX = lLeftXMod : lTZ = 0
                                Case Else : lTX = lRightXMod : lTZ = 0
                            End Select
                            lTX += uLocs(lCurrentIdx).ptRel.X
                            lTZ += uLocs(lCurrentIdx).ptRel.Y

                            .muItems(X).ptCenterPtRel.X = (lTX - lHalfSize) * lTrueCellSize
                            .muItems(X).ptCenterPtRel.Y = (lTZ - lHalfSize) * lTrueCellSize
                        Case 2
                            'VERTICAL FORMATION
                            '==================
                            'Increment our space usage
                            If bFirst = True Then
                                bFirst = False
                                lFrontZMod = 1
                                lLeftXMod = -1
                                lBackZMod = -1
                                lRightXMod = 1
                            Else
                                Select Case lPlaceID
                                    Case 0 : lFrontZMod += 1
                                    Case 1 : lBackZMod -= 1
                                End Select
                            End If

                            lPlaceID += 1
                            If lPlaceID > 1 Then lPlaceID = 0

                            'Get our next coordinate
                            Dim lTX As Int32 = 0
                            Dim lTZ As Int32 = 0
                            Select Case lPlaceID
                                Case 0 : lTX = 0 : lTZ = lFrontZMod
                                Case Else : lTX = 0 : lTZ = lBackZMod
                            End Select
                            lTX += uLocs(lCurrentIdx).ptRel.X
                            lTZ += uLocs(lCurrentIdx).ptRel.Y

                            .muItems(X).ptCenterPtRel.X = (lTX - lHalfSize) * lTrueCellSize
                            .muItems(X).ptCenterPtRel.Y = (lTZ - lHalfSize) * lTrueCellSize
                        Case 3
                            'MASS FORMATION
                            '==============
                            If bFirst = True Then
                                bFirst = False
                                lSpaceMod = 1
                                lSpaceModX = -2
                                lSpaceModZ = -1
                                lPlaceID = 0
                            Else

                                Select Case lPlaceID
                                    Case 0
                                        'moving right
                                        lSpaceModX += 1
                                        If lSpaceModX = lSpaceMod Then
                                            lPlaceID += 1
                                        End If
                                    Case 1
                                        'moving down
                                        lSpaceModZ += 1
                                        If lSpaceModZ = lSpaceMod Then
                                            lPlaceID += 1
                                        End If
                                    Case 2
                                        'moving left
                                        lSpaceModX -= 1
                                        If lSpaceModX = -lSpaceMod Then
                                            lPlaceID += 1
                                        End If
                                    Case 3
                                        'moving up
                                        lSpaceModZ -= 1
                                        If lSpaceModZ = -lSpaceMod Then
                                            lPlaceID = 0
                                            lSpaceMod += 1
                                            lSpaceModX = -lSpaceMod
                                            lSpaceModZ = -lSpaceMod
                                        End If
                                End Select

                                Dim lTX As Int32 = lSpaceModX + uLocs(lCurrentIdx).ptRel.X
                                Dim lTZ As Int32 = lSpaceModZ + uLocs(lCurrentIdx).ptRel.Y

                                .muItems(X).ptCenterPtRel.X = (lTX - lHalfSize) * lTrueCellSize
                                .muItems(X).ptCenterPtRel.Y = (lTZ - lHalfSize) * lTrueCellSize
                            End If
                        Case Else
                            'STAR FORMATION
                            '==============
                            'Increment our space usage
                            If bFirst = True Then
                                bFirst = False
                                lFrontZMod = 1
                                lLeftXMod = -1
                                lBackZMod = -1
                                lRightXMod = 1
                            Else
                                Select Case lPlaceID
                                    Case 0 : lFrontZMod += 1
                                    Case 1 : lLeftXMod -= 1
                                    Case 2 : lBackZMod -= 1
                                    Case 3 : lRightXMod += 1
                                End Select
                            End If

                            lPlaceID += 1
                            If lPlaceID = 8 Then lPlaceID = 0

                            'Get our next coordinate
                            Dim lTX As Int32 = 0
                            Dim lTZ As Int32 = 0
                            Select Case lPlaceID
                                Case 0 : lTX = 0 : lTZ = lFrontZMod
                                Case 1 : lTX = lLeftXMod : lTZ = 0
                                Case 2 : lTX = 0 : lTZ = lBackZMod
                                Case 3 : lTX = lRightXMod : lTZ = 0
                                Case 4 : lTX = lLeftXMod : lTZ = lFrontZMod
                                Case 5 : lTX = lRightXMod : lTZ = lFrontZMod
                                Case 6 : lTX = lLeftXMod : lTZ = lBackZMod
                                Case Else : lTX = lRightXMod : lTZ = lBackZMod
                            End Select
                            lTX += uLocs(lCurrentIdx).ptRel.X
                            lTZ += uLocs(lCurrentIdx).ptRel.Y

                            .muItems(X).ptCenterPtRel.X = (lTX - lHalfSize) * lTrueCellSize
                            .muItems(X).ptCenterPtRel.Y = (lTZ - lHalfSize) * lTrueCellSize
                    End Select
                   

				End If

				Dim fPtX As Single = .muItems(X).ptCenterPtRel.X
				Dim fPtZ As Single = .muItems(X).ptCenterPtRel.Y
				RotatePoint(0, 0, fPtX, fPtZ, fAngle)
				.muItems(X).ptCenterPtRel.X = CInt(fPtX)
				.muItems(X).ptCenterPtRel.Y = CInt(fPtZ)

			Next X
		End With

		oFormation.bSetFinalDest = False

		AddFormationInstance(oFormation)

		Dim yResp() As Byte
		Dim yFinal() As Byte = Nothing
		Dim lDestPos As Int32 = 0

        Dim iTempAngle As Int16 = CShort(LineAngleDegrees(lFirstPtX, lFirstPtZ, oFormation.ptFinalDest.X, oFormation.ptFinalDest.Y) * 10)

        If bMoveThenForm = True Then
            lFirstPtX = oFormation.ptFinalDest.X
            lFirstPtZ = oFormation.ptFinalDest.Y
        End If

        'iTempAngle += 1800S
        'If iTempAngle > 3600 Then iTempAngle -= 3600S

        'Now, we have everything we need... so, plot paths to the first point...
        For X As Int32 = 0 To oFormation.mlItemUB
            If oFormation.muItems(X).lMoverIdx <> -1 Then
                Dim oMover As Mover = goMovers(oFormation.muItems(X).lMoverIdx)
                If oMover Is Nothing = False Then
                    Dim uDestLoc As DestLoc
                    With uDestLoc
                        '.DestAngle = oFormation.iFormationDirection + 900S
                        'If .DestAngle > 3600 Then .DestAngle -= 3600S
                        .DestAngle = iTempAngle
                        .DestX = oFormation.muItems(X).ptCenterPtRel.X + lFirstPtX
                        .DestZ = oFormation.muItems(X).ptCenterPtRel.Y + lFirstPtZ
                        .iDestTypeID = oFormation.iDestTypeID
                        .lDestID = oFormation.lDestID
                    End With
                    If oMover.colDests Is Nothing = False Then oMover.colDests.Clear()
                    oMover.colDests = Nothing
                    Dim colTemp As Collection = New Collection
                    colTemp.Add(uDestLoc)
                    ''Try this to see if we can fake the final movement...
                    'Dim uDestLoc2 As DestLoc
                    'With uDestLoc2
                    '    Dim fTmpX As Single = 5
                    '    Dim fTmpZ As Single = 0
                    '    RotatePoint(0, 0, fTmpX, fTmpZ, fAngle)
                    '    .DestX = CInt(uDestLoc.DestX + fTmpX)
                    '    .DestZ = CInt(uDestLoc.DestZ + fTmpZ)
                    '    .ySpecialOp = 255
                    '    .DestAngle = oFormation.iFormationDirection
                    '    .iDestTypeID = oFormation.iDestTypeID
                    '    .lDestID = oFormation.lDestID
                    'End With
                    'colTemp.Add(uDestLoc2)
                    oMover.colDests = colTemp

                    yResp = oMover.ProcessStopMessage(oFormation.muItems(X).lLocX, oFormation.muItems(X).lLocZ, oFormation.muItems(X).iLocA)
                    lDestPos = MsgSystem.AppendLenAppendedMsg(yResp, yFinal, lDestPos, 1000)
                End If
            End If
        Next X
        If bMoveThenForm = True Then oFormation.bSetFinalDest = True

        yFinal = MsgSystem.FinalizeLenAppendedMsg(yFinal, lDestPos)
        'Now, send all responses in one packet
        If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then
            goMsgSys.SendLenReadyMsgToDomain(yFinal, oFormation.DomainServerIndex)
        End If

    End Sub

	Public Sub FillFromMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2	'for msgcode

		lFormationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Dim bIsDefault As Boolean = yDefault <> 0
		yDefault = yData(lPos) : lPos += 1

		yCriteria = CType(yData(lPos), CriteriaType) : lPos += 1
		'name, not needed here...
		lPos += 20
		lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lCellSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		lLocUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

        'If bIsDefault = True AndAlso yDefault = 0 Then
        '	yDefault = 1
        'ElseIf bIsDefault = False AndAlso yDefault <> 0 Then
        '	Try
        '		For X As Int32 = 0 To glFormationDefUB
        '			If glFormationDefIdx(X) <> -1 Then
        '				If goFormationDef(X).lPlayerID = Me.lPlayerID AndAlso goFormationDef(X).yDefault <> 0 AndAlso goFormationDef(X).lFormationID <> Me.lFormationID Then goFormationDef(X).yDefault = 0
        '			End If
        '		Next X
        '	Catch
        '	End Try
        'End If

		ReDim uLocs(lLocUB)
		For lIdx As Int32 = 0 To lLocUB
			Dim lSlotIdx As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim iSlotVal As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

			Dim lSlotY As Int32 = lSlotIdx \ ml_GRID_SIZE_WH
			Dim lSlotX As Int32 = ml_GRID_SIZE_WH - (lSlotIdx - (lSlotY * ml_GRID_SIZE_WH))

            If lSlotX <= ml_GRID_SIZE_WH AndAlso lSlotY <= ml_GRID_SIZE_WH AndAlso lSlotX > -1 AndAlso lSlotY > -1 Then
                uLocs(lIdx).iCriteriaSortOrder = iSlotVal
                uLocs(lIdx).ptRel.X = lSlotX
                uLocs(lIdx).ptRel.Y = lSlotY
            End If
		Next lIdx

		'now, sort our locs
		Dim lSorted() As Int32 = Nothing
		Dim lSortedUB As Int32 = -1
		For X As Int32 = 0 To lLocUB
			Dim lIdx As Int32 = -1

			For Y As Int32 = 0 To lSortedUB
				If uLocs(lSorted(Y)).iCriteriaSortOrder > uLocs(X).iCriteriaSortOrder Then
					lIdx = Y
					Exit For
				End If
			Next Y
			lSortedUB += 1
			ReDim Preserve lSorted(lSortedUB)
			If lIdx = -1 Then
				lSorted(lSortedUB) = X
			Else
				For Y As Int32 = lSortedUB To lIdx + 1 Step -1
					lSorted(Y) = lSorted(Y - 1)
				Next Y
				lSorted(lIdx) = X
			End If
		Next X

		Dim uTemp(lLocUB) As FormationPoint	'= uLocs
		For X As Int32 = 0 To lLocUB
			uTemp(X) = uLocs(lSorted(X))
		Next X
		uLocs = uTemp
	End Sub
End Class

