Option Strict On

Public Class HullTech
	Inherits Base_Tech

	Public Const ml_GRID_SIZE_WH As Int32 = 30

	Public HullName As String
	Private miModelID As Int16
	Public HullSize As Int32
	Public StructuralMineralID As Int32
	Public StructuralHitPoints As Int32

	Public yTypeID As Byte
	Public ySubTypeID As Byte

	Public yChassisType As Byte

	Public PowerRequired As Int32 = 0

	Public moSlots(,) As HullSlot

	Public Property ModelID() As Int16
		Get
			Return miModelID
		End Get
		Set(ByVal value As Int16)
			If value <> miModelID Then
				miModelID = value

				If goModelDefs Is Nothing Then Return

				Dim oModelDef As ModelDef = goModelDefs.GetModelDef(miModelID)
				If oModelDef Is Nothing = False Then
					Dim lTemp As Int32
					Dim X As Int32
					Dim Y As Int32

					With oModelDef
						'Forward
						For lIdx As Int32 = 0 To .FrontLocs.GetUpperBound(0)
							lTemp = .FrontLocs(lIdx)
							Y = lTemp \ ml_GRID_SIZE_WH
							X = lTemp - (Y * ml_GRID_SIZE_WH)
							moSlots(X, Y).yType = SlotType.eFront
						Next lIdx

						'Left
						For lIdx As Int32 = 0 To .LeftLocs.GetUpperBound(0)
							lTemp = .LeftLocs(lIdx)
							Y = lTemp \ ml_GRID_SIZE_WH
							X = lTemp - (Y * ml_GRID_SIZE_WH)
							moSlots(X, Y).yType = SlotType.eLeft
						Next lIdx

						'Rear
						For lIdx As Int32 = 0 To .RearLocs.GetUpperBound(0)
							lTemp = .RearLocs(lIdx)
							Y = lTemp \ ml_GRID_SIZE_WH
							X = lTemp - (Y * ml_GRID_SIZE_WH)
							moSlots(X, Y).yType = SlotType.eRear
						Next lIdx

						'Right
						For lIdx As Int32 = 0 To .RightLocs.GetUpperBound(0)
							lTemp = .RightLocs(lIdx)
							Y = lTemp \ ml_GRID_SIZE_WH
							X = lTemp - (Y * ml_GRID_SIZE_WH)
							moSlots(X, Y).yType = SlotType.eRight
						Next lIdx

						'All Arc
						For lIdx As Int32 = 0 To .AllArcLocs.GetUpperBound(0)
							lTemp = .AllArcLocs(lIdx)
							Y = lTemp \ ml_GRID_SIZE_WH
							X = lTemp - (Y * ml_GRID_SIZE_WH)
							moSlots(X, Y).yType = SlotType.eAllArc
						Next lIdx
					End With
				End If

			End If
		End Set
	End Property

	Public Sub New()
		ReDim moSlots(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1)
	End Sub

	Public Overrides Sub FillFromMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2	'first 2 bytes is my msg code

		ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If OwnerID = 0 Then OwnerID = glPlayerID
		ComponentDevelopmentPhase = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ErrorCodeReason = yData(lPos) : lPos += 1
		Researchers = yData(lPos) : lPos += 1
        MajorDesignFlaw = yData(lPos) : lPos += 1

        lSpecifiedHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedPower = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        blSpecifiedResCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        blSpecifiedResTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        blSpecifiedProdCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        blSpecifiedProdTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        lSpecifiedColonists = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedEnlisted = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedOfficers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedMin1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedMin2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedMin3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedMin4 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedMin5 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSpecifiedMin6 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		HullName = GetStringFromBytes(yData, lPos, 20)
		lPos += 20

		HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		StructuralMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		StructuralHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		yTypeID = yData(lPos) : lPos += 1
		ySubTypeID = yData(lPos) : lPos += 1
		ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		yChassisType = yData(lPos) : lPos += 1
		PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim X As Int32
		Dim Y As Int32
		For lIdx As Int32 = 0 To lCnt - 1
			X = yData(lPos) : lPos += 1
			Y = yData(lPos) : lPos += 1
			moSlots(X, Y).lConfig = CType(CInt(yData(lPos)), SlotConfig) : lPos += 1
			moSlots(X, Y).lGroupNum = CInt(yData(lPos)) : lPos += 1
		Next lIdx

	End Sub

	Public Function GetHullAllotment(ByVal yType As SlotType, ByVal lConfig As SlotConfig) As Single
		Dim lCnt As Int32 = 0
		Dim lTotal As Int32 = 0

		For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If moSlots(X, Y).yType <> SlotType.eUnused Then
					lTotal += 1
					If yType = SlotType.eNoChange Then
						If moSlots(X, Y).lConfig = lConfig Then lCnt += 1
					ElseIf moSlots(X, Y).yType = yType AndAlso moSlots(X, Y).lConfig = lConfig Then
						lCnt += 1
					End If
				End If
			Next X
		Next Y

		Dim fHullPerSlot As Single = CSng(HullSize / lTotal)
		Return fHullPerSlot * lCnt
	End Function

	Public Function GetWeaponHullAllotment(ByVal lGroupNum As Int32) As Single
		Dim lCnt As Int32 = 0
		Dim lTotal As Int32 = 0

		For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If moSlots(X, Y).yType <> SlotType.eUnused Then
					lTotal += 1
					If moSlots(X, Y).lConfig = SlotConfig.eWeapons AndAlso moSlots(X, Y).lGroupNum = lGroupNum Then
						lCnt += 1
					End If
				End If
			Next X
		Next Y

		Dim fHullPerSlot As Single = CSng(HullSize / lTotal)
		Return fHullPerSlot * lCnt
	End Function

	Public Function HasSlotConfig(ByVal lConfig As SlotConfig) As Boolean
		For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If moSlots(X, Y).lConfig = lConfig Then Return True
			Next X
		Next Y
		Return False
	End Function

	Public Function TotalSlots() As Int32
		Dim lCnt As Int32 = 0
		For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If moSlots(X, Y).yType <> SlotType.eUnused Then lCnt += 1
			Next X
		Next Y
		Return lCnt
	End Function

	Public Function GetHangarDoorDetails() As String
		Dim lGroup() As Int32 = Nothing
		Dim lGroupUB As Int32 = -1
		Dim lSize() As Int32 = Nothing
		Dim lSide1Cnt() As Int32 = Nothing
		Dim lSide2Cnt() As Int32 = Nothing
		Dim lSide3Cnt() As Int32 = Nothing
		Dim lSide4Cnt() As Int32 = Nothing

		For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If moSlots(X, Y).yType <> SlotType.eUnused AndAlso moSlots(X, Y).lConfig = SlotConfig.eHangarDoor Then
					Dim lIdx As Int32 = -1
					For lTempIdx As Int32 = 0 To lGroupUB
						If lGroup(lTempIdx) = moSlots(X, Y).lGroupNum Then
							lIdx = lTempIdx
							Exit For
						End If
					Next lTempIdx

					If lIdx = -1 Then
						lGroupUB += 1
						ReDim Preserve lGroup(lGroupUB)
						ReDim Preserve lSize(lGroupUB)
						ReDim Preserve lSide1Cnt(lGroupUB)
						ReDim Preserve lSide2Cnt(lGroupUB)
						ReDim Preserve lSide3Cnt(lGroupUB)
						ReDim Preserve lSide4Cnt(lGroupUB)
						lIdx = lGroupUB
					End If

					lGroup(lIdx) = moSlots(X, Y).lGroupNum
					lSize(lIdx) += 1

					Select Case moSlots(X, Y).yType
						Case SlotType.eFront
							lSide1Cnt(lIdx) += 1
						Case SlotType.eLeft
							lSide2Cnt(lIdx) += 1
						Case SlotType.eRear
							lSide3Cnt(lIdx) += 1
						Case SlotType.eRight
							lSide4Cnt(lIdx) += 1
					End Select
				End If
			Next X
		Next Y

		Dim fHullPerSlot As Single = CSng(Me.HullSize / TotalSlots())

		If lGroupUB <> -1 Then
			Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder
			Dim sSide As String = ""
			Dim lMax As Int32

            Dim fHangarCap As Single = GetHullAllotment(SlotType.eNoChange, SlotConfig.eHangar)

            Dim oModelDef As ModelDef = goModelDefs.GetModelDef(Me.ModelID)
            If oModelDef Is Nothing = False Then
                Select Case CType(oModelDef.lSpecialTraitID, elModelSpecialTrait)
                    Case elModelSpecialTrait.CargoAndHangar10
                        fHangarCap *= 1.1F
                    Case elModelSpecialTrait.CargoAndHangar3
                        fHangarCap *= 1.03F
                    Case elModelSpecialTrait.CargoAndHangar6
                        fHangarCap *= 1.06F
                    Case elModelSpecialTrait.Hangar10
                        fHangarCap *= 1.1F
                    Case elModelSpecialTrait.Hangar5
                        fHangarCap *= 1.05F
                End Select
            End If

            oSB.AppendLine("  Hanger Capacity" & Math.Floor(fHangarCap).ToString(" #,###").PadLeft(13, "."c))

			For X As Int32 = 0 To lGroupUB
				lMax = Int32.MinValue
				If lSide1Cnt(X) > lMax Then
					sSide = "Front" : lMax = lSide1Cnt(X)
				End If
				If lSide2Cnt(X) > lMax Then
                    sSide = "Left" : lMax = lSide2Cnt(X)
				End If
				If lSide3Cnt(X) > lMax Then
                    sSide = "Rear" : lMax = lSide3Cnt(X)
				End If
				If lSide4Cnt(X) > lMax Then
					sSide = "Right" : lMax = lSide4Cnt(X)
				End If

                Dim sLine As String = "    Door #" & lGroup(X).ToString & "(" & sSide & ")"
                oSB.AppendLine(sLine & CInt(lSize(X) * fHullPerSlot).ToString(" #,###").PadLeft(30 - sLine.Length, "."c))

            Next X
			Return oSB.ToString
		Else : Return ""
		End If
	End Function

	Public Overrides Function GetDesignFlawText() As String
		Return ""
	End Function

	Public Overrides Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lPos As Integer, ByVal yTechKnow As Byte, ByVal sName As String)
		HullName = sName

		Try
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
				'settings 1 or better
				HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				yTypeID = yData(lPos) : lPos += 1
				ySubTypeID = yData(lPos) : lPos += 1
				ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				yChassisType = yData(lPos) : lPos += 1
				PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				StructuralHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
				'settings 2 or better
				Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim X As Int32
				Dim Y As Int32
				For lIdx As Int32 = 0 To lCnt - 1
					X = yData(lPos) : lPos += 1
					Y = yData(lPos) : lPos += 1
					moSlots(X, Y).lConfig = CType(CInt(yData(lPos)), SlotConfig) : lPos += 1
					moSlots(X, Y).lGroupNum = CInt(yData(lPos)) : lPos += 1
				Next lIdx

				Me.oProductionCost = New ProductionCost()
				With Me.oProductionCost
					.ObjectID = Me.ObjectID
					.ObjTypeID = Me.ObjTypeID
					.ColonistCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.EnlistedCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.OfficerCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				End With
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
				'full knowledge
				StructuralMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			End If
		Catch
			'do nothing?
		End Try
	End Sub

	'Public Shared Function MaxWpnSlots(ByVal lHull As Int32) As Int32
	'	If lHull < 120 Then
	'		Return 2
	'	ElseIf lHull < 170 Then
	'		Return 5
	'	ElseIf lHull < 200 Then
	'		Return 3
	'	ElseIf lHull < 300 Then
	'		Return 1
	'	ElseIf lHull < 500 Then
	'		Return 4
	'	ElseIf lHull < 2000 Then
	'		Return 12
	'	ElseIf lHull < 8000 Then
	'		Return 8
	'	ElseIf lHull < 20000 Then
	'		Return 12
	'	ElseIf lHull < 40000 Then
	'		Return 15
	'	ElseIf lHull < 80000 Then
	'		Return 20
	'	ElseIf lHull < 250000 Then
	'		Return 20
	'	ElseIf lHull < 1000000 Then
	'		Return 25
	'	ElseIf lHull < 4000000 Then
	'		Return 20
	'	Else : Return 40
	'	End If
	'End Function
    Public Shared Function MaxWpnSlots(ByVal lTypeID As Int32, ByVal lSubTypeID As Int32) As Int32
        Dim lHullTypeID As eyHullType = GetHullTypeID(lTypeID, lSubTypeID)

        Dim lMaxGuns As Int32 = 1
        TechBuilderComputer.GetTypeValues(lHullTypeID, 0, lMaxGuns, 0, 0, 0, 0)
        Return lMaxGuns
        'Select Case lTypeID
        '    Case 0
        '        If lSubTypeID = 0 Then
        '            Return 20   'battleship
        '        ElseIf lSubTypeID = 2 Then
        '            Return 25   'battlecruiser
        '        End If
        '    Case 1
        '        If lSubTypeID = 0 Then Return 15
        '        If lSubTypeID = 1 Then Return 20
        '        If lSubTypeID = 2 Then Return 20
        '        If lSubTypeID = 3 Then Return 12
        '        Return 8
        '    Case 2
        '        If lSubTypeID = 9 Then
        '            Return 40       'space station
        '        Else : Return 20      'other facility
        '        End If
        '    Case 3
        '        If lSubTypeID = 0 Then Return 2 'small fighter
        '        If lSubTypeID = 1 Then Return 3 'medium fighter
        '        If lSubTypeID = 2 Then Return 4 'hvy fighter
        '    Case 4
        '        Return 5    'small vehicles
        '    Case 5
        '        Return 12   'tanks
        '    Case 6
        '        Return 4    'transports
        '    Case 7
        '        Return 1    'utility
        'End Select

        'Return 1
    End Function

	Public Overrides Function GetIntelReport(ByVal yLevel As PlayerTechKnowledge.KnowledgeType) As String
		Dim oSB As New System.Text.StringBuilder
		If yLevel > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
			'settings 1 or better
			oSB.AppendLine("Hull Size: " & HullSize.ToString("#,##0"))

			Dim oModelDef As ModelDef = goModelDefs.GetModelDef(ModelID)
			If oModelDef Is Nothing = False Then
				oSB.AppendLine("Chassis/Frame Name: " & oModelDef.FrameName)
			End If
			If (yChassisType And ChassisType.eGroundBased) <> 0 Then oSB.AppendLine("Ground-Based")
			If (yChassisType And ChassisType.eAtmospheric) <> 0 Then oSB.AppendLine("Atmospheric")
            If (yChassisType And ChassisType.eSpaceBased) <> 0 Then oSB.AppendLine("Space-Capable")
            If (yChassisType And ChassisType.eNavalBased) <> 0 Then oSB.AppendLine("Naval")
			oSB.AppendLine("Power Required: " & PowerRequired.ToString("#,##0"))
			oSB.AppendLine("Hitpoints: " & StructuralHitPoints.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'settings 2 or better
			Dim hulSlots As New UIHullSlots(goUILib)
			hulSlots.SetFromHullTech(Me)
			oSB.AppendLine(hulSlots.GetHullSummary(HullSize))
			hulSlots = Nothing
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
			'full knowledge 
			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = StructuralMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Structural Material: " & sName)
		End If
		Return oSB.ToString
    End Function

    Public Shared Function GetHullTypeID(ByVal plTypeID As Int32, ByVal plSubTypeID As Int32) As Base_Tech.eyHullType
        Select Case plTypeID
            Case 0  'capital
                If plSubTypeID = 2 Then Return eyHullType.BattleCruiser
                If plSubTypeID = 0 Then Return eyHullType.Battleship
            Case 1  'escort
                Select Case plSubTypeID
                    Case 0 : Return eyHullType.Corvette
                    Case 1 : Return eyHullType.Cruiser
                    Case 2 : Return eyHullType.Destroyer
                    Case 3 : Return eyHullType.Frigate
                    Case 4 : Return eyHullType.Escort
                End Select
            Case 2  'facility
                If plSubTypeID = 9 Then Return eyHullType.SpaceStation Else Return eyHullType.Facility
            Case 3  'fighter
                If plSubTypeID = 0 Then
                    Return eyHullType.LightFighter
                ElseIf plSubTypeID = 1 Then
                    Return eyHullType.MediumFighter
                ElseIf plSubTypeID = 2 Then
                    Return eyHullType.HeavyFighter
                End If
            Case 4  'small vehicle
                Return eyHullType.SmallVehicle
            Case 5  'tank
                Return eyHullType.Tank
            Case 6  'transport
                '????
                Return eyHullType.Utility
            Case 7  'utility vehicle
                Return eyHullType.Utility
            Case 8  'naval
                Select Case plSubTypeID
                    Case 0 : Return eyHullType.NavalBattleship
                    Case 1 : Return eyHullType.NavalCarrier
                    Case 2 : Return eyHullType.NavalCruiser
                    Case 3 : Return eyHullType.NavalDestroyer
                    Case 4 : Return eyHullType.NavalFrigate
                    Case 5 : Return eyHullType.NavalSub
                    Case 6 : Return eyHullType.Utility
                End Select
        End Select
        Return eyHullType.Utility
    End Function

    Public Shared Sub GetMinMaxHullOfHullType(ByVal plHullTypeID As Int32, ByRef lMinHull As Int32, ByRef lMaxHull As Int32)
        lMinHull = 0
        lMaxHull = 0

        If plHullTypeID = -1 Then Return

        lMinHull = GetMinHullForType(plHullTypeID)
        lMaxHull = GetMaxHullForType(plHullTypeID)
        'Select Case CType(plHullTypeID, eyHullType)
        '    Case eyHullType.BattleCruiser
        '        lMinHull = 110000 : lMaxHull = 253173
        '    Case eyHullType.Battleship
        '        lMinHull = 400000 : lMaxHull = 1100000
        '    Case eyHullType.Corvette
        '        lMinHull = 4500 : lMaxHull = 12500
        '    Case eyHullType.Cruiser
        '        lMinHull = 57000 : lMaxHull = 110000
        '    Case eyHullType.Destroyer
        '        lMinHull = 32000 : lMaxHull = 44000
        '    Case eyHullType.Escort
        '        lMinHull = 1200 : lMaxHull = 2700
        '    Case eyHullType.Facility
        '        lMinHull = 160 : lMaxHull = 1105000
        '    Case eyHullType.Frigate
        '        lMinHull = 5000 : lMaxHull = 28000
        '    Case eyHullType.HeavyFighter
        '        lMinHull = 135 : lMaxHull = 300
        '    Case eyHullType.LightFighter
        '        lMinHull = 40 : lMaxHull = 90
        '    Case eyHullType.MediumFighter
        '        lMinHull = 70 : lMaxHull = 140
        '    Case eyHullType.NavalBattleship
        '        lMinHull = 184000 : lMaxHull = 240000
        '    Case eyHullType.NavalCarrier
        '        lMinHull = 181000 : lMaxHull = 255000
        '    Case eyHullType.NavalCruiser
        '        lMinHull = 83000 : lMaxHull = 105000
        '    Case eyHullType.NavalDestroyer
        '        lMinHull = 31000 : lMaxHull = 50000
        '    Case eyHullType.NavalFrigate
        '        lMinHull = 12000 : lMaxHull = 20000
        '    Case eyHullType.NavalSub
        '        lMinHull = 42000 : lMaxHull = 127000
        '    Case eyHullType.SmallVehicle
        '        lMinHull = 75 : lMaxHull = 130
        '    Case eyHullType.SpaceStation
        '        lMinHull = 1500000 : lMaxHull = 11000000
        '    Case eyHullType.Tank
        '        lMinHull = 250 : lMaxHull = 620
        '    Case eyHullType.Utility
        '        lMinHull = 300 : lMaxHull = 36500
        'End Select
    End Sub
End Class

Public Structure HullSlot
    Public yType As SlotType
    Public lConfig As SlotConfig
    Public lGroupNum As Int32       'for separate groups

    Public bVerified As Boolean     'used during HasErrorList
    Public bFiltered As Boolean     'used during SetSelectingSlotConfig
End Structure

Public Enum SlotType As Byte
    eUnused = 0
    eFront = 1
    eLeft = 2
    eRear = 3
    eRight = 4
    eAllArc = 5

    eNoChange = 255
End Enum

Public Enum SlotConfig As Integer
    eArmorConfig = 0
    eCrewQuarters = 1
    eCargoBay = 2
    eRadar = 3
    eEngines = 4
    eHangar = 5
    eShields = 6
    eWeapons = 7
    eFuelBay = 8
    eHangarDoor = 9

    eConfigCnt      'must always be last
End Enum
