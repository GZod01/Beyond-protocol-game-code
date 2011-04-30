Public Class RadarTech
	Inherits Base_Tech

	Public sRadarName As String
	Public WeaponAcc As Byte
	Public ScanResolution As Byte
	Public OptimumRange As Byte
	Public MaximumRange As Byte
	Public DisruptionResistance As Byte
	Public lEmitterMineralID As Int32 = -1
	Public lDetectionMineralID As Int32 = -1
	Public lCollectionMineralID As Int32 = -1
	Public lCasingMineralID As Int32 = -1
	Public PowerRequired As Int32
	Public HullRequired As Int32

    'Public yRadarType As Byte
	Public JamImmunity As Byte
	Public JamStrength As Byte
	Public JamTargets As Byte
	Public JamEffect As Byte

    Public HullTypeID As Byte

	Public Overrides Sub FillFromMsg(ByRef yData() As Byte)
		With Me
			Dim lPos As Int32 = 2
			.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If .OwnerID = 0 Then .OwnerID = glPlayerID
			.ComponentDevelopmentPhase = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ErrorCodeReason = yData(lPos) : lPos += 1
			.Researchers = yData(lPos) : lPos += 1
			.MajorDesignFlaw = yData(lPos) : lPos += 1

            .lSpecifiedHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedPower = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .blSpecifiedResCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedResTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedProdCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedProdTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .lSpecifiedColonists = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedEnlisted = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedOfficers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin4 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin5 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin6 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.sRadarName = GetStringFromBytes(yData, lPos, 20)
			lPos += 20

			.WeaponAcc = yData(lPos) : lPos += 1
			.ScanResolution = yData(lPos) : lPos += 1
			.OptimumRange = yData(lPos) : lPos += 1
			.MaximumRange = yData(lPos) : lPos += 1
			.DisruptionResistance = yData(lPos) : lPos += 1
			.lCollectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lCasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lEmitterMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lDetectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .HullTypeID = yData(lPos) : lPos += 1
			.JamImmunity = yData(lPos) : lPos += 1
			.JamStrength = yData(lPos) : lPos += 1
			.JamTargets = yData(lPos) : lPos += 1
            .JamEffect = yData(lPos) : lPos += 1
        End With
	End Sub

	Public Overrides Function GetDesignFlawText() As String
		Dim sResult As String = "It would be better if the " & vbCrLf

		Dim yTemp As Byte = MajorDesignFlaw
		Dim bHigher As Boolean = False
		If (MajorDesignFlaw And eComponentDesignFlaw.eShift_Not_Study) <> 0 Then
			sResult = ""
			yTemp = CByte(yTemp - eComponentDesignFlaw.eShift_Not_Study)
		End If
		If (MajorDesignFlaw And eComponentDesignFlaw.eShift_Should_Be_Higher) <> 0 Then
			yTemp = CByte(yTemp - eComponentDesignFlaw.eShift_Should_Be_Higher)
			bHigher = True
		End If
		Dim bGoodDesign As Boolean = False
		If (MajorDesignFlaw And eComponentDesignFlaw.eGoodDesign) <> 0 Then
			yTemp = CByte(yTemp - eComponentDesignFlaw.eGoodDesign)
			sResult = "A good design, perhaps if the " & vbCrLf
			bGoodDesign = True
		End If

		Select Case CType(yTemp, eComponentDesignFlaw)
			Case eComponentDesignFlaw.eMat1_Prop1
				sResult &= "Emitter's Electrical Resistance "
			Case eComponentDesignFlaw.eMat1_Prop2
				sResult &= "Emitter's Magnetic Production "
			Case eComponentDesignFlaw.eMat1_Prop3
				sResult &= "Emitter's Magnetic Reactance "
			Case eComponentDesignFlaw.eMat2_Prop1
				sResult &= "Detector's Magnetic Reactance "
			Case eComponentDesignFlaw.eMat2_Prop2
				sResult &= "Detector's Magnetic Production "
			Case eComponentDesignFlaw.eMat2_Prop3
				sResult &= "Detector's Superconductive Point "
			Case eComponentDesignFlaw.eMat3_Prop1
				sResult &= "Collector's Electrical Resistance "
			Case eComponentDesignFlaw.eMat3_Prop2
				sResult &= "Collector's Malleable "
			Case eComponentDesignFlaw.eMat3_Prop3
				sResult &= "Collector's Magnetic Production "
			Case eComponentDesignFlaw.eMat4_Prop1
				sResult &= "Casing's Thermal Conductance "
			Case eComponentDesignFlaw.eMat4_Prop2
				sResult &= "Casing's Thermal Expansion "
			Case eComponentDesignFlaw.eMat4_Prop3
				sResult &= "Casing's Density "
			Case eComponentDesignFlaw.eMat4_Prop4
				sResult &= "Casing's Magnetic Production "
			Case eComponentDesignFlaw.eMat4_Prop5
				sResult &= "Casing's Magnetic Reactance "
			Case Else
				'sResult &= "overall design "
                Return "Your scientists have finished designing the component"
		End Select

		If (MajorDesignFlaw And eComponentDesignFlaw.eShift_Not_Study) <> 0 Then
			sResult &= vbCrLf & "not being fully studied gave us difficulties"
		Else
			If bGoodDesign = False Then
				If bHigher = True Then
					sResult &= "was higher"
				Else : sResult &= "was lower"
				End If
			Else
				If bHigher = True Then
					sResult &= "was slightly higher"
				Else : sResult &= "was slightly lower"
				End If
			End If
		End If

		Return sResult
	End Function

	Public Overrides Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lPos As Integer, ByVal yTechKnow As Byte, ByVal sName As String)
		sRadarName = sName

		Try
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
				'settings 1 or better
                HullTypeID = yData(lPos) : lPos += 1
				PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
				'settings 2 or better
				WeaponAcc = yData(lPos) : lPos += 1
				ScanResolution = yData(lPos) : lPos += 1
				OptimumRange = yData(lPos) : lPos += 1
				MaximumRange = yData(lPos) : lPos += 1
				DisruptionResistance = yData(lPos) : lPos += 1
				JamImmunity = yData(lPos) : lPos += 1
				JamStrength = yData(lPos) : lPos += 1
				JamTargets = yData(lPos) : lPos += 1
				JamEffect = yData(lPos) : lPos += 1

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
				lEmitterMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lDetectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lCollectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lCasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
			End If
		Catch
			'do nothing?
		End Try

	End Sub

	Public Overrides Function GetIntelReport(ByVal yLevel As PlayerTechKnowledge.KnowledgeType) As String
		Dim oSB As New System.Text.StringBuilder
		If yLevel > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
			'settings 1 or better
            'If HullTypeID = 0 Then oSB.AppendLine("Passive Radar") Else oSB.AppendLine("Active Radar")
            oSB.AppendLine(GetHullTypeName(HullTypeID) & " Radar")
			oSB.AppendLine("Power Required: " & PowerRequired.ToString("#,##0"))
			oSB.AppendLine("Hull Required: " & HullRequired.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'settings 2 or better
			oSB.AppendLine("Weapon Accuracy: " & WeaponAcc.ToString("#,##0"))
			oSB.AppendLine("Scan Resolution: " & ScanResolution.ToString("#,##0"))
			oSB.AppendLine("Optimum Range: " & OptimumRange.ToString("#,##0"))
			oSB.AppendLine("Maximum Range: " & MaximumRange.ToString("#,##0"))
			oSB.AppendLine("Disrupt Resist: " & DisruptionResistance.ToString("#,##0"))

			Dim lJamImmunes As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eJammingImmunityAvailable)
			If lJamImmunes >= JamImmunity Then
				Select Case JamImmunity
					Case 0
						oSB.AppendLine("Jamming Immunity: None")
					Case 1
						oSB.AppendLine("Jamming Immunity: System Degradation")
					Case 2
						oSB.AppendLine("Jamming Immunity: System Interference")
					Case 3
						oSB.AppendLine("Jamming Immunity: System Clutter")
					Case 4
						oSB.AppendLine("Jamming Immunity: Anti-Jamming")
					Case 5
						oSB.AppendLine("Jamming Immunity: Decoy Clutter")
				End Select 
			Else : oSB.AppendLine("Jamming Immunity: Unknown")
			End If

			Dim lJamEffects As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eJammingEffectAvailable)
			If lJamEffects >= JamEffect Then
				Select Case JamEffect
					Case 0
						oSB.AppendLine("Jamming Effect: None")
					Case 1
						oSB.AppendLine("Jamming Effect: System Degradation")
					Case 2
						oSB.AppendLine("Jamming Effect: System Interference")
					Case 3
						oSB.AppendLine("Jamming Effect: System Clutter")
					Case 4
						oSB.AppendLine("Jamming Effect: Anti-Jamming")
					Case 5
						oSB.AppendLine("Jamming Effect: Decoy Clutter")
				End Select
			Else : oSB.AppendLine("Jamming Effect: Unknown")
			End If

			oSB.AppendLine("Jamming Strength: " & JamStrength.ToString("#,##0"))
			oSB.AppendLine("Jamming Targets: " & JamTargets.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
 
			'full knowledge
			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lEmitterMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Emitter Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lDetectionMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Detection Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lCollectionMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Collection Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lCasingMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Casing Material: " & sName)
		End If
		Return oSB.ToString
	End Function
End Class
