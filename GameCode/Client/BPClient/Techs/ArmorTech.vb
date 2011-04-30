Public Class ArmorTech
	Inherits Base_Tech

	Public sArmorName As String
	Public BeamResist As Byte
	Public ImpactResist As Byte
	Public PiercingResist As Byte
	Public MagneticResist As Byte
	Public ChemicalResist As Byte
	Public BurnResist As Byte
	Public RadarResist As Byte
	Public HullUsagePerPlate As Int32
	Public HPPerPlate As Int32

	Public lIntegrity As Int32

	Public OuterLayerMineralID As Int32 = -1
	Public MiddleLayerMineralID As Int32 = -1
	Public InnerLayerMineralID As Int32 = -1

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

			'Now, the alloy tech specifics...
			.sArmorName = GetStringFromBytes(yData, lPos, 20)
			lPos += 20

			.PiercingResist = yData(lPos) : lPos += 1
			.ImpactResist = yData(lPos) : lPos += 1
			.BeamResist = yData(lPos) : lPos += 1
			.MagneticResist = yData(lPos) : lPos += 1
			.BurnResist = yData(lPos) : lPos += 1
			.ChemicalResist = yData(lPos) : lPos += 1
			.RadarResist = yData(lPos) : lPos += 1
			.HPPerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.HullUsagePerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.OuterLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.MiddleLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.InnerLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.lIntegrity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		End With
	End Sub

	Public Overrides Function GetDesignFlawText() As String
		Select Case Me.MajorDesignFlaw
			Case 1
				Return "The Outer Layer Mineral was not fully studied!"
			Case 2
				Return "The Middle Layer Mineral was not fully studied!"
			Case 3
				Return "The Inner Layer Mineral was not fully studied!"
			Case 16
				Return "Getting the required Beam Resistance" & vbCrLf & "proved to be very difficult."
			Case 17
				Return "Getting the required Impact Resistance" & vbCrLf & "proved to be very difficult."
			Case 18
				Return "Getting the required Piercing Resistance" & vbCrLf & "proved to be very difficult."
			Case 19
				Return "Getting the required Magnetic Resistance" & vbCrLf & "proved to be very difficult."
			Case 20
				Return "Getting the required Chemical Resistance" & vbCrLf & "proved to be very difficult."
			Case 21
				Return "Getting the required Burn Resistance" & vbCrLf & "proved to be very difficult."
			Case 22
				Return "Getting the required Radar Resistance" & vbCrLf & "proved to be very difficult."
			Case Else
				Return ""
		End Select
	End Function

	Public Overrides Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lPos As Integer, ByVal yTechKnow As Byte, ByVal sName As String)
		sArmorName = sName

		Try
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
				'settings 1 or better
				HPPerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				HullUsagePerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
				'settings 2 or better
				PiercingResist = yData(lPos) : lPos += 1
				ImpactResist = yData(lPos) : lPos += 1
				BeamResist = yData(lPos) : lPos += 1
				MagneticResist = yData(lPos) : lPos += 1
				BurnResist = yData(lPos) : lPos += 1
				ChemicalResist = yData(lPos) : lPos += 1
				RadarResist = yData(lPos) : lPos += 1
				lIntegrity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
				OuterLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				MiddleLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                InnerLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
			'ylevel = PlayerTechKnowledge.KnowledgeType.eSettingsLevel1			
			oSB.AppendLine("Hitpoint Per Plate: " & HPPerPlate.ToString("#,##0"))
			oSB.AppendLine("Hull Per Plate: " & HullUsagePerPlate.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'ylevel = PlayerTechKnowledge.KnowledgeType.eSettingsLevel2
			oSB.AppendLine("Beam Resist: " & BeamResist.ToString("##0"))
			oSB.AppendLine("Burn Resist: " & BurnResist.ToString("##0"))
			oSB.AppendLine("Chemical Resist: " & ChemicalResist.ToString("##0"))
			oSB.AppendLine("Impact Resist: " & ImpactResist.ToString("##0"))
			oSB.AppendLine("Magnetic Resist: " & MagneticResist.ToString("##0"))
			oSB.AppendLine("Piercing Resist: " & PiercingResist.ToString("##0"))
			oSB.AppendLine("Radar Resist: " & RadarResist.ToString("##0"))
			oSB.AppendLine("Integrity: " & lIntegrity.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
			'ylevel = PlayerTechKnowledge.KnowledgeType.eFullKnowledge 
			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = OuterLayerMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Outer Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = MiddleLayerMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Middle Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = InnerLayerMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Inner Material: " & sName)
		End If
		Return oSB.ToString
	End Function
End Class
