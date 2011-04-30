Public Class EngineTech
	Inherits Base_Tech

    Public sEngineName As String
	Public Thrust As Int32
	Public Maneuver As Byte
	Public Speed As Byte
	Public PowerProd As Int32
	Public lStructuralBodyMineralID As Int32 = -1
	Public lStructuralFrameMineralID As Int32 = -1
	Public lStructuralMeldMineralID As Int32 = -1
	Public lDriveBodyMineralID As Int32 = -1
	Public lDriveFrameMineralID As Int32 = -1
	Public lDriveMeldMineralID As Int32 = -1
	'Public lFuelCompositionMineralID As Int32 = -1
	'Public lFuelCatalystMineralID As Int32 = -1

	Public HullRequired As Int32

    Public ColorValue As Byte = 0
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
            'Return
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
			.sEngineName = GetStringFromBytes(yData, lPos, 20)
			lPos += 20

			.Thrust = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.Maneuver = yData(lPos) : lPos += 1
			.Speed = yData(lPos) : lPos += 1
			.PowerProd = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lStructuralBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lStructuralFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lStructuralMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lDriveBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lDriveFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lDriveMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			'.lFuelCompositionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			'.lFuelCatalystMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ColorValue = yData(lPos) : lPos += 1
            .HullTypeID = yData(lPos) : lPos += 1
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
				sResult &= "Structural Body's Magnetic Production "
			Case eComponentDesignFlaw.eMat1_Prop2
				sResult &= "Structural Body's Density "
			Case eComponentDesignFlaw.eMat1_Prop3
				sResult &= "Structural Body's Melting Point "
			Case eComponentDesignFlaw.eMat2_Prop1
				sResult &= "Structural Frame's Malleable "
			Case eComponentDesignFlaw.eMat2_Prop2
				sResult &= "Structural Frame's Hardness "
			Case eComponentDesignFlaw.eMat2_Prop3
				sResult &= "Structural Frame's Compress "
			Case eComponentDesignFlaw.eMat3_Prop1
				sResult &= "Structural Meld's Malleable "
			Case eComponentDesignFlaw.eMat3_Prop2
				sResult &= "Structural Meld's Hardness "
			Case eComponentDesignFlaw.eMat3_Prop3
				sResult &= "Structural Meld's Melting Point "
			Case eComponentDesignFlaw.eMat4_Prop1
				sResult &= "Drive Body's Magnetic Reactance "
			Case eComponentDesignFlaw.eMat4_Prop2
				sResult &= "Drive Body's Compressibility "
			Case eComponentDesignFlaw.eMat4_Prop3
				sResult &= "Drive Body's Density "
			Case eComponentDesignFlaw.eMat5_Prop1
				sResult &= "Drive Frame's Superconductive Point "
			Case eComponentDesignFlaw.eMat5_Prop2
				sResult &= "Drive Frame's Hardness "
			Case eComponentDesignFlaw.eMat6_Prop1
				sResult &= "Drive Meld's Chemical Reactance "
			Case eComponentDesignFlaw.eMat6_Prop2
				sResult &= "Drive Meld's Malleable "
			Case eComponentDesignFlaw.eMat6_Prop3
				sResult &= "Drive Meld's Combustiveness "
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
		sEngineName = sName

		Try
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
				'settings 1 or better
				HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                ColorValue = yData(lPos) : lPos += 1
                HullTypeID = yData(lPos) : lPos += 1
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
				'settings 2 or better
				Thrust = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Maneuver = yData(lPos) : lPos += 1
				Speed = yData(lPos) : lPos += 1
				PowerProd = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
				lStructuralBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lStructuralFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lStructuralMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lDriveBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lDriveFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lDriveMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
            oSB.AppendLine("Engine for a " & Base_Tech.GetHullTypeName(HullTypeID))
			oSB.AppendLine("Hull Required: " & HullRequired.ToString("#,##0"))

			Select Case ColorValue
				Case 0
					oSB.AppendLine("Color: Light Blue")
				Case 1
					oSB.AppendLine("Color: Bright Green")
				Case 2
					oSB.AppendLine("Color: Orange")
				Case 3
					oSB.AppendLine("Color: Blue")
				Case 4
					oSB.AppendLine("Color: Purple")
				Case 5
					oSB.AppendLine("Color: Dark Blue")
				Case 6
					oSB.AppendLine("Color: Red")
				Case 7
					oSB.AppendLine("Color: Yellow")
				Case 8
					oSB.AppendLine("Color: Dark Green")
			End Select 
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'settings 2 or better
			oSB.AppendLine("Thrust: " & Thrust.ToString("#,##0"))
			oSB.AppendLine("Power: " & PowerProd.ToString("#,##0"))
			oSB.AppendLine("Max Speed: " & Speed.ToString("#,##0"))
			oSB.AppendLine("Maneuver: " & Maneuver.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
			'full knowledge 
			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lStructuralBodyMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Structural Body: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lStructuralFrameMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Structural Frame: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lStructuralMeldMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Structural Meld: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lDriveBodyMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Drive Body: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lDriveFrameMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Drive Frame: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lDriveMeldMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Drive Meld: " & sName)
		End If
		Return oSB.ToString
	End Function
End Class
