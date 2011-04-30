Public Class ShieldTech
	Inherits Base_Tech

	Public sShieldName As String
	Public MaxHitPoints As Int32
	Public RechargeRate As Int32
	Public RechargeFreq As Int32
	Public lProjectionHullSize As Int32
	Public lCoilMineralID As Int32 = -1
	Public lAcceleratorMineralID As Int32 = -1
	Public lCasingMineralID As Int32 = -1
	Public PowerRequired As Int32
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

			.sShieldName = GetStringFromBytes(yData, lPos, 20)
			lPos += 20

			.MaxHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.RechargeRate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.RechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lProjectionHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lCasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lCoilMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lAcceleratorMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
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
				sResult &= "Coil's Density "
			Case eComponentDesignFlaw.eMat1_Prop2
				sResult &= "Coil's Superconductive Point "
			Case eComponentDesignFlaw.eMat1_Prop3
				sResult &= "Coil's Magnetic Production "
			Case eComponentDesignFlaw.eMat2_Prop1
				sResult &= "Accelerator's Quantum "
			Case eComponentDesignFlaw.eMat2_Prop2
				sResult &= "Accelerator's Superconductive Point "
			Case eComponentDesignFlaw.eMat2_Prop3
				sResult &= "Accelerator's Magnetic Reactance "
			Case eComponentDesignFlaw.eMat3_Prop1
				sResult &= "Casing's Density "
			Case eComponentDesignFlaw.eMat3_Prop2
				sResult &= "Casing's Thermal Conductance "
			Case eComponentDesignFlaw.eMat3_Prop3
				sResult &= "Casing's Temperature Sensitivity "
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
		sShieldName = sName

		Try
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
				'settings 1 or better
				PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                ColorValue = yData(lPos) : lPos += 1
                HullTypeID = yData(lPos) : lPos += 1
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
				'settings 2 or better
				MaxHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				RechargeRate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				RechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lProjectionHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
                lCoilMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lAcceleratorMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
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
            'ylevel = PlayerTechKnowledge.KnowledgeType.eSettingsLevel1		
            oSB.AppendLine("Shield for a " & Base_Tech.GetHullTypeName(HullTypeID))
			oSB.AppendLine("Power Required: " & PowerRequired.ToString("#,##0"))
			oSB.AppendLine("Hull Required: " & HullRequired.ToString("#,##0"))

			Select Case ColorValue
				Case 0
					oSB.AppendLine("Color: Teal")
				Case 1
					oSB.AppendLine("Color: White")
				Case 2
					oSB.AppendLine("Color: Yellow")
				Case 3
					oSB.AppendLine("Color: Orange")
				Case 4
					oSB.AppendLine("Color: Blue")
				Case 5
					oSB.AppendLine("Color: Red")
				Case 6
					oSB.AppendLine("Color: Purple")
			End Select
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'ylevel = PlayerTechKnowledge.KnowledgeType.eSettingsLevel2
			oSB.AppendLine("Max Hitpoints: " & MaxHitPoints.ToString("#,##0"))
			oSB.AppendLine("Recharge Rate: " & RechargeRate.ToString("#,##0"))
			oSB.AppendLine("Interval: " & (RechargeFreq / 30.0F).ToString("#,##0.##"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
			'ylevel = PlayerTechKnowledge.KnowledgeType.eFullKnowledge 
			oSB.AppendLine("Projection Size: " & lProjectionHullSize.ToString("#,##0"))

			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lCoilMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Coil Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lAcceleratorMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Accelerator Material: " & sName)
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
