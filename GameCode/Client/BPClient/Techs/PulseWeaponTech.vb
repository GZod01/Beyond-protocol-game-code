Option Strict On

Public Class PulseWeaponTech
	Inherits WeaponTech

	Public InputEnergy As Int32
	Public CompressFactor As Single
	Public MaxRange As Byte
	Public ROF As Int16
	Public ScatterRadius As Byte

	Public CoilMatID As Int32
	Public AcceleratorMatID As Int32
	Public CasingMatID As Int32
	Public FocuserMatID As Int32
	Public CompressChmbrMatID As Int32

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
			Case eComponentDesignFlaw.eMat1_Prop4
				sResult &= "Coil's Magnetic Reactance "
			Case eComponentDesignFlaw.eMat2_Prop1
				sResult &= "Accelerator's Quantum "
			Case eComponentDesignFlaw.eMat2_Prop2
				sResult &= "Accelerator's Superconductive Point "
			Case eComponentDesignFlaw.eMat2_Prop3
				sResult &= "Accelerator's Magnetic Reactance "
			Case eComponentDesignFlaw.eMat2_Prop4
				sResult &= "Accelerator's Magnetic Production "
			Case eComponentDesignFlaw.eMat3_Prop1
				sResult &= "Casing's Density "
			Case eComponentDesignFlaw.eMat3_Prop2
				sResult &= "Casing's Thermal Conductance "
			Case eComponentDesignFlaw.eMat3_Prop3
				sResult &= "Casing's Temperature Sensitivity "
			Case eComponentDesignFlaw.eMat3_Prop4
				sResult &= "Casing's Malleable "
			Case eComponentDesignFlaw.eMat4_Prop1
				sResult &= "Focuser's Refraction "
			Case eComponentDesignFlaw.eMat4_Prop2
				sResult &= "Focuser's Quantum "
			Case eComponentDesignFlaw.eMat4_Prop3
				sResult &= "Focuser's Temperature Sensitivity "
			Case eComponentDesignFlaw.eMat4_Prop4
				sResult &= "Focuser's Thermal Expansion "
			Case eComponentDesignFlaw.eMat5_Prop1
				sResult &= "Compression Chamber's Material "
			Case eComponentDesignFlaw.eMat5_Prop4
				sResult &= "Compression Chamber's Electrical Resistance "
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

    Public Overrides Sub GetPrototypeBuilderAttributeString(ByRef oSB As System.Text.StringBuilder, ByVal bPointDefense As Boolean)
        oSB.AppendLine("    " & WeaponName)
        oSB.AppendLine("    RoF" & (ROF / 30.0F).ToString(" #,##0.#0").PadLeft(23, "."c))
        Dim lRange As Int32 = CInt(MaxRange)
        Dim lMaxDmg As Int32 = CInt(InputEnergy * CompressFactor / 10)
        If bPointDefense = True Then
            lRange = CInt(lRange * 0.25F)
            lMaxDmg = CInt(lMaxDmg * 0.5F)
        End If
        oSB.AppendLine("    Max Range" & lRange.ToString(" #,###").PadLeft(17, "."c))
        oSB.AppendLine("    Max Damage" & lMaxDmg.ToString(" #,###").PadLeft(16, "."c))
    End Sub

    Public Overrides Function GetPrototypeBuilderString(ByVal bPointDefense As Boolean) As String
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim sFinal As String = ""
        With Me

            Dim lRange As Int32 = .MaxRange
            Dim lHull As Int32 = .HullRequired
            If bPointDefense = True Then
                lRange = CInt(lRange * 0.25F)
                lHull = CInt(lHull * 0.5F)
                oSB.AppendLine("50% Damage")
                oSB.AppendLine("+25 Accuracy")
            End If

            oSB.AppendLine("Max Range: " & .MaxRange)
            oSB.AppendLine("ROF: " & (.ROF / 30.0F).ToString("#,###.#0") & " s")
            oSB.AppendLine("Hull Usage: " & .HullRequired)
            oSB.AppendLine("Power Required: " & .PowerRequired & vbCrLf)

            oSB.AppendLine("COSTS:")
            With .oProductionCost
                If .CreditCost <> 0 Then oSB.AppendLine("  Credits: " & .CreditCost)
                If .ColonistCost <> 0 Then oSB.AppendLine("  Colonists: " & .ColonistCost)
                If .EnlistedCost <> 0 Then oSB.AppendLine("  Enlisted: " & .EnlistedCost)
                If .OfficerCost <> 0 Then oSB.AppendLine("  Officers: " & .OfficerCost)
            End With

            If .oProductionCost.ItemCostUB > -1 Then oSB.AppendLine(vbCrLf & "MATERIALS")
            For Y As Int32 = 0 To .oProductionCost.ItemCostUB
                oSB.AppendLine("  " & .oProductionCost.ItemCosts(Y).GetItemName & ": " & .oProductionCost.ItemCosts(Y).QuantityNeeded)
            Next Y
            sFinal = oSB.ToString
        End With

        oSB = Nothing
        Return sFinal
    End Function

	Protected Overrides Sub PopulateFromMsg(ByRef yData() As Byte, ByVal lPos As Integer)
		With Me
			.InputEnergy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.CompressFactor = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
			.MaxRange = yData(lPos) : lPos += 1
			.ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.ScatterRadius = yData(lPos) : lPos += 1

			.CoilMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.AcceleratorMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.CasingMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.FocuserMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.CompressChmbrMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			FillWeaponDefFromMsg(yData, lPos)
		End With
	End Sub

	Public Overrides Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lPos As Integer, ByVal yTechKnow As Byte, ByVal sName As String)
		WeaponName = sName

		Try
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
				'settings 1 or better
				WeaponTypeID = CType(yData(lPos), WeaponType) : lPos += 1
				PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                HullTypeID = yData(lPos) : lPos += 1
			End If
			If yTechKnow > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
				'settings 2 or better
				InputEnergy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				CompressFactor = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
				MaxRange = yData(lPos) : lPos += 1
				ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				ScatterRadius = yData(lPos) : lPos += 1

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
				CoilMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				AcceleratorMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				CasingMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				FocuserMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                CompressChmbrMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
            oSB.AppendLine("Pulse Beam for a " & Base_Tech.GetHullTypeName(HullTypeID))
			oSB.AppendLine("Power Required: " & PowerRequired.ToString("#,##0"))
			oSB.AppendLine("Hull Required: " & HullRequired.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'settings 2 or better
			oSB.AppendLine("Input Energy: " & InputEnergy.ToString("#,##0"))
			oSB.AppendLine("Compression: " & CompressFactor.ToString("#,##0.##"))
			oSB.AppendLine("Max Range: " & MaxRange.ToString("#,##0"))
			oSB.AppendLine("Rate of Fire: " & (ROF / 30.0F).ToString("#,##0.##"))
			oSB.AppendLine("Scatter: " & ScatterRadius.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
			'full knowledge
			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = CoilMatID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Coil Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = AcceleratorMatID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Accelerator Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = CasingMatID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Casing Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = FocuserMatID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Focuser Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = CompressChmbrMatID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Chamber Material: " & sName)
		End If
		Return oSB.ToString
	End Function
End Class
