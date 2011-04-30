Option Strict On

Public Class ProjectileWeaponTech
	Inherits WeaponTech

	'wpn specific values
	Public ProjectionType As Byte
	Public CartridgeSize As Int32
	Public PierceRatio As Byte
	Public ROF As Int16
	Public MaxRange As Int16
	Public PayloadType As Byte
	Public ExplosionRadius As Byte

	Public lMinPierce As Int32
	Public lMaxPierce As Int32
	Public lMinImpact As Int32
	Public lMaxImpact As Int32
	Public lMinPayload As Int32
	Public lMaxPayload As Int32

    'Public fAmmoSize As Single
    'Public fAmmoMin1Cost As Single = 0.0F
    'Public fAmmoMin2Cost As Single = 0.0F
    'Public fAmmoMin3Cost As Single = 0.0F
    'Public fAmmoMin4Cost As Single = 0.0F
    'Public fAmmoMin5Cost As Single = 0.0F

	Public BarrelMineralID As Int32
	Public CasingMineralID As Int32
	Public Payload1MineralID As Int32
	Public Payload2MineralID As Int32
	Public ProjectionMineralID As Int32

    Public Overrides Sub GetPrototypeBuilderAttributeString(ByRef oSB As System.Text.StringBuilder, ByVal bPointDefense As Boolean)
        oSB.AppendLine("    " & WeaponName)
        oSB.AppendLine("    RoF" & (ROF / 30.0F).ToString(" #,###.#0").PadLeft(23, "."c))
        Dim lMaxDmg As Int32 = (lMaxPierce + lMaxImpact + lMaxPayload)
        Dim lRange As Int32 = MaxRange
        If bPointDefense = True Then
            lMaxDmg = CInt(lMaxDmg * 0.5F)
            lRange = CInt(lRange * 0.25F)
        End If
        oSB.AppendLine("    Max Range" & lRange.ToString(" #,###").PadLeft(17, "."c))
        oSB.AppendLine("    Max Damage" & lMaxDmg.ToString(" #,###").PadLeft(16, "."c))
    End Sub

    Public Overrides Function GetPrototypeBuilderString(ByVal bPointDefense As Boolean) As String
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim sFinal As String = ""
        With Me

            Dim lDamage As Int32 = (.lMaxPierce + .lMaxImpact + .lMaxPayload)
            Dim lRange As Int32 = .MaxRange
            Dim lHull As Int32 = .HullRequired

            If bPointDefense = True Then
                lDamage = CInt(lDamage * 0.5F)
                lHull = CInt(lHull * 0.5F)
                lRange = CInt(lRange * 0.25F)
            End If

            oSB.AppendLine("Damage: " & lDamage)
            oSB.AppendLine("ROF: " & (.ROF / 30.0F).ToString("#,###.#0") & " s")
            oSB.AppendLine("Max Range: " & lRange)
            oSB.AppendLine("Explosion Radius: " & .ExplosionRadius)

            If bPointDefense = True Then
                oSB.AppendLine("Accuracy +25")
            End If

            oSB.AppendLine("Hull Usage: " & lHull)
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
			.ProjectionType = yData(lPos) : lPos += 1
			.CartridgeSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PierceRatio = yData(lPos) : lPos += 1
			.ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.MaxRange = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.PayloadType = yData(lPos) : lPos += 1
			.ExplosionRadius = yData(lPos) : lPos += 1

			.BarrelMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.CasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.Payload1MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.Payload2MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ProjectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.lMinPierce = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lMaxPierce = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lMinImpact = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lMaxImpact = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lMinPayload = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lMaxPayload = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            '.fAmmoSize = System.BitConverter.ToSingle(yData, lPos) : lPos += 4

            '.blAmmoProdCredits = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            '.blAmmoProdPoints = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            '         .fAmmoMin1Cost = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
            '.fAmmoMin2Cost = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
            '.fAmmoMin3Cost = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
            '.fAmmoMin4Cost = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
            '.fAmmoMin5Cost = System.BitConverter.ToSingle(yData, lPos) : lPos += 4

			FillWeaponDefFromMsg(yData, lPos)
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
				sResult &= "Barrel's Temperature Sensitivity "
			Case eComponentDesignFlaw.eMat1_Prop2
				sResult &= "Barrel's Thermal Conductance "
			Case eComponentDesignFlaw.eMat1_Prop3
				sResult &= "Barrel's Thermal Expansion "
			Case eComponentDesignFlaw.eMat2_Prop1
				sResult &= "Casing's Temperature Sensitivity "
			Case eComponentDesignFlaw.eMat2_Prop2
				sResult &= "Casing's Density "
			Case eComponentDesignFlaw.eMat2_Prop3
				sResult &= "Casing's Thermal Conductance "
			Case eComponentDesignFlaw.eMat3_Prop1
				sResult &= "Payload 1's Malleable "
			Case eComponentDesignFlaw.eMat3_Prop2
				sResult &= "Payload 1's Hardness "
			Case eComponentDesignFlaw.eMat3_Prop3
				sResult &= "Payload 1's Density "
			Case eComponentDesignFlaw.eMat4_Prop1
				sResult &= "Payload 2's Properties "
			Case eComponentDesignFlaw.eMat5_Prop1
				sResult &= "Projection's properties "
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
				ProjectionType = yData(lPos) : lPos += 1
				CartridgeSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				PierceRatio = yData(lPos) : lPos += 1
				ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				MaxRange = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				PayloadType = yData(lPos) : lPos += 1
				ExplosionRadius = yData(lPos) : lPos += 1

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
				BarrelMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				CasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Payload1MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Payload2MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				ProjectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

				lMinPierce = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lMaxPierce = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lMinImpact = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lMaxImpact = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lMinPayload = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lMaxPayload = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'fAmmoSize = System.BitConverter.ToSingle(yData, lPos) : lPos += 4

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
            oSB.AppendLine("Projectile for a " & Base_Tech.GetHullTypeName(HullTypeID))
			oSB.AppendLine("Power Required: " & PowerRequired.ToString("#,##0"))
			oSB.AppendLine("Hull Required: " & HullRequired.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'settings 2 or better
			If ProjectionType = 0 Then oSB.AppendLine("Explosive Projection") Else oSB.AppendLine("Magnetic Projection")
			oSB.AppendLine("Cartridge Size: " & CartridgeSize.ToString("#,##0"))
			oSB.AppendLine("Pierce Ratio: " & PierceRatio.ToString("#,##0"))
			oSB.AppendLine("Rate of Fire: " & (ROF / 30.0F).ToString("#,##0.##"))
			oSB.AppendLine("Max Range: " & MaxRange.ToString("#,##0"))

			Select Case PayloadType
				Case 0
					oSB.AppendLine("Payload Type: No additional")
				Case 1
					oSB.AppendLine("Payload Type: Explosive")
				Case 2
					oSB.AppendLine("Payload Type: Chemical")
				Case 3
					oSB.AppendLine("Payload Type: Magnetic")
			End Select
			oSB.AppendLine("Explosion Radius: " & ExplosionRadius.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then

			oSB.AppendLine("Pierce Damage: " & lMinPierce.ToString & " to " & lMaxPierce.ToString)
			oSB.AppendLine("Impact Damage: " & lMinImpact.ToString & " to " & lMaxImpact.ToString)
			oSB.AppendLine("Payload Damage: " & lMinPayload.ToString & " to " & lMaxPayload.ToString)
 
 
			'full knowledge
			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = BarrelMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Barrel Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = CasingMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Casing Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = Payload1MineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Payload 1 Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = Payload2MineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Payload 2 Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = ProjectionMineralID Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Projection Material: " & sName)
		End If
		Return oSB.ToString
	End Function
End Class
