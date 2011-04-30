Option Strict On

Public Class MissileWeaponTech
	Inherits WeaponTech

	'Player Enter Values
	Public MaximumDamage As Int32
	Public MissileHullSize As Int32
	Public MaxSpeed As Byte
	Public Maneuver As Byte
	Public ROF As Int16
	Public FlightTime As Int16
	Public HomingAccuracy As Byte
	Public PayloadType As Byte
	Public ExplosionRadius As Byte
	Public StructuralHP As Int32

	Public BodyMaterialID As Int32
	Public NoseMaterialID As Int32
	Public FlapsMaterialID As Int32
	Public FuelMaterialID As Int32
	Public PayloadMaterialID As Int32

    'Public blAmmoCostCredits As Int64 = 0
    'Public blAmmoCostPoints As Int64 = 0
    'Public fAmmoMin1Cost As Single = 0.0F
    'Public fAmmoMin2Cost As Single = 0.0F
    'Public fAmmoMin3Cost As Single = 0.0F
    'Public fAmmoMin4Cost As Single = 0.0F
    'Public fAmmoMin5Cost As Single = 0.0F

    Public Overrides Sub GetPrototypeBuilderAttributeString(ByRef oSB As System.Text.StringBuilder, ByVal bPointDefense As Boolean)
        oSB.AppendLine("    " & WeaponName)

        If ROF < 1 Then
            oSB.AppendLine("    Single-Shot Missile")
        Else : oSB.AppendLine("    RoF" & (ROF / 30.0F).ToString(" #,##0.#0").PadLeft(23, "."c))
        End If
        Dim lDamage As Int32 = MaximumDamage
        If bPointDefense = True Then lDamage = CInt(lDamage * 0.5F)
        oSB.AppendLine("    Max Dmg" & lDamage.ToString(" #,###").PadLeft(19, "."c))
    End Sub

    Public Overrides Function GetPrototypeBuilderString(ByVal bPointDefense As Boolean) As String
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim sFinal As String = ""
        With Me
            oSB.AppendLine("Damage: " & .MaximumDamage)
            If .ROF = -1 Then
                oSB.AppendLine("Single Shot")
            Else : oSB.AppendLine("ROF: " & (.ROF / 30.0F).ToString("#,###.#0") & " s")
            End If
            oSB.AppendLine("Speed: " & .MaxSpeed)
            oSB.AppendLine("Maneuver: " & .Maneuver)
            oSB.AppendLine("Range: " & .FlightTime)

            oSB.AppendLine("Hull Usage: " & .HullRequired)
            oSB.AppendLine("Missile Size: " & .MissileHullSize)
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
			.ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.MaximumDamage = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.MissileHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.MaxSpeed = yData(lPos) : lPos += 1
			.Maneuver = yData(lPos) : lPos += 1
			.FlightTime = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.HomingAccuracy = yData(lPos) : lPos += 1
			.PayloadType = yData(lPos) : lPos += 1
			.ExplosionRadius = yData(lPos) : lPos += 1
			.StructuralHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.BodyMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.NoseMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.FlapsMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.FuelMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PayloadMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            '.blAmmoCostCredits = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            '.blAmmoCostPoints = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            '.fAmmoMin1Cost = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
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
		If (yTemp And eComponentDesignFlaw.eShift_Not_Study) <> 0 Then
			sResult = ""
			yTemp = CByte(yTemp - eComponentDesignFlaw.eShift_Not_Study)
		End If
		If (MajorDesignFlaw And eComponentDesignFlaw.eShift_Should_Be_Higher) <> 0 Then
			yTemp = CByte(yTemp - eComponentDesignFlaw.eShift_Should_Be_Higher)
			bHigher = True
		End If

		Select Case CType(yTemp, Base_Tech.eComponentDesignFlaw)
			Case eComponentDesignFlaw.eMat1_Prop1
				sResult &= "Body's Hardness "
			Case eComponentDesignFlaw.eMat1_Prop2
				sResult &= "Body's Density "
			Case eComponentDesignFlaw.eMat1_Prop3
				sResult &= "Body's Malleable "
			Case eComponentDesignFlaw.eMat2_Prop1
				sResult &= "Nose's Hardness "
			Case eComponentDesignFlaw.eMat2_Prop2
				sResult &= "Nose's Compress "
			Case eComponentDesignFlaw.eMat3_Prop1
				sResult &= "Flap's Density "
			Case eComponentDesignFlaw.eMat3_Prop2
				sResult &= "Flap's Thermal Expansion "
			Case eComponentDesignFlaw.eMat3_Prop3
				sResult &= "Flap's Temperature Sensitivity "
			Case eComponentDesignFlaw.eMat4_Prop1
				sResult &= "Fuel's Chemical Reactance "
			Case eComponentDesignFlaw.eMat4_Prop2
				sResult &= "Fuel's Boiling Point "
			Case eComponentDesignFlaw.eMat4_Prop3
				sResult &= "Fuel's Density "
			Case eComponentDesignFlaw.eMat5_Prop1
				sResult &= "Payload's properties "
			Case Else
				'sResult &= "overall design "
                Return "Your scientists have finished designing the component"
		End Select

		If (MajorDesignFlaw And eComponentDesignFlaw.eShift_Not_Study) <> 0 Then
			sResult &= vbCrLf & "not being fully studied gave us difficulties"
		Else
			If bHigher = True Then
				sResult &= "was higher"
			Else : sResult &= "was lower"
			End If
			'sResult &= "to work"
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
				ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				MaximumDamage = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				MissileHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				MaxSpeed = yData(lPos) : lPos += 1
				Maneuver = yData(lPos) : lPos += 1
				FlightTime = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				HomingAccuracy = yData(lPos) : lPos += 1
				PayloadType = yData(lPos) : lPos += 1
				ExplosionRadius = yData(lPos) : lPos += 1
				StructuralHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
				BodyMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				NoseMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				FlapsMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				FuelMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                PayloadMaterialID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
            oSB.AppendLine("Missile weapon for a " & Base_Tech.GetHullTypeName(HullTypeID))
            oSB.AppendLine("Power Required: " & PowerRequired.ToString("#,##0"))
            oSB.AppendLine("Hull Required: " & HullRequired.ToString("#,##0"))
        End If
        If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            'settings 2 or better
 
            oSB.AppendLine("Rate of Fire: " & (ROF / 30.0F).ToString("#,##0.##"))
            oSB.AppendLine("Max Damage: " & MaximumDamage.ToString("#,##0.##"))
            oSB.AppendLine("Missile Size: " & MissileHullSize.ToString("#,##0.##"))
            oSB.AppendLine("Max Speed: " & MaxSpeed.ToString)
            oSB.AppendLine("Maneuver: " & Maneuver.ToString)
            oSB.AppendLine("Range: " & FlightTime.ToString)
            oSB.AppendLine("Homing Accuracy: " & HomingAccuracy.ToString)
            oSB.AppendLine("Explosion Radius: " & ExplosionRadius.ToString)
            oSB.AppendLine("Structural HP: " & StructuralHP.ToString)
        End If
        If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
            'full knowledge
            Dim sName As String = "Unknown Mineral"
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = BodyMaterialID Then
                    sName = goMinerals(X).MineralName
                    Exit For
                End If
            Next X
            oSB.AppendLine("Body Material: " & sName)
            sName = "Unknown Mineral"
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = NoseMaterialID Then
                    sName = goMinerals(X).MineralName
                    Exit For
                End If
            Next X
            oSB.AppendLine("Nose Material: " & sName)
            sName = "Unknown Mineral"
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = FlapsMaterialID Then
                    sName = goMinerals(X).MineralName
                    Exit For
                End If
            Next X
            oSB.AppendLine("Flaps Material: " & sName)
            sName = "Unknown Mineral"
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = FuelMaterialID Then
                    sName = goMinerals(X).MineralName
                    Exit For
                End If
            Next X
            oSB.AppendLine("Fuel Material: " & sName)
            sName = "Unknown Mineral"
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = PayloadMaterialID Then
                    sName = goMinerals(X).MineralName
                    Exit For
                End If
            Next X
            oSB.AppendLine("Payload Material: " & sName)
        End If
        Return oSB.ToString

	End Function
End Class
