Option Strict On

Public Class BombWeaponTech
	Inherits WeaponTech

	Public lPayloadSize As Int32
	Public yAOE As Byte
	Public yGuidance As Byte
	Public yRange As Byte
	Public yPayloadType As Byte
	Public iROF As Int16

	Public lPayloadMat As Int32
	Public lGuidanceMat As Int32
	Public lCasingMat As Int32
	
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
				lPayloadSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				yAOE = yData(lPos) : lPos += 1
				yGuidance = yData(lPos) : lPos += 1
				yRange = yData(lPos) : lPos += 1
				yPayloadType = yData(lPos) : lPos += 1
				iROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

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
				lPayloadMat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				lGuidanceMat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lCasingMat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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
				sResult &= "Casing's Density "
			Case eComponentDesignFlaw.eMat1_Prop2
				sResult &= "Casing's Compressibility "
			Case eComponentDesignFlaw.eMat1_Prop3
				sResult &= "Casing's Thermal Conductance "
			Case eComponentDesignFlaw.eMat1_Prop4
				sResult &= "Casing's Thermal Expansion "
			Case eComponentDesignFlaw.eMat1_Prop5
				sResult &= "Casing's Combustiveness "
			Case eComponentDesignFlaw.eMat2_Prop1
				sResult &= "Guidance's Electrical Resistance "
			Case eComponentDesignFlaw.eMat2_Prop2
				sResult &= "Guidance's Malleable "
			Case eComponentDesignFlaw.eMat2_Prop3
				sResult &= "Guidance's Magnetic Production "
			Case eComponentDesignFlaw.eMat2_Prop4
				sResult &= "Guidance's Superconductive Point "
			Case eComponentDesignFlaw.eMat3_Prop1
				sResult &= "Payload's Density "
			Case eComponentDesignFlaw.eMat3_Prop2
				sResult &= "Payload's Compressibility "
			Case eComponentDesignFlaw.eMat3_Prop3
				sResult &= "Payload's Payload Property "
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

	Public Overrides Function GetIntelReport(ByVal yLevel As PlayerTechKnowledge.KnowledgeType) As String
		Dim oSB As New System.Text.StringBuilder
		If yLevel > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
			'settings 1 or better
            oSB.AppendLine("Bomb for a " & Base_Tech.GetHullTypeName(HullTypeID))
			oSB.AppendLine("Power Required: " & PowerRequired.ToString("#,##0"))
			oSB.AppendLine("Hull Required: " & HullRequired.ToString("#,##0"))
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
			'settings 2 or better
			oSB.AppendLine("Payload Size: " & lPayloadSize.ToString("#,##0"))
			oSB.AppendLine("Rate of Fire: " & (iROF / 30.0F).ToString("#,##0.##"))
			oSB.AppendLine("Max Range: " & yRange.ToString("#,##0"))
			oSB.AppendLine("Area Effect: " & yAOE.ToString("#,##0"))
			oSB.AppendLine("Guidance: " & yGuidance.ToString("#,##0"))

			Select Case yPayloadType
				Case 0
					oSB.AppendLine("Payload Type: Explosive")
				Case 1
					oSB.AppendLine("Payload Type: Toxic")
				Case 2
					oSB.AppendLine("Payload Type: Concussive")
				Case 3
					oSB.AppendLine("Payload Type: EMP")
			End Select 
		End If
		If yLevel > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
			'full knowledge
			Dim sName As String = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lPayloadMat Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Payload Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lGuidanceMat Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Guidance Material: " & sName)
			sName = "Unknown Mineral"
			For X As Int32 = 0 To glMineralUB
				If glMineralIdx(X) = lCasingMat Then
					sName = goMinerals(X).MineralName
					Exit For
				End If
			Next X
			oSB.AppendLine("Casing Material: " & sName)
			 
		End If
		Return oSB.ToString
	End Function

    Public Overrides Sub GetPrototypeBuilderAttributeString(ByRef oSB As System.Text.StringBuilder, ByVal bPointDefense As Boolean)
        oSB.AppendLine("    " & WeaponName)
        oSB.AppendLine("    RoF" & (iROF / 30.0F).ToString(" #,##0.#0").PadLeft(23, "."c))
        oSB.AppendLine("    Max Damage" & lPayloadSize.ToString(" #,###").PadLeft(16, "."c))
    End Sub

    Public Overrides Function GetPrototypeBuilderString(ByVal bPointDefense As Boolean) As String
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder()
        Dim sFinal As String = ""
        With Me
            oSB.AppendLine("ROF: " & (.iROF / 30.0F).ToString("#,###.#0") & " s")
            Dim lRange As Int32 = .yRange
            If bPointDefense = True Then lRange = CInt(lRange * 0.25F)
            oSB.AppendLine("Max Range: " & lRange)
            oSB.AppendLine("Explosion Radius: " & .yAOE)

            Dim lHull As Int32 = .HullRequired
            If bPointDefense = True Then lHull = CInt(lHull * 0.5F)

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

            If bPointDefense = True Then
                oSB.AppendLine()
                oSB.AppendLine("As A Point Defense Weapon:")
                oSB.AppendLine("Damage reduced by 50%")
                oSB.AppendLine("Accuracy Increased by 25")
            End If

            sFinal = oSB.ToString
        End With

        oSB = Nothing
        Return sFinal
    End Function

	Protected Overrides Sub PopulateFromMsg(ByRef yData() As Byte, ByVal lPos As Integer)
		With Me
			.lPayloadSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.yAOE = yData(lPos) : lPos += 1
			.yGuidance = yData(lPos) : lPos += 1
			.yRange = yData(lPos) : lPos += 1
			.yPayloadType = yData(lPos) : lPos += 1
			.iROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.lPayloadMat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lGuidanceMat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.lCasingMat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			FillWeaponDefFromMsg(yData, lPos)
		End With
	End Sub
End Class
