Public Class AlloyTech
	Inherits Base_Tech

	Public sAlloyName As String
	Public AlloyResultID As Int32 = -1
	Public Mineral1ID As Int32 = -1
	Public Mineral2ID As Int32 = -1
	Public Mineral3ID As Int32 = -1
	Public Mineral4ID As Int32 = -1
	Public PropertyID1 As Int32 = -1
	Public PropertyID2 As Int32 = -1
	Public PropertyID3 As Int32 = -1
	'Public bHigher1 As Boolean
	'Public bHigher2 As Boolean
	'Public bHigher3 As Boolean
	Public yValue1 As Byte = 255
	Public yValue2 As Byte = 255
	Public yValue3 As Byte = 255
	Public ResearchLevel As Byte

	Public oExpectedResult As Mineral = Nothing

	Public Overrides Sub FillFromMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		With Me
			.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			.OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If .OwnerID = 0 Then .OwnerID = glPlayerID

			.ComponentDevelopmentPhase = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.ErrorCodeReason = yData(lPos) : lPos += 1
			.Researchers = yData(lPos) : lPos += 1
			.MajorDesignFlaw = yData(lPos) : lPos += 1

			'Now, the alloy tech specifics...
			.sAlloyName = GetStringFromBytes(yData, lPos, 20)
			lPos += 20

			'MSC - alloy result is now at the end...
			'.AlloyResultID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

			.Mineral1ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.Mineral2ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.Mineral3ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.Mineral4ID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PropertyID1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PropertyID2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PropertyID3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			'.bHigher1 = (yData(lPos) <> 0) : lPos += 1
			'.bHigher2 = (yData(lPos) <> 0) : lPos += 1
			'.bHigher3 = (yData(lPos) <> 0) : lPos += 1
			.yValue1 = yData(lPos) : lPos += 1
			.yValue2 = yData(lPos) : lPos += 1
			.yValue3 = yData(lPos) : lPos += 1
			.ResearchLevel = yData(lPos) : lPos += 1

			.AlloyResultID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		End With

		If AlloyResultID = -1 AndAlso lPos < yData.Length Then
			'Ok, special... the rest of this message is essentially the player mineral object
			'  well, in that it has a CNT
			Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			oExpectedResult = New Mineral()
			With oExpectedResult
				.MineralName = Me.sAlloyName
				For X As Int32 = 0 To lCnt - 1
					'4 bytes for propid
					Dim lPropID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					'1 byte for percentage
					Dim yVal As Byte = yData(lPos) : lPos += 1
					.SetMineralPropertyValue(lPropID, "", "", yVal)
				Next X
			End With
		End If
	End Sub

	Public Overrides Function GetDesignFlawText() As String
		Return ""
	End Function

	Public Overrides Sub FillFromPlayerTechKnowledge(ByRef yData() As Byte, ByVal lOffset As Integer, ByVal yTechKnow As Byte, ByVal sName As String)
		'TODO: Should never be called???
	End Sub

	Public Overrides Function GetIntelReport(ByVal yLevel As PlayerTechKnowledge.KnowledgeType) As String
		Return "Not yet implemented, bug me"
	End Function
End Class
