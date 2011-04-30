Public Enum NewsItemType As Byte
	eSpaceCombat = 0
	eSpaceCombatUpdate = 1
	eLowMorale = 2
	eTerrorism = 3
	eTitlePromotionLow = 4
	ePlanetFall = 5
	eMajorProdStart = 6
	ePlanetControlShifts = 7
	ePlanetWideLowResources = 8
	eLostColony = 9
	ePlayerDeath = 10
	ePirateElimination = 11
	eNewSpaceStation = 12
	eWarDeclaration = 13
	eCeaseHostilities = 14
	eMajorAlliances = 15
	eOrbitalBombardment = 16
	eSpaceCombatEnd = 17
	ePlanetCombat = 18
	ePlanetCombatUpdate = 19
	ePlanetCombatEnd = 20
	eOfflineWarDec = 21
	eTitlePromotionMed = 22
    eTitlePromotionHi = 23
    eTitleDemotion = 24
End Enum

Public Class GNSTemplate
	Public yNewsType As NewsItemType
	Public lSectionID As Int32		'section this story should appear
	Public sTitle As String
	Public sHTMLText As String		'the template

	Public sSummaryText As String

	Public sAssociated As String

	Public sInformant As String

	Private msParsedLines() As String
	Private msParsedTitle() As String
	Private msParsedSummary() As String

	Public Function GetNewsTitleText(ByRef oData As GNSData) As String
		If msParsedTitle Is Nothing Then ParseTitleLines()
		If msParsedTitle Is Nothing Then Return ""

		Dim oSB As New System.Text.StringBuilder()
		Dim lExclude As Int32 = 0
		oSB.Length = 0
		For X As Int32 = 0 To msParsedTitle.GetUpperBound(0)
			If msParsedTitle(X).StartsWith("[") = True Then
				'ok, a command... 
				Try
					Dim sResult As String = oData.GetTagData(msParsedTitle(X))
					If sResult Is Nothing Then sResult = ""
					If sResult.ToUpper = "EXCLUDE" Then
						lExclude += 1
					ElseIf sResult.ToUpper = "UNEXCLUDE" Then
						lExclude -= 1
					ElseIf lExclude < 1 AndAlso sResult <> "" Then
						oSB.Append(sResult)
					End If
				Catch
					'do nothing for now
				End Try
			ElseIf lExclude < 1 Then
				'just normal text
				oSB.Append(msParsedTitle(X))
			End If
		Next X
		Return oSB.ToString
	End Function

	Public Function GetNewsStoryText(ByRef oData As GNSData) As String
		'ok, let's parse the HTMLText
		If msParsedLines Is Nothing Then ParseLines()
		If msParsedLines Is Nothing Then Return ""

		Dim oSB As New System.Text.StringBuilder()
		Dim lExclude As Int32 = 0
		oSB.Length = 0
		For X As Int32 = 0 To msParsedLines.GetUpperBound(0)
			If msParsedLines(X).StartsWith("[") = True Then
				'ok, a command... 
				Try
					Dim sResult As String = oData.GetTagData(msParsedLines(X))
					If sResult Is Nothing Then sResult = ""
					If sResult.ToUpper = "EXCLUDE" Then
						lExclude += 1
					ElseIf sResult.ToUpper = "UNEXCLUDE" Then
						lExclude -= 1
					ElseIf lExclude < 1 AndAlso sResult <> "" Then
						oSB.Append(sResult)
					End If
				Catch
					'do nothing for now
				End Try
			ElseIf lExclude < 1 Then
				'just normal text
				oSB.Append(msParsedLines(X))
			End If
		Next X
		Return oSB.ToString
	End Function

	Public Function GetSummaryText(ByRef oData As GNSData) As String
		If msParsedSummary Is Nothing Then ParseSummaryLines()
		If msParsedSummary Is Nothing Then Return ""

		Dim oSB As New System.Text.StringBuilder()
		Dim lExclude As Int32 = 0
		oSB.Length = 0
		For X As Int32 = 0 To msParsedSummary.GetUpperBound(0)
			If msParsedSummary(X).StartsWith("[") = True Then
				'ok, a command... 
				Try
					Dim sResult As String = oData.GetTagData(msParsedSummary(X))
					If sResult Is Nothing Then sResult = ""
					If sResult.ToUpper = "EXCLUDE" Then
						lExclude += 1
					ElseIf sResult.ToUpper = "UNEXCLUDE" Then
						lExclude -= 1
					ElseIf lExclude < 1 AndAlso sResult <> "" Then
						oSB.Append(sResult)
					End If
				Catch
					'do nothing for now
				End Try
			ElseIf lExclude < 1 Then
				'just normal text
				oSB.Append(msParsedSummary(X))
			End If
		Next X
		Return oSB.ToString
	End Function

	Private Sub ParseTitleLines()
		'Ok, split up our text to strings with [] items their own strings
		Dim lUB As Int32 = 0

		ReDim msParsedTitle(0)
		msParsedTitle(0) = ""

		For X As Int32 = 0 To sTitle.Length - 1
			Dim sCur As String = sTitle(X)
			If sCur = "[" Then
				'end our current value
				lUB += 1
				ReDim Preserve msParsedTitle(lUB)
				msParsedTitle(lUB) = ""
			End If
			msParsedTitle(lUB) &= sCur
			If sCur = "]" Then
				'end our current value
				lUB += 1
				ReDim Preserve msParsedTitle(lUB)
				msParsedTitle(lUB) = ""
			End If
		Next X
	End Sub

	Private Sub ParseLines()
		'Ok, split up our text to strings with [] items their own strings
		Dim lUB As Int32 = 0

		ReDim msParsedLines(0)
		msParsedLines(0) = ""

		For X As Int32 = 0 To sHTMLText.Length - 1
			Dim sCur As String = sHTMLText(X)
			If sCur = "[" Then
				'end our current value
				lUB += 1
				ReDim Preserve msParsedLines(lUB)
				msParsedLines(lUB) = ""
			End If
			msParsedLines(lUB) &= sCur
			If sCur = "]" Then
				'end our current value
				lUB += 1
				ReDim Preserve msParsedLines(lUB)
				msParsedLines(lUB) = ""
			End If
		Next X
	End Sub

	Private Sub ParseSummaryLines()
		'Ok, split up our text to strings with [] items their own strings
		Dim lUB As Int32 = 0

		ReDim msParsedSummary(0)
		msParsedSummary(0) = ""

		For X As Int32 = 0 To sSummaryText.Length - 1
			Dim sCur As String = sSummaryText(X)
			If sCur = "[" Then
				'end our current value
				lUB += 1
				ReDim Preserve msParsedSummary(lUB)
				msParsedSummary(lUB) = ""
			End If
			msParsedSummary(lUB) &= sCur
			If sCur = "]" Then
				'end our current value
				lUB += 1
				ReDim Preserve msParsedSummary(lUB)
				msParsedSummary(lUB) = ""
			End If
		Next X
	End Sub

	Private Function GetAuthorName() As String
		Dim sAuthorName As String = sInformant
		If sAuthorName.Trim.ToUpper = "RANDOM" Then
			Dim bMale As Boolean = CInt(Rnd() * 100) < 50
			Dim sFirstName As String = ""
			Dim sLastName As String = ""

			GenerateName(bMale, sFirstName, sLastName)
			sAuthorName = sFirstName & " " & sLastName
			Return sAuthorName
		Else : Return sAuthorName.Trim
		End If
	End Function

	Public Function GetInformant(ByVal sEnvirName As String) As String
		Dim sResult As String = sEnvirName
		If sResult Is Nothing Then sResult = ""
		If sResult = "" Then Return GetAuthorName() & ", GNS News"
		sResult = sResult.Trim & " Correspondent"

		sResult = GetAuthorName() & ", " & sResult
		Return sResult
	End Function
End Class