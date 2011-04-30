''' <summary>
''' Manages all GNS functionality, should only be instantiated once, so use the shared method GetGNSMgr to acquire the current manager object
''' </summary>
''' <remarks></remarks>
Public Class GNSMgr
#Region "  Singleton Object  "
	Private Shared moMgr As GNSMgr
	Public Shared Function GetGNSMgr() As GNSMgr
		If moMgr Is Nothing Then moMgr = New GNSMgr()
		Return moMgr
	End Function
	'by making this private, the gnsmgr can only be declared thru our shared GetGNSMgr function, and since the GetGNSMgr function
	'  returns a shared object reference, this object becomes a singleton
	Private Sub New()
		'
	End Sub
#End Region

	Private moTemplates() As GNSTemplate
	Private myTemplateType() As Byte
	Private mlTemplateUB As Int32 = -1

	Private moCombats() As GNSCombat
	Private mlCombatUB As Int32 = -1

	Public Sub ParseGNSMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2	'for msgcode
		'MsgCode (2)
		'NewsType (1)
		'EnvirGUID (6)
		'NewsItemSpecifics...

		Dim yType As Byte = yData(lPos) : lPos += 1

		LogEvent("Received GNS News Item: " & CType(yType, NewsItemType).ToString)

		'ok, 
		Dim oData As New GNSData() 
		With oData
			'This GUID is not the GUID where the particulare event occurred... it is the GUID of the location where the server
			'  has determined will be required in order to view the GNS news story
			.LocationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.LocationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		End With
		'Now, based on our type, lets call the specific parse method
		Select Case CType(yType, NewsItemType)
			Case NewsItemType.eCeaseHostilities
				'TODO: Finish this one
			Case NewsItemType.eSpaceCombat, NewsItemType.eSpaceCombatEnd, NewsItemType.eSpaceCombatUpdate, NewsItemType.ePlanetCombat, NewsItemType.ePlanetCombatEnd, NewsItemType.ePlanetCombatUpdate
				ParseCombat(oData.LocationID, oData.LocationTypeID, yData, lPos, yType)
				Return
			Case NewsItemType.eLostColony
				ParseLostColony(yData, lPos, oData)
			Case NewsItemType.eLowMorale
				ParseLowMorale(yData, lPos, oData)
			Case NewsItemType.eMajorAlliances
				'TODO: Finish this
			Case NewsItemType.eMajorProdStart
				'TODO: Finish this
			Case NewsItemType.eNewSpaceStation
				ParseNewSpaceStation(yData, lPos, oData)
			Case NewsItemType.eOrbitalBombardment
				'TODO: Finish this
			Case NewsItemType.ePirateElimination
				ParsePirateElimNewsItem(yData, lPos, oData)
			Case NewsItemType.ePlanetControlShifts
				ParsePlanetControlShiftNewsItem(yData, lPos, oData)
			Case NewsItemType.ePlanetFall
				ParsePlayerStart(yData, lPos, oData)
			Case NewsItemType.ePlanetWideLowResources
				ParsePlanetLowRes(yData, lPos, oData)
			Case NewsItemType.ePlayerDeath
				ParsePlayerDeath(yData, lPos, oData)
				If oData.PlayerName Is Nothing OrElse oData.PlayerName = "" Then Return
			Case NewsItemType.eTerrorism
				'TODO: Finish this
			Case NewsItemType.eTitlePromotionHi, NewsItemType.eTitlePromotionLow, NewsItemType.eTitlePromotionMed
                ParsePlayerPromotion(yData, lPos, oData)
            Case NewsItemType.eTitleDemotion
                ParsePlayerDemotion(yData, lPos, oData)
            Case NewsItemType.eWarDeclaration
                'TODO: Finish this
			Case Else
				Return
		End Select

		If oData Is Nothing = False Then
			'ok, Now get our template
			Dim oTemplate As GNSTemplate = GetTemplate(CType(yType, NewsItemType))
			If oTemplate Is Nothing = False Then
				Dim sText As String = oTemplate.GetNewsStoryText(oData)
				If sText Is Nothing Then Return

				oData.TitleText = oTemplate.GetNewsTitleText(oData)
				oData.SectionID = oTemplate.lSectionID
				oData.BodyText = sText
				oData.sAssociated = GNSData.GetAssociated(oTemplate.sAssociated, oTemplate.lSectionID, oData.LocationTypeID)
				oData.SummaryText = oTemplate.GetSummaryText(oData)
				oData.sInformantText = oTemplate.GetInformant(oData.GetTagData("ENVIRNAME"))

				'here, we save the data to the database
				If oData.PostNewsItem() = False Then
					LogEvent("GNS Post News Item Failed")
				End If
			End If
		End If

	End Sub

	Public Function GetTemplate(ByVal lNewsType As NewsItemType) As GNSTemplate
		If mlTemplateUB = -1 Then LoadTemplates()

		Dim lIdx() As Int32 = Nothing
		Dim lUB As Int32 = -1

		For X As Int32 = 0 To mlTemplateUB
			If myTemplateType(X) = lNewsType Then
				lUB += 1
				ReDim Preserve lIdx(lUB)
				lIdx(lUB) = X
			End If
		Next X
		If lUB = -1 Then Return Nothing

		Dim lVal As Int32 = CInt(Rnd() * lUB)
		Return moTemplates(lIdx(lVal))
	End Function

	Private Sub LoadTemplates()
		Dim oINI As New InitFile()
		Dim sTemplatesDir As String = oINI.GetString("SETTINGS", "TemplatesDir", "Templates")
		Dim sDir As String = AppDomain.CurrentDomain.BaseDirectory
		If sDir.EndsWith("\") = False Then sDir &= "\"
		If sTemplatesDir.StartsWith("\") = True Then sTemplatesDir = sTemplatesDir.Substring(1)
		sDir &= sTemplatesDir
		If sDir.EndsWith("\") = False Then sDir &= "\"
		oINI = Nothing

		Me.mlTemplateUB = My.Computer.FileSystem.GetFiles(sDir).Count - 1
		ReDim moTemplates(mlTemplateUB)
		ReDim myTemplateType(mlTemplateUB)

		Dim lTemplateIdx As Int32 = -1

		'Now, get all of our templates...
		For Each sFile As String In My.Computer.FileSystem.GetFiles(sDir)
			If sFile.ToUpper.EndsWith(".TXT") = False Then
				mlTemplateUB -= 1
				Continue For
			End If

			lTemplateIdx += 1
			moTemplates(lTemplateIdx) = New GNSTemplate()

			Dim oFS As IO.FileStream = Nothing
			Dim oReader As IO.StreamReader = Nothing
			Try
				oFS = New IO.FileStream(sFile, IO.FileMode.Open)
				oReader = New IO.StreamReader(oFS)
				Dim lIdx As Int32 = 0
				Dim sBodyText As String = ""
				While oReader.EndOfStream = False
					If lIdx = 0 Then
						lIdx += 1
						moTemplates(lTemplateIdx).sTitle = oReader.ReadLine().Trim
					ElseIf lIdx = 1 Then
						lIdx += 1
						moTemplates(lTemplateIdx).sSummaryText = oReader.ReadLine.Trim
					ElseIf lIdx = 2 Then
						lIdx += 1
						Dim yVal As Byte = CByte(Val(oReader.ReadLine()))
						myTemplateType(lTemplateIdx) = yVal
						moTemplates(lTemplateIdx).yNewsType = CType(yVal, NewsItemType)
					ElseIf lIdx = 3 Then
						'informant
						moTemplates(lTemplateIdx).sInformant = oReader.ReadLine().Trim
						lIdx += 1
					ElseIf lIdx = 4 Then
						lIdx += 1
						Dim lVal As Int32 = CInt(Val(oReader.ReadLine()))
						moTemplates(lTemplateIdx).lSectionID = lVal
					ElseIf lIdx = 5 Then
						lIdx += 1
						moTemplates(lTemplateIdx).sAssociated = oReader.ReadLine().Trim
					Else
						sBodyText &= oReader.ReadLine()
					End If
				End While
				moTemplates(lTemplateIdx).sHTMLText = sBodyText
			Catch ex As Exception
				LogEvent("LoadTemplates: " & ex.Message)
			Finally
				If oFS Is Nothing = False Then oFS.Close()
				If oReader Is Nothing = False Then oReader.Close()
				oReader = Nothing
				oFS = Nothing
			End Try
		Next

	End Sub

	Public Sub ForceReloadTemplates()
		LoadTemplates()
	End Sub

#Region "  Message Parsers  "
	Private Sub ParseCombat(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByRef yData() As Byte, ByVal lPos As Int32, ByVal yType As Byte)
		'let's parse the data
		Dim uData As GNSCombatData = New GNSCombatData
		uData.SetFromGNSMsg(yData, lPos)

		Dim lIdx As Int32 = -1
		Dim lFirstIdx As Int32 = -1
		For X As Int32 = 0 To mlCombatUB
			If moCombats(X) Is Nothing = False Then
				If moCombats(X).lCombatLocationID = uData.lFightLocID AndAlso moCombats(X).iCombatLocationTypeID = uData.iFightLocTypeID AndAlso moCombats(X).lStoryEnvirID = lEnvirID AndAlso moCombats(X).iStoryEnvirTypeID = iEnvirTypeID Then
					lIdx = X
					Exit For
				End If
			ElseIf lFirstIdx = -1 Then
				lFirstIdx = X
			End If
		Next X
		If lIdx = -1 Then
			If lFirstIdx <> -1 Then
				lIdx = lFirstIdx
			Else
				SyncLock Me
					mlCombatUB += 1
					ReDim Preserve moCombats(mlCombatUB)
					lIdx = mlCombatUB
				End SyncLock
			End If
			moCombats(lIdx) = New GNSCombat
			moCombats(lIdx).lCombatLocationID = uData.lFightLocID
			moCombats(lIdx).iCombatLocationTypeID = uData.iFightLocTypeID
			moCombats(lIdx).lStoryEnvirID = lEnvirID
			moCombats(lIdx).iStoryEnvirTypeID = iEnvirTypeID
		End If
		moCombats(lIdx).AddGNSCombatData(uData, yType)
	End Sub

	Private Sub ParseLowMorale(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		With oData
			.SetPlayer1(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
			.ColonyName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.yLowMoraleReason = CType(yData(lPos), LowMoraleReason) : lPos += 1

			Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			Dim yGender As Byte = yData(lPos) : lPos += 1
			Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.SetPlayer1Specific(sName, yGender = 1, sEmpire)
		End With
	End Sub

	Private Sub ParseNewSpaceStation(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim sPlanet As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim sPlayerName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim yGender As Byte = yData(lPos) : lPos += 1
		Dim sEmpireName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

		'oData.SetPlayer1(lPlayerID)
		oData.SetPlayer1Specific(sPlayerName, yGender = 1, sEmpireName)

		oData.PlanetName = sPlanet
	End Sub

	Private Sub ParsePirateElimNewsItem(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		Dim sPlanet As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		oData.SetPlayer1(lPlayerID)
		oData.PlanetName = sPlanet

		Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim yGender As Byte = yData(lPos) : lPos += 1
		Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		If sName Is Nothing = False AndAlso sName <> "" Then oData.SetPlayer1Specific(sName, yGender = 1, sEmpire)
	End Sub

	Private Sub ParsePlanetControlShiftNewsItem(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		Dim sPlanet As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim sSystem As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lPrevOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		oData.SetPlayer1(lPlayerID)
		If lPrevOwnerID > 0 Then oData.SetPlayer2(lPrevOwnerID)
		oData.PlanetName = sPlanet
		oData.SystemName = sSystem

		Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim yGender As Byte = yData(lPos) : lPos += 1
		Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		If sName Is Nothing = False AndAlso sName <> "" Then oData.SetPlayer1Specific(sName, yGender = 1, sEmpire)

		sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		yGender = yData(lPos) : lPos += 1
		If sName Is Nothing = False AndAlso sName <> "" Then oData.SetPlayer2Specific(sName, yGender = 1)

	End Sub

	Private Sub ParsePlanetLowRes(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		Dim sPlanet As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim sPrimaryComp As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		oData.PlanetPrimaryComposition = sPrimaryComp
		oData.PlanetName = sPlanet
		If lOwnerID > 0 Then oData.SetPlayer1(lOwnerID)

		Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim yGender As Byte = yData(lPos) : lPos += 1
		Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		If sName Is Nothing = False AndAlso sName <> "" Then oData.SetPlayer1Specific(sName, yGender = 1, sEmpire)
	End Sub

	Private Sub ParsePlayerStart(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		Dim sPlayerName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim sPlanet As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		oData.PlayerName = sPlayerName
		oData.PlanetName = sPlanet
    End Sub

    Private Sub ParsePlayerDemotion(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
        Dim yTitle As Byte = yData(lPos) : lPos += 1
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        Dim yGender As Byte = yData(lPos) : lPos += 1
        Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

        oData.SetPlayer1(lPlayerID)
        oData.PlayerTitleScore = yTitle
        oData.SetPlayer1Specific(sName, yGender = 1, sEmpire)
    End Sub

	Private Sub ParsePlayerPromotion(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		Dim yTitle As Byte = yData(lPos) : lPos += 1
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lPrevPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim sPlanetName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

		oData.PlanetName = sPlanetName
		oData.SetPlayer1(lPlayerID)
		If lPrevPlayerID > 0 Then oData.SetPlayer2(lPrevPlayerID)
		oData.PlayerTitleScore = yTitle

		Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim yGender As Byte = yData(lPos) : lPos += 1
		Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		If sName Is Nothing = False AndAlso sName <> "" Then oData.SetPlayer1Specific(sName, yGender = 1, sEmpire)

		sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		yGender = yData(lPos) : lPos += 1
		If sName Is Nothing = False AndAlso sName <> "" Then oData.SetPlayer2Specific(sName, yGender = 1)
	End Sub

	Private Sub ParseLostColony(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		With oData
			.ColonyName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.PlanetName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.SetPlayer1(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
			.StartOfBattle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.MaxTotalPopulation = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.yColonyLostReason = CType(yData(lPos), ColonyLostReason) : lPos += 1

			Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			Dim yGender As Byte = yData(lPos) : lPos += 1
			Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.SetPlayer1Specific(sName, yGender = 1, sEmpire)
		End With
	End Sub

	Private Sub ParsePlayerDeath(ByRef yData() As Byte, ByVal lPos As Int32, ByRef oData As GNSData)
		With oData
			.SetPlayer1(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
			.MaxTotalPopulation = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.PlanetName = GetStringFromBytes(yData, lPos, 20) : lPos += 20

			Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			Dim yGender As Byte = yData(lPos) : lPos += 1
			Dim sEmpire As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			.SetPlayer1Specific(sName, yGender = 1, sEmpire)
		End With
	End Sub
#End Region
End Class
