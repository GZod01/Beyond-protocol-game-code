Public Enum LowMoraleReason As Byte
	UnemploymentRate = 1
	TaxRate = 2
	Homeless = 4
	UnpoweredHomes = 8
	Sentiment = 16
End Enum
Public Enum ColonyLostReason As Byte
	Neglect = 1
	Destruction = 2
	LowMorale = 3
	Abandoned = 4
End Enum
Public Enum PlayerRank As Byte
	Magistrate = 0
	Governor = 1
	Overseer = 2
	Duke = 3
	Baron = 4
	King = 5
	Emperor = 6
End Enum

Public Class GNSData

	'News Item Specifics
	Public LocationID As Int32
	Public LocationTypeID As Int16
	Public SectionID As Int32
	Public sAssociated As String
	Public TitleText As String
	Public BodyText As String
	Public SummaryText As String
	'End of news item specifics

	'contains the data dictionary properties...
	Public PlanetName As String
	Public SystemName As String

	Public PlayerName As String
	Public PlayerIsMale As Boolean = False
	Public PlayerName2 As String
	Public Player2IsMale As Boolean = False

	Public PlayerSideListA As String
	Public PlayerSideListB As String
	Public StartOfBattle As Int32			'date as number
	Public EndOfBattle As Int32 = 0			'date as number

	''' <summary>
	''' Returns an int32 of number of seconds elapsed
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public ReadOnly Property SpanOfBattle() As Int32
		Get
			Dim dtStart As Date = GetDateFromNumber(StartOfBattle)
			Dim dtEnd As Date = Now
			If EndOfBattle > 0 Then dtEnd = GetDateFromNumber(EndOfBattle)
			Return CInt(dtEnd.Subtract(dtStart).TotalSeconds)
		End Get
	End Property

	Public TotalUnitLostSideA As Int64
	Public TotalFacilitiesLostSideA As Int64
	Public TotalResidenceLostSideA As Int64
	Public TotalJobsLostSideA As Int64
	Public TotalLostSideA As Int64		'enlisted and officers

	Public TotalUnitLostSideB As Int64
	Public TotalFacilitiesLostSideB As Int64
	Public TotalResidenceLostSideB As Int64
	Public TotalJobsLostSideB As Int64
	Public TotalLostSideB As Int64		'enlisted and officers

	Public ColonyName As String
	Public yLowMoraleReason As LowMoraleReason
	Public yColonyLostReason As ColonyLostReason

	Public PlanetPrimaryComposition As String		'minerals that compose the most mineable caches of the planet
	Public MaxTotalPopulation As Int64

	Public WarDecList As String			'comma-delimited player list of players at war with parties involved in this news item
	Public CorporationName As String

	Public PlayerTitleScore As Byte
	Public EmpireName As String
	Public TotalInvolved As Int32

	Public sInformantText As String


	Public ReadOnly Property TotalUnitLostBothSides() As Int64
		Get
			Return TotalUnitLostSideA + TotalUnitLostSideB
		End Get
	End Property
	Public ReadOnly Property TotalFacilitiesLostBothSides() As Int64
		Get
			Return TotalFacilitiesLostSideA + TotalFacilitiesLostSideB
		End Get
	End Property
	Public ReadOnly Property TotalResidenceLostBothSides() As Int64
		Get
			Return TotalResidenceLostSideA + TotalResidenceLostSideB
		End Get
	End Property
	Public ReadOnly Property TotalJobsLostBothSides() As Int64
		Get
			Return TotalJobsLostSideA + TotalJobsLostSideB
		End Get
	End Property
	Public ReadOnly Property TotalLostBothSides() As Int64
		Get
			Return TotalLostSideA + TotalLostSideB
		End Get
	End Property
	Public ReadOnly Property TotalEntitiesLostSideA() As Int64
		Get
			Return TotalUnitLostSideA + TotalFacilitiesLostSideA
		End Get
	End Property
	Public ReadOnly Property TotalEntitiesLostSideB() As Int64
		Get
			Return TotalUnitLostSideB + TotalFacilitiesLostSideB
		End Get
	End Property
	Public ReadOnly Property TotalEntitiesLostBothSides() As Int64
		Get
			Return TotalEntitiesLostSideA + TotalEntitiesLostSideB
		End Get
	End Property
	Public ReadOnly Property LowMoraleReasonText() As String
		Get
			Dim sResult As String = ""
			If (yLowMoraleReason And LowMoraleReason.Homeless) <> 0 Then
				If sResult <> "" Then sResult &= " and "
				sResult &= "Homeless rate"
			End If
			If (yLowMoraleReason And LowMoraleReason.Sentiment) <> 0 Then
				If sResult <> "" Then sResult &= " and "
				sResult &= "Lack of Government Commitment"
			End If
			If (yLowMoraleReason And LowMoraleReason.TaxRate) <> 0 Then
				If sResult <> "" Then sResult &= " and "
				sResult &= "Tax Rate"
			End If
			If (yLowMoraleReason And LowMoraleReason.UnemploymentRate) <> 0 Then
				If sResult <> "" Then sResult &= " and "
				sResult &= "Unemployment Rate"
			End If
			If (yLowMoraleReason And LowMoraleReason.UnpoweredHomes) <> 0 Then
				If sResult <> "" Then sResult &= " and "
				sResult &= "Unpowered Residence"
			End If

			Return sResult
		End Get
	End Property
	Public ReadOnly Property PlayerTitle() As String
		Get
			Return GetPlayerTitle(PlayerTitleScore, PlayerIsMale)
		End Get
	End Property
	Public ReadOnly Property ColonyLostReasonText() As String
		Get
			Select Case yColonyLostReason
				Case ColonyLostReason.Abandoned
					Return "abandonment"
				Case ColonyLostReason.Destruction
					Return "destruction"
				Case ColonyLostReason.LowMorale
					Return "low morale"
				Case ColonyLostReason.Neglect
					Return "neglect"
				Case Else
					Return ""
			End Select
		End Get
	End Property
	Public Function GetTagData(ByVal sTag As String) As String
		sTag = sTag.Replace("[", "").Replace("]", "").ToUpper

		Dim bPossessive As Boolean = False
		If sTag.EndsWith("POSSESSIVE") = True Then
			bPossessive = True
			sTag = sTag.Substring(0, sTag.Length - 10)
		End If

		If sTag.StartsWith("EXCLUDE:") = True Then
			'format of exclude is:
			'[EXCLUDE:<attrname><operator><value>]
			'we've removed the [ and the ] already
			'pull of the EXCLUDE:
			sTag = sTag.Substring(8)
			If CheckExclude(sTag) = True Then Return "EXCLUDE" Else Return ""
		ElseIf sTag.StartsWith("/EXCLUDE") = True Then
			Return "UNEXCLUDE"
		End If
		sTag = sTag.Trim

		Select Case sTag
			Case "PLANETNAME", "NEARESTPLANETNAME"
				If bPossessive = True Then
					If PlanetName.EndsWith("s") = True Then
						Return PlanetName & "'"
					Else : Return PlanetName & "'s"
					End If
				Else : Return PlanetName
				End If
			Case "SYSTEMNAME"
				If bPossessive = True Then
					If SystemName.EndsWith("s") = True Then
						Return SystemName & "'"
					Else : Return SystemName & "'s"
					End If
				Else : Return SystemName
				End If
			Case "PLAYERNAME", "PLAYERNAME2", "PREVIOUSPLAYERNAME", "OWNEROFTITLEPREVIOUSLY"
				Dim sTemp As String
				If sTag = "PLAYERNAME" Then sTemp = PlayerName Else sTemp = PlayerName2
				If bPossessive = True Then
					If sTemp.EndsWith("s") = True Then
						Return sTemp & "'"
					Else : Return sTemp & "'s"
					End If
				Else : Return sTemp
				End If
			Case "PLAYERSIDELISTA", "PLAYERLISTSIDEA"
				If PlayerSideListA Is Nothing OrElse PlayerSideListA = "" Then
					If PlayerName Is Nothing = False AndAlso PlayerName <> "" Then Return PlayerName
				End If
				Return PlayerSideListA
			Case "PLAYERSIDELISTB", "PLAYERLISTSIDEB"
				If PlayerSideListB Is Nothing OrElse PlayerSideListB = "" Then
					If PlayerName2 Is Nothing = False AndAlso PlayerName2 <> "" Then Return PlayerName2
				End If
				Return PlayerSideListB
			Case "STARTOFBATTLE"
				Return GetDateFromNumberAsString(StartOfBattle)
			Case "ENDOFBATTLE"
				Return GetDateFromNumberAsString(EndOfBattle)
			Case "SPANOFBATTLE"
				Return GetDurationFromSeconds(SpanOfBattle(), True)
			Case "TOTALUNITLOSTBOTHSIDES", "TOTALUNITSLOSTBOTHSIDES"
				Return TotalUnitLostBothSides.ToString("#,##0")
			Case "TOTALUNITLOSTSIDEA", "TOTALUNITSLOSTSIDEA"
				Return TotalUnitLostSideA.ToString("#,##0")
			Case "TOTALUNITLOSTSIDEB", "TOTALUNITSLOSTSIDEB"
				Return TotalUnitLostSideB.ToString("#,##0")
			Case "TOTALLOSTRESIDENCEBOTHSIDES"
				Return TotalResidenceLostBothSides.ToString("#,##0")
			Case "TOTALLOSTRESIDENCESIDEA"
				Return TotalResidenceLostSideA.ToString("#,##0")
			Case "TOTALLOSTRESIDENCESIDEB"
				Return TotalResidenceLostSideB.ToString("#,##0")
			Case "TOTALLOSTJOBSBOTHSIDES"
				Return TotalJobsLostBothSides.ToString("#,##0")
			Case "TOTALLOSTJOBSSIDEA"
				Return TotalJobsLostSideA.ToString("#,##0")
			Case "TOTALLOSTJOBSSIDEB"
				Return TotalJobsLostSideB.ToString("#,##0")
			Case "TOTALFACILITIESLOSTBOTHSIDES"
				Return TotalFacilitiesLostBothSides.ToString("#,##0")
			Case "TOTALFACILITIESLOSTSIDEA"
				Return TotalFacilitiesLostSideA.ToString("#,##0")
			Case "TOTALFACILITIESLOSTSIDEB"
				Return TotalFacilitiesLostSideB.ToString("#,##0")
			Case "TOTALLOSTBOTHSIDES"
				Return TotalLostBothSides.ToString("#,##0")
			Case "TOTALLOSTSIDEA"
				Return TotalLostSideA.ToString("#,##0")
			Case "TOTALLOSTSIDEB"
				Return TotalLostSideB.ToString("#,##0")
			Case "TOTALENTITIESLOSTBOTHSIDES"
				Return TotalEntitiesLostBothSides.ToString("#,##0")
			Case "TOTALENTITIESLOSTSIDEA"
				Return TotalEntitiesLostSideA.ToString("#,##0")
			Case "TOTALENTITIESLOSTSIDEB"
				Return TotalEntitiesLostSideB.ToString("#,##0")
			Case "COLONYNAME"
				If bPossessive = True Then
					If ColonyName.EndsWith("s") = True Then
						Return ColonyName & "'"
					Else : Return ColonyName & "'s"
					End If
				Else : Return ColonyName
				End If
			Case "LOWMORALEREASON"
				Return LowMoraleReasonText()
			Case "WARDECLIST"
				Return WarDecList
			Case "CORPORATIONNAME"
				If bPossessive = True Then
					If CorporationName.EndsWith("s") = True Then
						Return CorporationName & "'"
					Else : Return CorporationName & "'s"
					End If
				Else : Return CorporationName
				End If
			Case "COLONYENVIRLIST"
				Return ColonyName
			Case "MAXTOTALPOPULATION"
				Return MaxTotalPopulation.ToString("#,##0")
			Case "PLAYERTITLE"
				Return PlayerTitle
			Case "PLAYERTITLESCORE"
				Return PlayerTitleScore.ToString
			Case "EMPIRENAME"
				If bPossessive = True Then
					If EmpireName.EndsWith("s") = True Then
						Return EmpireName & "'"
					Else : Return EmpireName & "'s"
					End If
				Else : Return EmpireName
				End If
			Case "TOTALINVOLVED"
				Return TotalInvolved.ToString
			Case "ENVIRNAME"
				If PlanetName <> "" Then
					Return GetTagData("PLANETNAME")
				ElseIf SystemName <> "" Then
					Return GetTagData("SYSTEMNAME")
				End If
			Case "COLONYLOSTREASON"
				Return ColonyLostReasonText()
			Case "COLONYSTART"
				Return GetDateFromNumberAsString(StartOfBattle)
			Case "OWNEROFTITLEPREVIOUSLYTITLE"
				Return GetPlayerTitle(PlayerTitleScore, Player2IsMale)
		End Select

		Return ""
	End Function
	Private Function CheckExclude(ByVal sExpression As String) As Boolean
		'should already be formatted
		'expression operator value		'the spaces may not exist...

		'so, here's our valid operator list...
		'=, <>, >, <, >=, <=

		Dim cChecks(2) As Char
		cChecks(0) = "="c
		cChecks(1) = ">"c
		cChecks(2) = "<"c

		Dim lIdx As Int32 = sExpression.IndexOfAny(cChecks)
		If lIdx <> -1 Then
			Dim sVariable As String = ""
			Dim sValue As String = ""
			Dim sNext As String = sExpression(lIdx + 1)
			Dim sOperator As String = sExpression(lIdx)

			If sNext = ">" Then
				If sOperator = "=" Then sOperator = ">=" Else sOperator &= sNext
			ElseIf sNext = "=" Then
				If sOperator <> "=" Then sOperator &= sNext
			ElseIf sNext = "<" Then
				If sOperator = ">" Then
					sOperator = "<>"
				ElseIf sOperator = "=" Then
					sOperator = "<="
				End If
			End If

			'ok, get our variable and value
			sVariable = sExpression.Substring(0, lIdx).Trim
			If sOperator.Length = 1 Then
				sValue = sExpression.Substring(lIdx + 1).Trim
			Else
				sValue = sExpression.Substring(lIdx + 2).Trim
			End If

			Dim sDataValue As String = GetTagData(sVariable)
			If sValue = """""" OrElse sValue = """" Then sValue = ""
			If sDataValue Is Nothing Then sDataValue = ""

			'now, do our comparison
			If IsNumeric(sValue) AndAlso (IsNumeric(sDataValue) OrElse sDataValue = "") Then
				Dim lValue As Int32 = CInt(Val(sValue))
				Dim lDataValue As Int32 = 0
				If sDataValue <> "" Then lDataValue = CInt(Val(sDataValue))

				Select Case sOperator
					Case "<"
						Return lDataValue < lValue
					Case ">"
						Return lDataValue > lValue
					Case "<="
						Return lDataValue <= lValue
					Case ">="
						Return lDataValue >= lValue
					Case "<>"
						Return lDataValue <> lValue
					Case Else	'equals
						Return lDataValue = lValue
				End Select
			Else
				Select Case sOperator
					Case "<"
						Return sDataValue < sValue
					Case ">"
						Return sDataValue > sValue
					Case "<="
						Return sDataValue <= sValue
					Case ">="
						Return sDataValue >= sValue
					Case "<>"
						Return sDataValue <> sValue
					Case Else	'equals
						Return sDataValue = sValue
				End Select
			End If

		End If
		Return False
	End Function

	Public Sub SetPlayer1(ByVal lPlayerID As Int32)
		For X As Int32 = 0 To glPlayerUB
			If glPlayerIdx(X) = lPlayerID Then
				PlayerName = goPlayer(X).sPlayerName
				PlayerIsMale = goPlayer(X).yGender = 1
				EmpireName = goPlayer(X).sEmpireName
				Exit For
			End If
		Next X
	End Sub
	Public Sub SetPlayer2(ByVal lPlayerID As Int32)
		For X As Int32 = 0 To glPlayerUB
			If glPlayerIdx(X) = lPlayerID Then
				PlayerName2 = goPlayer(X).sPlayerName
				Player2IsMale = goPlayer(X).yGender = 1
				Exit For
			End If
		Next X
	End Sub

	Public Sub SetPlayer1Specific(ByVal sName As String, ByVal bMale As Boolean, ByVal sEmpire As String)
		PlayerName = sName
		PlayerIsMale = bMale
		EmpireName = sEmpire
	End Sub
	Public Sub SetPlayer2Specific(ByVal sName As String, ByVal bMale As Boolean)
		PlayerName2 = sName
		Player2IsMale = bMale
	End Sub

	Public Function PostNewsItem() As Boolean
		Dim bResult As Boolean = False

		'Dim lSortOrder As Int32 = 0		'todo: figure this out
		'INSERT INTO `nuke_stories` 
		'(`sid`, `catid`, `aid`, `title`, `time`,`hometext`, `bodytext`, `comments`, `counter`, `topic`, `informant`,`notes`, 
		'`ihome`, `alanguage`, `acomm`, `haspoll`, `pollID`, `score`,`ratings`, `rating_ip`, `associated`) 
		'VALUES 
		'(2, 0, 'GNS', 'Galactic News Service Returns Soon', '2007-12-2010:08:44', '<newssummary goes here>', '<news body goes here>',
		'0, 6, 2, 'BPGod', '', 0, '', 0, 0, 0, 3, 1, '80.136.60.228', '2-');


		'Dim sSummary As String = Mid$(Me.BodyText, 1, 75)
		'Dim lTempIdx As Int32 = sSummary.LastIndexOf(" "c)
		'If sSummary.LastIndexOf("."c) > lTempIdx Then lTempIdx = sSummary.LastIndexOf("."c)
		'If lTempIdx > -1 Then
		'	sSummary = sSummary.Substring(0, lTempIdx)
		'End If
		'sSummary &= "..."

		''Now, we need to strip out the html
		'Dim lHTMLLvl As Int32 = 0
		'Dim sTemp As String = sSummary
		'sSummary = ""
		'For X As Int32 = 0 To sTemp.Length - 1
		'	If sTemp(X) = "<"c Then
		'		lHTMLLvl += 1
		'	ElseIf sTemp(X) = ">"c Then
		'		lHTMLLvl -= 1
		'	ElseIf lHTMLLvl = 0 Then
		'		sSummary &= sTemp(X)
		'	End If
		'Next X

        'Now, put our stuff in for the summary

        If gb_IS_TEST_SERVER = True Then Return True

		SummaryText = "<div style=""font-size: 12pt; color:#FFF;"">" & SummaryText & "</div>"

        Dim sSQL As String = "INSERT INTO `nuke_stories` (`catid`, `aid`, `title`, `time`,`hometext`, `bodytext`, `comments`, " & _
          "`counter`, `topic`, `informant`,`associated`) VALUES (0, 'MMORTS', '" & MakeDBStr(TitleText) & "', '" & _
          Date.UtcNow.ToString("yyyy-MM-dd HH:mm:ss")
		sSQL &= "', '" & MakeDBStr(SummaryText) & "', '" & MakeDBStr(BodyText) & "', 0, 0, " & SectionID & ", '" & MakeDBStr(sInformantText) & "', '" & sAssociated & "')"

		Dim oComm As Odbc.OdbcCommand = Nothing
		Try
			oComm = New Odbc.OdbcCommand(sSQL, goGNS_CN)
			Dim lRecs As Int32 = oComm.ExecuteNonQuery()
			If lRecs < 1 Then Err.Raise(-1, "PostNewsItem", "No records affected with insert")
			bResult = True
		Catch ex As Exception
			LogEvent("PostNewsItem: " & ex.Message)
			bResult = False
		Finally
			If oComm Is Nothing = False Then oComm.Dispose()
			oComm = Nothing
		End Try
		Return bResult

		'Dim sSQL As String = "INSERT INTO tblNewsArticles (SectionID, Title, TextHTML, SortOrder, CreatedBy, DateCreated, ModifiedBy, "
		'sSQL &= "DateModified, MinAuthView, MinAuthEdit, LocationID, LocationTypeID) VALUES (" & SectionID & ", '" & MakeDBStr(TitleText) & _
		' "', '" & MakeDBStr(BodyText) & "', " & lSortOrder & ", " & l_GNS_USER_NAME & ", '" & Now.ToString("MM/dd/yyyy HH:mm") & "', " & _
		' l_GNS_USER_NAME & ", '" & Now.ToString("MM/dd/yyyy HH:mm") & "', 0, 50, " & LocationID & ", " & LocationTypeID & ")"
		'Dim oComm As OleDb.OleDbCommand = Nothing
		'Try
		'	oComm = New OleDb.OleDbCommand(sSQL, goGNS_CN)
		'	Dim lRecs As Int32 = oComm.ExecuteNonQuery()
		'	If lRecs < 1 Then Err.Raise(-1, "PostNewsItem", "No records affected with insert")
		'	bResult = True
		'Catch ex As Exception
		'	LogEvent("PostNewsItem: " & ex.Message)
		'	bResult = False
		'Finally
		'	If oComm Is Nothing = False Then oComm.Dispose()
		'	oComm = Nothing
		'End Try
		'Return bResult
	End Function

	'Private Function GetTopicNumber() As Int32
	'	'--GNS Global is 2
	'	'--GNS Space battle is 3
	'	'--GNS Planet battle is 4
	'	'--GNS Politics is 5
	'	'--GNS Technology is 6
	'	'--GNS Breaking News is 7
	'	'--GNS Travel is 8 
	'End Function
	Public Shared Function GetAssociated(ByVal psBaseAssoc As String, ByVal plSectionID As Int32, ByVal piLocTypeID As Int16) As String
		Dim sResult As String = psBaseAssoc
		If sResult Is Nothing Then sResult = ""
		If sResult.Trim = "" Then sResult = plSectionID.ToString
		If sResult.EndsWith("-") = False Then sResult &= "-"

		'relates multiple topic numbers to this story...
		If piLocTypeID = ObjectType.eGalaxy Then
			If sResult.StartsWith("2-") = False AndAlso sResult.Contains("-2-") = False Then
				sResult &= "2-"
			End If
		End If

		'must always end with a -
		If sResult.EndsWith("-") = False Then sResult &= "-"
		Return sResult
	End Function

	Private Shared Function MakeDBStr(ByVal sStr As String) As String
		Return sStr.Replace("'", "''")
	End Function
End Class
