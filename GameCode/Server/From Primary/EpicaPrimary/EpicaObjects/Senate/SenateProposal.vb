'Option Strict On

'Public Enum eyProposalState As Byte
'	EmperorsChamber = 0
'	AwaitingApproval = 1		'emperor's chamber voting is done, waiting approval from devs
'	OnSenateFloor = 2			'dev approved on senate floor ready for votes
'	DevDeclined = 4				'developers declined, returned to emperors chamber
'	SmallMsgBitShift = 128
'End Enum

'Public Class SenateProposal
'	Inherits Epica_GUID

'	Public yProposalState As eyProposalState
'	Public yTitle() As Byte
'	Public yDescription() As Byte

'	Public lProposedBy As Int32
'	Public lProposedOn As Int32			'date as number in gmt

'	Public lVotesFor As Int32 = 0		'not saved, cummulated over time
'	Public lVotesAgainst As Int32 = 0	'not saved, cummulated over time

'	Private muMsgs() As SenateProposalMessage
'	Private moVotes() As SenateVote

'	Private moSystemVotes() As SystemVote		'not saved, kept track of as players change their votes and planet/system ownership changes
'	Private Class SystemVote
'		Public SystemID As Int32

'		Public PlanetsFor As Int32 = 0
'		Public PlanetsAgainst As Int32 = 0

'		Private moSystem As SolarSystem = Nothing
'		Public ReadOnly Property oSystem() As SolarSystem
'			Get
'				If moSystem Is Nothing OrElse moSystem.ObjectID <> SystemID Then
'					moSystem = GetEpicaSystem(SystemID)
'				End If
'				Return moSystem
'			End Get
'		End Property
'		Public ReadOnly Property ySystemVote() As eyVoteValue
'			Get
'				Dim yResult As eyVoteValue = eyVoteValue.AbstainVote
'				If oSystem Is Nothing = False Then
'					Dim lHalf As Int32 = (oSystem.mlPlanetUB + 1) \ 2
'					If PlanetsFor > lHalf Then
'						yResult = eyVoteValue.YesVote
'					ElseIf PlanetsAgainst > lHalf Then
'						yResult = eyVoteValue.NoVote
'					End If
'				End If
'			End Get
'		End Property
'		Private myVoteRating As Byte = 255
'		Public ReadOnly Property VoteRating() As Byte
'			Get
'				If myVoteRating = 255 Then
'					Dim oSystem As SolarSystem = GetEpicaSystem(SystemID)
'					If oSystem Is Nothing = False Then
'						myVoteRating = CByte(oSystem.mlPlanetUB + 1)
'					End If
'				End If
'				Return myVoteRating
'			End Get
'		End Property
'	End Class

'	Public Function GetSmallMsg() As Byte()
'		Dim lTitleLen As Int32 = 0
'		If yTitle Is Nothing = False Then lTitleLen = yTitle.Length

'		Dim yMsg(24 + lTitleLen) As Byte
'		Dim lPos As Int32 = 0
'		System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
'		Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
'		yMsg(lPos) = yProposalState Or eyProposalState.SmallMsgBitShift : lPos += 1
'		System.BitConverter.GetBytes(lVotesFor).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(lVotesAgainst).CopyTo(yMsg, lPos) : lPos += 4

'		Dim lMsgCnt As Int32 = 0
'		If muMsgs Is Nothing = False Then lMsgCnt = muMsgs.Length
'		System.BitConverter.GetBytes(lMsgCnt).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(lTitleLen).CopyTo(yMsg, lPos) : lPos += 4
'		If yTitle Is Nothing = False Then yTitle.CopyTo(yMsg, lPos) : lPos += lTitleLen

'		Return yMsg
'	End Function

'	Public Function GetBigMsg() As Byte()
'		'Ok, if the state is in the emperor's chamber, then we do simple votes
'		'however, if the state is in Senate Floor, we categorize the votes by system
'		Dim lTitleLen As Int32 = 0
'		If yTitle Is Nothing = False Then lTitleLen = yTitle.Length
'		Dim lDescLen As Int32 = 0
'		If yDescription Is Nothing = False Then lDescLen = yDescription.Length
'		Dim lMsgCnt As Int32 = 0
'		If muMsgs Is Nothing = False Then lMsgCnt = muMsgs.Length

'		Dim lPerVoteLen As Int32 = 0
'		Dim lSysVoteCnt As Int32 = 0

'		If moVotes Is Nothing = False Then
'			If (Me.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
'				'Emperor's Chambers
'				lPerVoteLen = 21
'				lSysVoteCnt = 0
'				If moVotes Is Nothing = False Then lSysVoteCnt = moVotes.Length
'			Else
'				'On senate floor... we need to categorize by system
'				lPerVoteLen = 6
'				lSysVoteCnt = 0
'				If moSystemVotes Is Nothing = False Then lSysVoteCnt = moSystemVotes.Length
'			End If
'		End If

'		Dim yMsg(52 + lTitleLen + lDescLen + (lSysVoteCnt * lPerVoteLen) + (lMsgCnt * 28)) As Byte
'		Dim lPos As Int32 = 0

'		System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
'		Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
'		yMsg(lPos) = yProposalState : lPos += 1

'		System.BitConverter.GetBytes(lTitleLen).CopyTo(yMsg, lPos) : lPos += 4
'		If yTitle Is Nothing = False Then yTitle.CopyTo(yMsg, lPos)
'		lPos += lTitleLen

'		System.BitConverter.GetBytes(lDescLen).CopyTo(yMsg, lPos) : lPos += 4
'		If yDescription Is Nothing = False Then yDescription.CopyTo(yMsg, lPos)
'		lPos += lDescLen

'		If ProposedBy Is Nothing = False Then ProposedBy.PlayerName.CopyTo(yMsg, lPos)
'		lPos += 20

'		System.BitConverter.GetBytes(lProposedBy).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(lProposedOn).CopyTo(yMsg, lPos) : lPos += 4

'		System.BitConverter.GetBytes(lSysVoteCnt).CopyTo(yMsg, lPos) : lPos += 4
'		If (Me.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
'			'Emperors Chambers
'			If moVotes Is Nothing = False Then
'				For X As Int32 = 0 To moVotes.GetUpperBound(0)
'					If moVotes(X) Is Nothing = False Then
'						With moVotes(X)
'							If .Voter Is Nothing = False Then
'								.Voter.PlayerName.CopyTo(yMsg, lPos)
'							End If
'							lPos += 20
'							yMsg(lPos) = .yVote : lPos += 1
'						End With
'					End If
'				Next X
'			End If
'		Else
'			'On Senate Floor
'			If moSystemVotes Is Nothing = False Then
'				For X As Int32 = 0 To moSystemVotes.GetUpperBound(0)
'					If moSystemVotes(X) Is Nothing = False Then
'						With moSystemVotes(X)
'							System.BitConverter.GetBytes(.SystemID).CopyTo(yMsg, lPos) : lPos += 4
'							yMsg(lPos) = .ySystemVote : lPos += 1
'							yMsg(lPos) = .VoteRating : lPos += 1
'						End With
'					End If
'				Next X
'			End If
'		End If

'		System.BitConverter.GetBytes(lMsgCnt).CopyTo(yMsg, lPos) : lPos += 4
'		If muMsgs Is Nothing = False Then
'			For X As Int32 = 0 To muMsgs.GetUpperBound(0)
'				With muMsgs(X)
'					Dim oPoster As Player = .PostedBy
'					If oPoster Is Nothing = False Then oPoster.PlayerName.CopyTo(yMsg, lPos)
'					lPos += 20
'					System.BitConverter.GetBytes(.lPosterID).CopyTo(yMsg, lPos) : lPos += 4
'					System.BitConverter.GetBytes(.lPostedOn).CopyTo(yMsg, lPos) : lPos += 4
'				End With
'			Next X
'		End If

'		Return yMsg
'	End Function

'	Private moProposedBy As Player = Nothing
'	Public ReadOnly Property ProposedBy() As Player
'		Get
'			If moProposedBy Is Nothing OrElse moProposedBy.ObjectID <> lProposedBy Then
'				moProposedBy = GetEpicaPlayer(lProposedBy)
'			End If
'			Return moProposedBy
'		End Get
'	End Property

'	Public Function SaveObject() As Boolean
'		Dim bResult As Boolean = False
'		Dim sSQL As String
'		Dim oComm As OleDb.OleDbCommand

'		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

'		Try
'			If ObjectID = -1 Then
'				'INSERT
'				sSQL = "INSERT INTO tblSenateProposal (ProposalState, ProposalTitle, ProposalDescription, ProposedBy, ProposedOn) VALUES ("
'				sSQL &= CByte(yProposalState) & ", '" & MakeDBStr(BytesToString(yTitle)) & "', '" & MakeDBStr(BytesToString(yDescription)) & _
'				   "', " & lProposedBy & ", " & lProposedOn & ")"
'			Else
'				'UPDATE
'				sSQL = "UPDATE tblSenateProposal SET ProposalState = " & CByte(yProposalState) & ", ProposalTitle = '" & _
'				 MakeDBStr(BytesToString(yTitle)) & "', ProposalDescription = '" & MakeDBStr(BytesToString(yDescription)) & _
'				 "', ProposedBy = " & lProposedBy & ", ProposedOn = " & lProposedOn & " WHERE ProposalID = " & Me.ObjectID
'			End If
'			oComm = New OleDb.OleDbCommand(sSQL, goCN)
'			If oComm.ExecuteNonQuery() = 0 Then
'				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
'			End If
'			If ObjectID = -1 Then
'				Dim oData As OleDb.OleDbDataReader
'				oComm = Nothing
'				sSQL = "SELECT MAX(ProposalID) FROM tblSenateProposal WHERE ProposedBy = " & lProposedBy
'				oComm = New OleDb.OleDbCommand(sSQL, goCN)
'				oData = oComm.ExecuteReader(CommandBehavior.Default)
'				If oData.Read Then
'					ObjectID = CInt(oData(0))
'				End If
'				oData.Close()
'				oData = Nothing
'			End If

'			sSQL = "DELETE FROM tblSenateProposalMsg WHERE ProposalID = " & Me.ObjectID
'			oComm = New OleDb.OleDbCommand(sSQL, goCN)
'			oComm.ExecuteNonQuery()
'			oComm = Nothing

'			sSQL = "DELETE FROM tblSenateProposalVote WHERE ProposalID = " & Me.ObjectID
'			oComm = New OleDb.OleDbCommand(sSQL, goCN)
'			oComm.ExecuteNonQuery()
'			oComm = Nothing

'			If muMsgs Is Nothing = False Then
'				For X As Int32 = 0 To muMsgs.GetUpperBound(0)
'					If muMsgs(X).SaveObject(Me.ObjectID) = False Then
'						LogEvent(LogEventType.CriticalError, "Could not save senate proposal msg for proposalid: " & Me.ObjectID)
'					End If
'				Next X
'			End If

'			If moVotes Is Nothing = False Then
'				For X As Int32 = 0 To moVotes.GetUpperBound(0)
'					If moVotes(X).SaveObject(Me.ObjectID) = False Then
'						LogEvent(LogEventType.CriticalError, "Could not save senate system vote for proposalid: " & Me.ObjectID)
'					End If
'				Next X
'			End If

'			bResult = True
'		Catch
'			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
'		Finally
'			oComm = Nothing
'		End Try
'		Return bResult
'	End Function

'	Public Function GetSaveObjectText() As String
'		Dim sSQL As String

'		If ObjectID = -1 Then
'			SaveObject()
'			Return ""
'		End If

'		Try
'			Dim oSB As New System.Text.StringBuilder

'			'UPDATE
'			sSQL = "UPDATE tblSenateProposal SET ProposalState = " & CByte(yProposalState) & ", ProposalTitle = '" & _
'			 MakeDBStr(BytesToString(yTitle)) & "', ProposalDescription = '" & MakeDBStr(BytesToString(yDescription)) & _
'			 "', ProposedBy = " & lProposedBy & ", ProposedOn = " & lProposedOn & " WHERE ProposalID = " & Me.ObjectID
'			oSB.AppendLine(sSQL)

'			sSQL = "DELETE FROM tblSenateProposalMsg WHERE ProposalID = " & Me.ObjectID
'			oSB.AppendLine(sSQL)

'			sSQL = "DELETE FROM tblSenateProposalVote WHERE ProposalID = " & Me.ObjectID
'			oSB.AppendLine(sSQL)

'			If muMsgs Is Nothing = False Then
'				For X As Int32 = 0 To muMsgs.GetUpperBound(0)
'					oSB.AppendLine(muMsgs(X).GetSaveObjectText(Me.ObjectID))
'				Next X
'			End If

'			If moVotes Is Nothing = False Then
'				For X As Int32 = 0 To moVotes.GetUpperBound(0)
'					oSB.AppendLine(moVotes(X).GetSaveObjectText(Me.ObjectID))
'				Next X
'			End If

'			Return oSB.ToString
'		Catch
'			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
'		End Try
'		Return ""

'	End Function

'	Public Sub HandlePlayerVote(ByVal lPlayerID As Int32, ByVal yVoteValue As eyVoteValue, ByVal bNoRecalculate As Boolean)
'		If yVoteValue = eyVoteValue.AbstainVote Then
'			LogEvent(LogEventType.PossibleCheat, "Vote of Abstain: " & lPlayerID)
'			Return
'		End If
'		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
'		If oPlayer Is Nothing Then Return

'		'Ok, first, is this vote on the senate floor?
'		If (Me.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
'			'No, ok, is the player an emperor?
'			If oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor Then
'				LogEvent(LogEventType.PossibleCheat, "Non-Emperor voting in emperor's chambers: " & lPlayerID)
'				Return
'			End If
'		End If

'		If moVotes Is Nothing Then ReDim moVotes(-1)
'		For X As Int32 = 0 To moVotes.GetUpperBound(0)
'			If moVotes(X) Is Nothing = False AndAlso moVotes(X).VoterID = lPlayerID Then
'				If moVotes(X).yVote <> yVoteValue Then
'					'ok, player is changing their vote...
'					If moVotes(X).yVote = eyVoteValue.YesVote Then
'						lVotesFor -= 1
'					ElseIf moVotes(X).yVote = eyVoteValue.NoVote Then
'						lVotesAgainst -= 1
'					End If
'					moVotes(X).yVote = yVoteValue
'					If yVoteValue = eyVoteValue.YesVote Then
'						lVotesFor += 1
'					ElseIf yVoteValue = eyVoteValue.NoVote Then
'						lVotesAgainst -= 1
'					End If
'					If bNoRecalculate = False Then Recalculate()
'				End If
'				Return
'			End If
'		Next X

'		'Ok, we are here, add a new vote
'		Dim oVote As New SenateVote()
'		With oVote
'			.VoterID = lPlayerID
'			.yVote = yVoteValue
'		End With
'		ReDim Preserve moVotes(moVotes.GetUpperBound(0) + 1)
'		moVotes(moVotes.GetUpperBound(0)) = oVote

'		If yVoteValue = eyVoteValue.YesVote Then
'			lVotesFor += 1
'		ElseIf yVoteValue = eyVoteValue.NoVote Then
'			lVotesAgainst -= 1
'		End If

'		'Ok, a new vote... need to determine what the player voting owns...
'		If bNoRecalculate = False Then Recalculate()
'	End Sub

'	Public Sub Recalculate()
'		'ok, here, we need to calculate our vote out and recalculate our system votes
'		Dim oSysVotes(-1) As SystemVote

'		If moVotes Is Nothing = False Then
'			For X As Int32 = 0 To moVotes.GetUpperBound(0)
'				If moVotes(X) Is Nothing = False AndAlso (moVotes(X).yVote = eyVoteValue.YesVote OrElse moVotes(X).yVote = eyVoteValue.NoVote) Then
'					'ok, need to get this guy's controlled environments
'					Dim oPlanets() As Planet = moVotes(X).Voter.GetControlledPlanetList()
'					If oPlanets Is Nothing = False Then
'						'ok, got our planet list
'						For Y As Int32 = 0 To oPlanets.GetUpperBound(0)
'							If oPlanets(Y) Is Nothing = False AndAlso oPlanets(Y).ObjectID < 500000000 Then
'								'ok, this planet follows this player's vote
'								Dim lIdx As Int32 = -1
'								Dim lSysID As Int32 = oPlanets(Y).ParentSystem.ObjectID
'								For Z As Int32 = 0 To oSysVotes.GetUpperBound(0)
'									If oSysVotes(Z).SystemID = lSysID Then
'										lIdx = Z
'										Exit For
'									End If
'								Next Z
'								If lIdx = -1 Then
'									lIdx = oSysVotes.GetUpperBound(0) + 1
'									ReDim Preserve oSysVotes(lIdx)
'									oSysVotes(lIdx) = New SystemVote
'									oSysVotes(lIdx).SystemID = lSysID
'								End If
'								If moVotes(X).yVote = eyVoteValue.YesVote Then
'									oSysVotes(lIdx).PlanetsFor += 1
'								Else
'									oSysVotes(lIdx).PlanetsAgainst += 1
'								End If
'							End If
'						Next Y
'					End If
'				End If
'			Next X
'		End If

'		moSystemVotes = oSysVotes
'	End Sub

'	Public Sub AddMsg(ByVal lPosterID As Int32, ByVal lPostedOn As Int32, ByVal sMsgData As String)
'		If muMsgs Is Nothing Then ReDim muMsgs(-1)

'		Dim oMsg As New SenateProposalMessage()
'		With oMsg
'			.lPostedOn = lPostedOn
'			.lPosterID = lPosterID
'			.yMsgData = StringToBytes(sMsgData)
'		End With
'		ReDim Preserve muMsgs(muMsgs.GetUpperBound(0) + 1)
'		muMsgs(muMsgs.GetUpperBound(0)) = oMsg
'	End Sub

'	Public Function GetProposalMsgDetails(ByVal lPosterID As Int32, ByVal lPostedOn As Int32) As Byte()
'		If muMsgs Is Nothing = False Then
'			For X As Int32 = 0 To muMsgs.GetUpperBound(0)
'				If muMsgs(X).lPosterID = lPosterID AndAlso muMsgs(X).lPostedOn = lPostedOn Then
'					Return muMsgs(X).GetMsgDetail(Me.ObjectID)
'				End If
'			Next X
'		End If
'		Return Nothing
'	End Function

'	Public Function GetSystemVoteDetails(ByVal lSystemID As Int32) As Byte()
'		If moSystemVotes Is Nothing = False Then
'			For X As Int32 = 0 To moSystemVotes.GetUpperBound(0)
'				If moSystemVotes(X) Is Nothing = False AndAlso moSystemVotes(X).SystemID = lSystemID Then
'					Dim oSystem As SolarSystem = moSystemVotes(X).oSystem
'					If oSystem Is Nothing Then Return Nothing

'					Dim lVoterID(oSystem.mlPlanetUB) As Int32
'					Dim yVoterRating(oSystem.mlPlanetUB) As Byte
'					Dim lVoterUB As Int32 = -1
'					For Y As Int32 = 0 To oSystem.mlPlanetUB
'						Dim lPlanetIdx As Int32 = oSystem.GetPlanetIdx(Y)
'						If glPlanetIdx(lPlanetIdx) > -1 Then
'							Dim oPlanet As Planet = goPlanet(lPlanetIdx)
'							If oPlanet Is Nothing = False Then
'								If oPlanet.OwnerID > 0 Then
'									Dim lIdx As Int32 = -1
'									For Z As Int32 = 0 To lVoterUB
'										If lVoterID(Z) = oPlanet.OwnerID Then
'											lIdx = Z
'											Exit For
'										End If
'									Next Z
'									If lIdx = -1 Then
'										lVoterUB += 1
'										lIdx = lVoterUB
'										lVoterID(lIdx) = oPlanet.OwnerID
'										yVoterRating(lIdx) = 0
'									End If
'									yVoterRating(lIdx) += CByte(1)
'								End If
'							End If
'						End If
'					Next Y

'					Dim lPlanetCnt As Int32 = lVoterUB + 1

'					Dim yMsg(15 + (lPlanetCnt * 26)) As Byte
'					Dim lPos As Int32 = 0
'					System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
'					yMsg(lPos) = eySenateRequestDetailsType.SystemVote : lPos += 1
'					System.BitConverter.GetBytes(lSystemID).CopyTo(yMsg, lPos) : lPos += 4
'					System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4

'					yMsg(lPos) = CByte(oSystem.mlPlanetUB + 1) : lPos += 1
'					System.BitConverter.GetBytes(lPlanetCnt).CopyTo(yMsg, lPos) : lPos += 4

'					For Y As Int32 = 0 To lVoterUB
'						Dim oPlayer As Player = GetEpicaPlayer(lVoterID(Y))
'						If oPlayer Is Nothing = False Then
'							System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'							oPlayer.PlayerName.CopyTo(yMsg, lPos) : lPos += 20
'							Dim oVote As SenateVote = GetVote(oPlayer.ObjectID)
'							If oVote Is Nothing = False Then
'								yMsg(lPos) = oVote.yVote
'							Else : yMsg(lPos) = eyVoteValue.AbstainVote				'TODO: Replace abstain with the default vote
'							End If
'							lPos += 1
'							yMsg(lPos) = yVoterRating(Y) : lPos += 1
'						End If
'					Next Y

'					Return yMsg
'				End If
'			Next X
'		End If
'		Return Nothing
'	End Function

'	Public Function GetVote(ByVal lPlayerID As Int32) As SenateVote
'		If moVotes Is Nothing = False Then
'			For X As Int32 = 0 To moVotes.GetUpperBound(0)
'				If moVotes(X) Is Nothing = False AndAlso moVotes(X).VoterID = lPlayerID Then Return moVotes(X)
'			Next X
'		End If
'		Return Nothing
'	End Function
'End Class
