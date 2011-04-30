'Option Strict On

'Public Enum eySenateRequestDetailsType As Byte
'	SystemVote = 0
'	SenateProposal = 1
'	ProposalMsg = 2
'End Enum

'''' <summary>
'''' Contains all of the functions needed for senate management
'''' </summary>
'''' <remarks></remarks>
'Public Class Senate

'	Private Shared moProposals() As SenateProposal
'	Private Shared mlProposalUB As Int32 = -1

'	Public Shared Function GetSaveObjectText() As String
'		Dim oSB As New System.Text.StringBuilder
'		For X As Int32 = 0 To mlProposalUB
'			If moProposals(X) Is Nothing = False Then
'				oSB.AppendLine(moProposals(X).GetSaveObjectText())
'			End If
'		Next X
'		Return oSB.ToString
'	End Function

'	Public Shared Sub AddNewProposal(ByRef oProposal As SenateProposal)
'		If moProposals Is Nothing Then ReDim moProposals(-1)
'		SyncLock moProposals
'			mlProposalUB += 1
'			ReDim Preserve moProposals(mlProposalUB)
'			moProposals(mlProposalUB) = oProposal
'		End SyncLock
'	End Sub

'	Public Shared Function HandleCreateProposalMsg(ByRef yData() As Byte, ByVal lProposerID As Int32) As Byte()
'		Dim lPos As Int32 = 2
'		Dim lTitleLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'		If lTitleLen > 255 Then Return Nothing
'		Dim sTitle As String = GetStringFromBytes(yData, lPos, lTitleLen) : lPos += lTitleLen
'		Dim lDescLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'		If lDescLen > 1000 Then Return Nothing
'		Dim sDesc As String = GetStringFromBytes(yData, lPos, lDescLen) : lPos += lDescLen

'		Dim lProposedOn As Int32 = GetDateAsNumber(Now.ToUniversalTime)

'		Dim oProposal As New SenateProposal()
'		With oProposal
'			.lProposedBy = lProposerID
'			.ObjectID = -1
'			.ObjTypeID = ObjectType.eSenateLaw
'			.yDescription = StringToBytes(sDesc)
'			.yProposalState = eyProposalState.EmperorsChamber
'			.yTitle = StringToBytes(sTitle)
'			.lProposedOn = lProposedOn
'			If .SaveObject() = False Then Return Nothing
'		End With
'		AddNewProposal(oProposal)

'		Return oProposal.GetSmallMsg()
'	End Function

'	Public Shared Function HandleRequestSenate(ByVal bIncludeEmpChamber As Boolean) As Byte()
'		'ok, here, we need to send the player all of the msgs...
'		Dim yCache(200000) As Byte
'		Dim yFinal() As Byte
'		Dim lPos As Int32
'		Dim lSingleMsgLen As Int32

'		Dim yTemp() As Byte

'		lPos = 0
'		lSingleMsgLen = -1
'		For X As Int32 = 0 To mlProposalUB
'			If moProposals(X) Is Nothing = False AndAlso (bIncludeEmpChamber = True OrElse (moProposals(X).yProposalState And eyProposalState.OnSenateFloor) <> 0) Then
'				yTemp = moProposals(X).GetSmallMsg()
'				lSingleMsgLen = yTemp.Length
'				'Ok, before we continue, check if we need to increase our cache
'				If lPos + lSingleMsgLen + 2 > yCache.Length Then
'					'increase it
'					ReDim Preserve yCache(yCache.Length + 200000)
'				End If
'				System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
'				lPos += 2
'				yTemp.CopyTo(yCache, lPos)
'				lPos += lSingleMsgLen
'			End If
'		Next X
'		If lPos <> 0 Then
'			ReDim yFinal(lPos - 1)
'			Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
'			Return yFinal
'		End If
'		Return Nothing
'	End Function

'	Public Shared Sub HandleAddProposalMessage(ByRef yData() As Byte, ByVal lPlayerID As Int32)
'		Dim lPos As Int32 = 2	'for msgcode
'		Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'		Dim lMsgLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'		If lMsgLen > 500 Then
'			LogEvent(LogEventType.PossibleCheat, "Player posting message larger than 500 characters: " & lPlayerID)
'			Return
'		End If
'		Dim sMsg As String = GetStringFromBytes(yData, lPos, lMsgLen)
'		Dim lPostedOn As Int32 = GetDateAsNumber(Now.ToUniversalTime)
'		Dim oProposal As SenateProposal = GetProposal(lProposalID)
'		If oProposal Is Nothing = False Then
'			oProposal.AddMsg(lPlayerID, lPostedOn, sMsg)
'		End If
'	End Sub

'	Public Shared Sub SetPlayerVote(ByVal lProposalID As Int32, ByVal yVoteValue As eyVoteValue, ByVal lPlayerID As Int32)
'		For X As Int32 = 0 To mlProposalUB
'			If moProposals(X) Is Nothing = False AndAlso moProposals(X).ObjectID = lProposalID Then
'				moProposals(X).HandlePlayerVote(lPlayerID, yVoteValue, False)
'				Exit For
'			End If
'		Next X
'	End Sub

'	Public Shared Function GetProposal(ByVal lProposalID As Int32) As SenateProposal
'		For X As Int32 = 0 To mlProposalUB
'			If moProposals(X) Is Nothing = False AndAlso moProposals(X).ObjectID = lProposalID Then
'				Return moProposals(X)
'			End If
'		Next X
'		Return Nothing
'	End Function

'	Public Shared Sub RecalculateAllProposals()
'		For X As Int32 = 0 To mlProposalUB
'			If moProposals(X) Is Nothing = False Then
'				moProposals(X).Recalculate()
'			End If
'		Next X
'	End Sub

'	Public Shared Function HandleGetSenateObjectDetails(ByRef yData() As Byte, ByVal lPlayerID As Int32) As Byte()
'		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
'		If oPlayer Is Nothing Then Return Nothing
'		Dim lPos As Int32 = 2	'for msgcode
'		Dim yType As Byte = yData(lPos) : lPos += 1
'		Select Case yType
'			Case eySenateRequestDetailsType.ProposalMsg
'				Dim lPosterID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'				Dim lPostedOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'				Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'				Dim oProposal As SenateProposal = GetProposal(lProposalID)
'				If oProposal Is Nothing Then Return Nothing
'				If (oProposal.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
'					If oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor Then
'						LogEvent(LogEventType.PossibleCheat, "Player requesting details from proposal without access: " & lPlayerID)
'						Return Nothing
'					End If
'				End If
'				Return oProposal.GetProposalMsgDetails(lPosterID, lPostedOn)
'			Case eySenateRequestDetailsType.SenateProposal
'				Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'				Dim oProposal As SenateProposal = GetProposal(lProposalID)
'				If oProposal Is Nothing Then Return Nothing
'				If (oProposal.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
'					If oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor Then
'						LogEvent(LogEventType.PossibleCheat, "Player requesting details from proposal without access: " & lPlayerID)
'						Return Nothing
'					End If
'				End If
'				Return oProposal.GetBigMsg()
'			Case eySenateRequestDetailsType.SystemVote
'				Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'				Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'				Dim oProposal As SenateProposal = GetProposal(lProposalID)
'				If oProposal Is Nothing Then Return Nothing
'				If (oProposal.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
'					If oPlayer.yPlayerTitle <> Player.PlayerRank.Emperor Then
'						LogEvent(LogEventType.PossibleCheat, "Player requesting details from proposal without access: " & lPlayerID)
'						Return Nothing
'					End If
'				End If
'				Return oProposal.GetSystemVoteDetails(lSystemID)
'		End Select
'		Return Nothing
'	End Function
'End Class
