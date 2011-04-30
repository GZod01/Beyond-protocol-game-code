Option Strict On

Public Class SenateVote
    'Public SystemID As Int32
    'Private moSolarSystem As SolarSystem = Nothing
    'Public ReadOnly Property oSystem() As SolarSystem
    '	Get
    '		If moSolarSystem Is Nothing OrElse moSolarSystem.ObjectID <> SystemID Then
    '			moSolarSystem = GetEpicaSystem(SystemID)
    '		End If
    '		Return moSolarSystem
    '	End Get
    'End Property

    'Public ySystemVote As eyVoteValue

    'Private myVoteRating As Byte = 255
    'Public ReadOnly Property VoteRating() As Byte
    '	Get
    '		If myVoteRating = 255 Then
    '			Dim oSystem As SolarSystem = GetEpicaSystem(SystemID)
    '			If oSystem Is Nothing = False Then
    '				myVoteRating = CByte(oSystem.mlPlanetUB + 1)
    '			End If
    '		End If
    '		Return myVoteRating
    '	End Get
    'End Property

    Public VoterID As Int32
    Public yVote As eyVoteValue
    Public yVotePriority As Byte = 0

    Private moVoter As Player = Nothing
    Public ReadOnly Property Voter() As Player
        Get
            If moVoter Is Nothing OrElse moVoter.ObjectID <> VoterID Then
                moVoter = GetEpicaPlayer(VoterID)
            End If
            Return moVoter
        End Get
    End Property

    'Public Function GetSystemVoteDetail(ByVal lProposalID As Int32) As Byte()
    '	Dim lVoterCnt As Int32 = 0
    '	If muVoters Is Nothing = False Then lVoterCnt = muVoters.Length

    '	Dim yMsg(14 + (lVoterCnt * 21)) As Byte
    '	Dim lPos As Int32 = 0

    '	'System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2		'what msgcode?
    '	System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
    '	System.BitConverter.GetBytes(SystemID).CopyTo(yMsg, lPos) : lPos += 4
    '	yMsg(lPos) = VoteRating : lPos += 1

    '	For X As Int32 = 0 To oSystem.mlPlanetUB
    '		If oSystem.GetPlanetIdx(X) > -1 Then
    '			Dim oPlanet As Planet = goPlanet(oSystem.GetPlanetIdx(X))
    '			If oPlanet Is Nothing = False Then
    '				If oPlanet.OwnerID > -1 Then

    '				End If
    '			End If
    '		End If
    '	Next X

    '	System.BitConverter.GetBytes(lVoterCnt).CopyTo(yMsg, lPos) : lPos += 4
    '	If muVoters Is Nothing = False Then
    '		For X As Int32 = 0 To muVoters.GetUpperBound(0)
    '			With muVoters(X)
    '				Dim oVoter As Player = .Voter
    '				Dim yTitle As Byte = 0
    '				If oVoter Is Nothing = False Then
    '					oVoter.PlayerName.CopyTo(yMsg, lPos)
    '					yTitle = oVoter.yPlayerTitle
    '				End If
    '				lPos += 20

    '				If .yVote = eyVoteValue.YesVote Then
    '					yTitle = CByte(yTitle Or 128)
    '				Else
    '					yTitle = CByte(yTitle Or 64)
    '				End If
    '				yMsg(lPos) = yTitle : lPos += 1
    '			End With
    '		Next X
    '	End If

    '	Return yMsg
    'End Function

    Public Function SaveObject(ByVal lProposalID As Int32) As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try

            Try
                sSQL = "UPDATE tblSenateProposalVote SET VoteValue = " & CByte(yVote) & ", VotePriority = " & yVotePriority & " WHERE ProposalID = " & _
                    lProposalID & " AND VoterID = " & VoterID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    oComm.Dispose()
                    oComm = Nothing

                    sSQL = "INSERT INTO tblSenateProposalVote (ProposalID, VoterID, VoteValue, VotePriority) VALUES (" & lProposalID & ", " & _
                      VoterID & ", " & CByte(yVote) & ", " & yVotePriority & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                    End If
                End If
            Catch ex As Exception
                LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object SenateProposalMessage. Reason: " & ex.Message)
            End Try

            bResult = True
        Catch ex As Exception
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object SenateSystemVote. Reason: " & ex.Message)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function GetSaveObjectText(ByVal lProposalID As Int32) As String
        Try

            Return "INSERT INTO tblSenateProposalVote (ProposalID, VoterID, VoteValue, VotePriority) VALUES (" & lProposalID & ", " & _
             VoterID & ", " & CByte(yVote) & ", " & yVotePriority & ")"
        Catch ex As Exception
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object SenateSystemVote. Reason: " & ex.Message)
        End Try
        Return ""
    End Function

End Class
