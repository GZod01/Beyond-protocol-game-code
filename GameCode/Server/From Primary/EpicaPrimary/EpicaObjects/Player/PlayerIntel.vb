Option Strict On

Public Class PlayerIntel
    Inherits Epica_GUID

    Public Enum eIntelStaticVariables As Integer
        eHomeworldKnowledge = 1
    End Enum

    'ObjectID is the TargetPlayer

    Public lPlayerID As Int32               'player who has this intel
    Public StaticVariables As Int32         'static variables
    Public TechnologyScore As Int32 = Int32.MinValue
    Public TechnologyUpdate As Int32
    Public DiplomacyScore As Int32 = Int32.MinValue
    Public DiplomacyUpdate As Int32
    Public MilitaryScore As Int32 = Int32.MinValue
    Public MilitaryUpdate As Int32
    Public PopulationScore As Int32 = Int32.MinValue
    Public PopulationUpdate As Int32
    Public ProductionScore As Int32 = Int32.MinValue
    Public ProductionUpdate As Int32
    Public WealthScore As Int32 = Int32.MinValue
    Public WealthUpdate As Int32

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try

            'Try an update first...
            sSQL = "UPDATE tblPlayerIntel SET StaticVariables = " & StaticVariables & ", TechnologyScore = " & _
              TechnologyScore & ", TechnologyUpdate = " & TechnologyUpdate & ", DiplomacyScore = " & DiplomacyScore & _
              ", DiplomacyUpdate = " & DiplomacyUpdate & ", MilitaryScore = " & MilitaryScore & ", MilitaryUpdate = " & _
              MilitaryUpdate & ", PopulationScore = " & PopulationScore & ", PopulationUpdate = " & PopulationUpdate & _
              ", ProductionScore = " & ProductionScore & ", ProductionUpdate = " & ProductionUpdate & ", WealthScore = " & _
              WealthScore & ", WealthUpdate = " & WealthUpdate & " WHERE PlayerID = " & lPlayerID & " AND TargetID = " & _
              Me.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                oComm = Nothing
                sSQL = "INSERT INTO tblPlayerIntel (PlayerID, TargetID, StaticVariables, TechnologyScore, TechnologyUpdate, " & _
                  "DiplomacyScore, DiplomacyUpdate, MilitaryScore, MilitaryUpdate, PopulationScore, PopulationUpdate, " & _
                  "ProductionScore, ProductionUpdate, WealthScore, WealthUpdate) VALUES (" & lPlayerID & ", " & Me.ObjectID & _
                  ", " & StaticVariables & ", " & TechnologyScore & ", " & TechnologyUpdate & ", " & DiplomacyScore & _
                  ", " & DiplomacyUpdate & ", " & MilitaryScore & ", " & MilitaryUpdate & ", " & PopulationScore & _
                  ", " & PopulationUpdate & ", " & ProductionScore & ", " & ProductionUpdate & ", " & WealthScore & _
                  ", " & WealthUpdate & ")"

                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save Player Intel (player = " & lPlayerID & ", target = " & ObjectID & "). Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Private moTarget As Player = Nothing
    Public ReadOnly Property oTarget() As Player
        Get
            If moTarget Is Nothing Then
                moTarget = GetEpicaPlayer(Me.ObjectID)
            End If
            Return moTarget
        End Get
    End Property

	Private Sub CheckForBloodAlly()
		Dim oOther As Player = oTarget
		If oOther Is Nothing = False Then
			Dim oRel As PlayerRel = oOther.GetPlayerRel(lPlayerID)
            If oRel Is Nothing = False AndAlso oRel.WithThisScore > elRelTypes.eAlly Then
                'ok, force the score update
                Dim lNow As Int32 = GetDateAsNumber(Now)

                With oOther
                    'If .lLastGlobalRequestTurnIn = 0 OrElse glCurrentCycle - .lLastGlobalRequestTurnIn > 9000 Then
                    '    goMsgSys.SendRequestGlobalPlayerScores(.ObjectID, 64, .ObjectID)
                    '    If .lLastGlobalRequestTurnIn = 0 Then
                    .lLGTechScore = .TechnologyScore
                    .lLGDiplomacyScore = .DiplomacyScore
                    .lLGWealthScore = .WealthScore
                    .lLGProductionScore = .ProductionScore
                    .lLGPopulationScore = .PopulationScore
                    .lLGMilitaryScore = .lMilitaryScore
                    .lLGTotalScore = .TotalScore
                    '    End If
                    'End If
                End With

                TechnologyScore = oOther.lLGTechScore : TechnologyUpdate = lNow
                DiplomacyScore = oOther.lLGDiplomacyScore : DiplomacyUpdate = lNow
                MilitaryScore = oOther.lLGMilitaryScore \ 50 : MilitaryUpdate = lNow
                PopulationScore = oOther.lLGPopulationScore : PopulationUpdate = lNow
                ProductionScore = oOther.lLGProductionScore : ProductionUpdate = lNow
                WealthScore = oOther.lLGWealthScore : WealthUpdate = lNow
            End If
		End If
	End Sub

    Public Function GetObjAsString() As Byte()
        Dim yMsg(85) As Byte
		Dim lPos As Int32 = 0

		'ok, before proceeding, check the player to player rel
		CheckForBloodAlly()

        Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4

        If (StaticVariables And eIntelStaticVariables.eHomeworldKnowledge) <> 0 Then
            System.BitConverter.GetBytes(oTarget.lStartedEnvirID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(oTarget.iStartedEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
        Else
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
        End If

		System.BitConverter.GetBytes(TechnologyScore).CopyTo(yMsg, lPos) : lPos += 4
		Dim lSeconds As Int32 = 0
		If TechnologyUpdate > 0 Then lSeconds = CInt(Now.Subtract(GetDateFromNumber(TechnologyUpdate)).TotalSeconds) Else lSeconds = 0
		System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(DiplomacyScore).CopyTo(yMsg, lPos) : lPos += 4
		If DiplomacyUpdate > 0 Then lSeconds = CInt(Now.Subtract(GetDateFromNumber(DiplomacyUpdate)).TotalSeconds) Else lSeconds = 0
		System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(MilitaryScore).CopyTo(yMsg, lPos) : lPos += 4
		If MilitaryUpdate > 0 Then lSeconds = CInt(Now.Subtract(GetDateFromNumber(MilitaryUpdate)).TotalSeconds) Else lSeconds = 0
		System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(PopulationScore).CopyTo(yMsg, lPos) : lPos += 4
		If PopulationUpdate > 0 Then lSeconds = CInt(Now.Subtract(GetDateFromNumber(PopulationUpdate)).TotalSeconds) Else lSeconds = 0
		System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(ProductionScore).CopyTo(yMsg, lPos) : lPos += 4
		If ProductionUpdate > 0 Then lSeconds = CInt(Now.Subtract(GetDateFromNumber(ProductionUpdate)).TotalSeconds) Else lSeconds = 0
		System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(WealthScore).CopyTo(yMsg, lPos) : lPos += 4
		If WealthUpdate > 0 Then lSeconds = CInt(Now.Subtract(GetDateFromNumber(WealthUpdate)).TotalSeconds) Else lSeconds = 0
        System.BitConverter.GetBytes(lSeconds).CopyTo(yMsg, lPos) : lPos += 4

        With oTarget
            'If .lLastGlobalRequestTurnIn = 0 OrElse glCurrentCycle - .lLastGlobalRequestTurnIn > 9000 Then
            'goMsgSys.SendRequestGlobalPlayerScores(.ObjectID, 64, .ObjectID)
            'If .lLastGlobalRequestTurnIn = 0 Then
            .lLGTotalScore = .TotalScore
            'End If
            'End If
        End With
        System.BitConverter.GetBytes(oTarget.lLGTotalScore).CopyTo(yMsg, lPos) : lPos += 4

        Dim oRel As PlayerRel = oTarget.GetPlayerRel(Me.lPlayerID)
        If oRel Is Nothing = False Then
            yMsg(lPos) = oRel.WithThisScore
        ElseIf Me.lPlayerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
            yMsg(lPos) = elRelTypes.eBloodWar
        Else
            yMsg(lPos) = elRelTypes.eNeutral
        End If
        lPos += 1

        yMsg(lPos) = oTarget.yGender : lPos += 1

        System.BitConverter.GetBytes(oTarget.lPlayerIcon).CopyTo(yMsg, lPos) : lPos += 4

        yMsg(lPos) = oTarget.yPlayerTitle : lPos += 1


        Dim lTemp As Int32 = oTarget.lCelebrationEnds - glCurrentCycle
        If lTemp < 0 Then lTemp = 0
        System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, lPos) : lPos += 4

        yMsg(lPos) = oTarget.yCustomTitle : lPos += 1

        System.BitConverter.GetBytes(oTarget.lGuildID).CopyTo(yMsg, lPos) : lPos += 4

        System.BitConverter.GetBytes(CShort(oTarget.lTotalSenateVoteStrength)).CopyTo(yMsg, lPos) : lPos += 2

        Return yMsg

    End Function

	Public Function CloneForPlayer(ByRef oOtherPlayer As Player) As PlayerIntel
        Dim oPI As PlayerIntel = oOtherPlayer.GetOrAddPlayerIntel(ObjectID, False)
		If oPI Is Nothing = False Then
			With Me
				oPI.lPlayerID = oOtherPlayer.ObjectID
				If .TechnologyScore <> Int32.MinValue AndAlso .TechnologyUpdate > oPI.TechnologyUpdate Then
					oPI.TechnologyScore = .TechnologyScore
					oPI.TechnologyUpdate = .TechnologyUpdate
				End If
				If .DiplomacyScore <> Int32.MinValue AndAlso .DiplomacyUpdate > oPI.DiplomacyUpdate Then
					oPI.DiplomacyScore = .DiplomacyScore
					oPI.DiplomacyUpdate = .DiplomacyUpdate
				End If
				If .MilitaryScore <> Int32.MinValue AndAlso .MilitaryUpdate > oPI.MilitaryUpdate Then
                    oPI.MilitaryScore = .MilitaryScore
					oPI.MilitaryUpdate = .MilitaryUpdate
				End If
				If .PopulationScore <> Int32.MinValue AndAlso .PopulationUpdate > oPI.PopulationUpdate Then
					oPI.PopulationScore = .PopulationScore
					oPI.PopulationUpdate = .PopulationUpdate
				End If
				If .ProductionScore <> Int32.MinValue AndAlso .ProductionUpdate > oPI.ProductionUpdate Then
					oPI.ProductionScore = .ProductionScore
					oPI.ProductionUpdate = .ProductionUpdate
				End If
				If .WealthScore <> Int32.MinValue AndAlso .WealthUpdate > oPI.WealthUpdate Then
					oPI.WealthScore = .WealthScore
					oPI.WealthUpdate = .WealthUpdate
				End If

				oPI.StaticVariables = oPI.StaticVariables Or .StaticVariables
			End With
			oPI.SaveObject()
		End If
		Return oPI
	End Function

	Public Shared Function DeleteAllPlayerIntels(ByVal lPlayerID As Int32) As Boolean
		Dim sSQL As String = ""
		Dim oComm As OleDb.OleDbCommand = Nothing
		Dim bResult As Boolean = False

		Try
			sSQL = "DELETE FROM tblPlayerIntel WHERE PlayerID = " & lPlayerID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm.Dispose()
			oComm = Nothing

			bResult = True
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, ex.Message)
		Finally
			If oComm Is Nothing = False Then oComm.Dispose()
			oComm = Nothing
		End Try

		Return bResult
	End Function
End Class
