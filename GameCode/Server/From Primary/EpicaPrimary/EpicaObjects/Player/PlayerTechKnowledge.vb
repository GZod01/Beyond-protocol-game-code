Option Strict On

''' <summary>
''' This class contains all necessary functionality for managing a player's knowledge of another player's technological designs/components
''' </summary>
''' <remarks></remarks>
Public Class PlayerTechKnowledge
    Public Enum KnowledgeType As Byte
        eNameOnly = 0                   'indicates that only the name of the tech is known
        eSettingsLevel1 = 1             'indicates that only lvl 1 stats are known, this varies per component
        eSettingsLevel2 = 2             'indicates that all stats are known except minerals used
		eFullKnowledge = 3				'indicates all stats are known including minerals used
	End Enum

    Public oPlayer As Player                    'reference to the player who has the knowledge
    Public oTech As Epica_Tech                  'reference to the technology that the knowledge is about
	Public yKnowledgeType As KnowledgeType		'an indication of what knowledge level the player has about the tech

    'Public bArchived As Boolean = False
    Public yArchived As Byte = 0

    Public Function SaveObject() As Boolean
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim bResult As Boolean = False

        Try
            sSQL = "UPDATE tblPlayerTechKnowledge SET TechOwnerID = " & oTech.Owner.ObjectID & ", KnowledgeType = " & _
              CByte(yKnowledgeType) & ", bArchived = " & yArchived & " WHERE PlayerID = " & oPlayer.ObjectID & " AND TechID = " & oTech.ObjectID & _
              " AND TechTypeID = " & oTech.ObjTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                'Ok, try to insert it
                oComm.Dispose()
                oComm = Nothing
                sSQL = "INSERT INTO tblPlayerTechKnowledge (PlayerID, TechID, TechTypeID, TechOwnerID, KnowledgeType, bArchived) VALUES (" & _
                  oPlayer.ObjectID & ", " & oTech.ObjectID & ", " & oTech.ObjTypeID & ", " & oTech.Owner.ObjectID & ", " & _
                  CByte(yKnowledgeType) & ", " & yArchived & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
            End If
            bResult = True
        Catch ex As Exception
            If oPlayer Is Nothing = False AndAlso oTech Is Nothing = False AndAlso oTech.Owner Is Nothing = False Then
                LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object PlayerTechKnowledge (" & oPlayer.ObjectID & ", " & oTech.ObjectID & ", " & oTech.ObjTypeID & "). Reason: " & ex.Message)
            Else
                LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object PlayerTechKnowledge. Reason: " & ex.Message)
            End If
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

        Return bResult
    End Function

	Public Shared Function CreateAndAddToPlayer(ByRef poPlayer As Player, ByRef poTech As Epica_Tech, ByVal yKnowType As KnowledgeType, ByVal pbNoSave As Boolean) As PlayerTechKnowledge
		Dim oPTK As New PlayerTechKnowledge()
		oPTK.oPlayer = poPlayer
		oPTK.oTech = poTech
		oPTK.yKnowledgeType = yKnowType
		If poPlayer Is Nothing = False Then poPlayer.AddPlayerTechKnowledge(oPTK, pbNoSave)
		Return oPTK
	End Function

    Public Sub SendMsgToPlayer()
        If oPlayer Is Nothing = False AndAlso (oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eViewAgents Or AliasingRights.eViewTechDesigns) = True) Then
            Dim yTemp() As Byte = oTech.GetPlayerTechKnowledgeMsg(yKnowledgeType)
            If yTemp Is Nothing = False Then
                Dim yMsg(7 + yTemp.Length) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eAddPlayerTechKnow).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = CByte(yKnowledgeType) : lPos += 1
                'If Me.bArchived = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
                yMsg(lPos) = yArchived
                lPos += 1
                yTemp.CopyTo(yMsg, lPos)

                oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewAgents Or AliasingRights.eViewTechDesigns)
            End If
        End If
	End Sub

	Public Function GetAddMsg() As Byte()
		Dim yTemp() As Byte = oTech.GetPlayerTechKnowledgeMsg(yKnowledgeType)
		If yTemp Is Nothing = False Then
			Dim yMsg(7 + yTemp.Length) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eAddPlayerTechKnow).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			yMsg(lPos) = CByte(yKnowledgeType) : lPos += 1
            'If Me.bArchived = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
            yMsg(lPos) = Me.yArchived
			lPos += 1
			yTemp.CopyTo(yMsg, lPos)

			Return yMsg
		End If
		Return Nothing
	End Function

	Public Shared Function DeleteAllPlayerTechKnowledge(ByVal lPlayerID As Int32) As Boolean
		Dim sSQL As String = ""
		Dim oComm As OleDb.OleDbCommand = Nothing
		Dim bResult As Boolean = False

		Try
			sSQL = "DELETE FROM tblPlayerTechKnowledge WHERE PlayerID = " & lPlayerID
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
