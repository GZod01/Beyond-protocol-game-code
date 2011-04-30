'Option Strict On

'Public Class GuildRel
'	'Guild Rel can be another player or another guild
'	Public lEntityID As Int32
'	Public iEntityTypeID As Int16

'	Public yRelTowardsUs As Byte		'their rel towards us
'	Public yRelTowardsThem As Byte		'our rel towards them

'	Public lWhoMadeFirstContact As Int32 = -1
'	Public lWhoFirstContactWasMadeWith As Int32 = -1
'	Public dtWhenFirstContactMade As Date = Date.MinValue		'in GMT
'	Public lLocationIDOfFC As Int32 = -1
'	Public iLocationTypeIDOfFC As Int16 = -1
'	Public lLocXOfFC As Int32 = 0
'	Public lLocZOfFC As Int32 = 0

'	Public yNote() As Byte = Nothing

'	''Guild-Specific
'	'Public lKnownMemberCount As Int32
'	'Public lLeaderID As Int32 = -1

'	''Player-Specific
'	'Public lPlayerGuildID As Int32 = -1
'	'Public yPlayerTitle As Byte = 0
'	'Public bIsMale As Boolean = False

'	Public Function GetObjAsSmallString() As Byte()
'		Dim yMsg(11) As Byte
'		Dim lPos As Int32 = 0

'		Dim lIcon As Int32 = 0
'		If iEntityTypeID = ObjectType.ePlayer Then
'			Dim oPlayer As Player = GetEpicaPlayer(lEntityID)
'			If oPlayer Is Nothing = False Then lIcon = oPlayer.lPlayerIcon
'			oPlayer = Nothing
'		ElseIf iEntityTypeID = ObjectType.eGuild Then
'			Dim oGuild As Guild = GetEpicaGuild(lEntityID)
'			If oGuild Is Nothing = False Then lIcon = oGuild.lIcon
'			oGuild = Nothing
'		End If

'		System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, lPos) : lPos += 2
'		yMsg(lPos) = yRelTowardsUs : lPos += 1
'		yMsg(lPos) = yRelTowardsThem : lPos += 1
'		System.BitConverter.GetBytes(lIcon).CopyTo(yMsg, lPos) : lPos += 4
'		Return yMsg
'	End Function

'	Public Function SaveObject(ByVal lGuildID As Int32) As Boolean
'		Dim bResult As Boolean = False
'		Dim sSQL As String
'		Dim oComm As OleDb.OleDbCommand

'		Try

'			Dim lDtOfFC As Int32 = 0
'			If dtWhenFirstContactMade <> Date.MinValue Then lDtOfFC = GetDateAsNumber(dtWhenFirstContactMade)
'			sSQL = "UPDATE tblGuildRel SET RelTowardsUs = " & yRelTowardsUs & ", RelTowardsThem = " & yRelTowardsThem & _
'			 ", WhoMadeFirstContact = " & lWhoMadeFirstContact & ", WhoFirstContactWasMadeWith = " & lWhoFirstContactWasMadeWith & _
'			 ", LocationIDOfFC = " & lLocationIDOfFC & ", LocTypeIDOfFC = " & iLocationTypeIDOfFC & ", LocXOfFC = " & lLocXOfFC & _
'			 ", LocZOfFC = " & lLocZOfFC & ", DateOfFC = " & lDtOfFC & ", CustomNotes = '"
'			If yNote Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yNote))
'			sSQL &= "' WHERE GuildID = " & lGuildID & " AND EntityID = " & lEntityID & " AND EntityTypeID = " & iEntityTypeID

'			oComm = New OleDb.OleDbCommand(sSQL, goCN)
'			If oComm.ExecuteNonQuery() = 0 Then
'				oComm = Nothing

'				sSQL = "INSERT INTO tblGuildRel (EntityID, EntityTypeID, GuildID, RelTowardsUs, RelTowardsThem, WhoMadeFirstContact, " & _
'				 "WhoFirstContactWasMadeWith, LocationIDOfFC, LocTypeIDOfFC, LocXOfFC, LocZOfFC, DateOfFC, CustomNotes) VALUES (" & _
'				 lEntityID & ", " & iEntityTypeID & ", " & lGuildID & ", " & yRelTowardsUs & ", " & yRelTowardsThem & ", " & _
'				 lWhoMadeFirstContact & ", " & lWhoFirstContactWasMadeWith & ", " & lLocationIDOfFC & ", " & iLocationTypeIDOfFC & _
'				 ", " & lLocXOfFC & ", " & lLocZOfFC & ", " & lDtOfFC & ", '"
'				If yNote Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yNote))
'				sSQL &= "')"

'				oComm = New OleDb.OleDbCommand(sSQL, goCN)
'				If oComm.ExecuteNonQuery() = 0 Then
'					Err.Raise(-1, "SaveObject", "No records affected when saving object!")
'				End If
'			End If

'			bResult = True
'		Catch
'			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object GuildRel. Reason: " & Err.Description)
'		Finally
'			oComm = Nothing
'		End Try
'		Return bResult
'	End Function

'	Public Function GetFullObjAsString() As Byte()
'		Dim lNoteLen As Int32 = 0
'		If yNote Is Nothing = False Then lNoteLen = yNote.Length


'		Dim lIcon As Int32 = 0
'		Dim lAdditionalDataLen As Int32 = 0
'		Dim yAddData() As Byte = Nothing
'		If iEntityTypeID = ObjectType.ePlayer Then
'			Dim oPlayer As Player = GetEpicaPlayer(lEntityID)
'			If oPlayer Is Nothing = False Then
'				lIcon = oPlayer.lPlayerIcon
'				lAdditionalDataLen = 6
'				ReDim yAddData(lAdditionalDataLen - 1)
'				System.BitConverter.GetBytes(oPlayer.lGuildID).CopyTo(yAddData, 0)
'				yAddData(4) = oPlayer.yPlayerTitle
'				yAddData(5) = oPlayer.yGender
'			End If
'			oPlayer = Nothing

'		ElseIf iEntityTypeID = ObjectType.eGuild Then
'			Dim oGuild As Guild = GetEpicaGuild(lEntityID)
'			If oGuild Is Nothing = False Then
'				lIcon = oGuild.lIcon
'				lAdditionalDataLen = 8
'				ReDim yAddData(lAdditionalDataLen - 1)
'				System.BitConverter.GetBytes(oGuild.GetApprovedMemberCount).CopyTo(yAddData, 0)
'				System.BitConverter.GetBytes(-1I).CopyTo(yAddData, 4)
'			End If
'			oGuild = Nothing

'		End If

'		Dim yMsg(41 + lNoteLen + lAdditionalDataLen) As Byte
'		Dim lPos As Int32 = 0


'		System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, lPos) : lPos += 2
'		yMsg(lPos) = yRelTowardsUs : lPos += 1
'		yMsg(lPos) = yRelTowardsThem : lPos += 1
'		System.BitConverter.GetBytes(lIcon).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(lWhoMadeFirstContact).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(lWhoFirstContactWasMadeWith).CopyTo(yMsg, lPos) : lPos += 4
'		Dim lVal As Int32 = 0
'		If dtWhenFirstContactMade <> Date.MinValue Then lVal = GetDateAsNumber(dtWhenFirstContactMade)
'		System.BitConverter.GetBytes(lVal).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(lLocationIDOfFC).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(iLocationTypeIDOfFC).CopyTo(yMsg, lPos) : lPos += 2
'		System.BitConverter.GetBytes(lLocXOfFC).CopyTo(yMsg, lPos) : lPos += 4
'		System.BitConverter.GetBytes(lLocZOfFC).CopyTo(yMsg, lPos) : lPos += 4

'		System.BitConverter.GetBytes(lNoteLen).CopyTo(yMsg, lPos) : lPos += 4
'		If yNote Is Nothing = False Then
'			yNote.CopyTo(yMsg, lPos) : lPos += yNote.Length
'		End If

'		If yAddData Is Nothing = False Then yAddData.CopyTo(yMsg, lPos)

'		Return yMsg
'	End Function

'End Class
