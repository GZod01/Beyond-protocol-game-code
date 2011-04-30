Option Strict On

''' <summary>
''' Contains functionality for managing a players intel on other player's assets that are NOT technology designs. Use PlayerTechKnowledge for tech designs.
''' </summary>
''' <remarks></remarks>
Public Class PlayerItemIntel
    Public Enum PlayerItemIntelType As Byte
        eExistance = 0      'player knows of it
        eLocation = 1       'player knows where it is located
        eStatus = 2         'player knows the item's status
        eFullKnowledge = 4  'player knows all things
    End Enum
    Public lItemID As Int32                     'ID of the item
    Public iItemTypeID As Int16                 'TypeID of the item
	'Public lItemIdx As Int32                    'index in the global array for this item (some items may not use this)
    Public yIntelType As PlayerItemIntelType    'Intel Type indicating knowledge known
    Public lOtherPlayerID As Int32              'other player ID who this intel is about
    Public lPlayerID As Int32                   'player who knows this intel
    Public dtTimeStamp As DateTime              'real life date/time when knowledge was acquired
	Public lValue As Int32						'arbitrary value at time of timestamp

	Public LocX As Int32 = Int32.MinValue		'Location of this item when the intel was discovered
	Public LocZ As Int32 = Int32.MinValue		'Location of this item when the intel was discovered
	Public EnvirID As Int32 = -1				'Environment where this item was located when this intel was discovered
	Public EnvirTypeID As Int16 = -1			'Environment type 
	Public StatusValue As Int32 = 0				'status value (if the intel type is status or better)
	Public FullKnowledge() As Byte = Nothing	'string text of the intel report (if the intel type is fullknowledge or better)

    'Public bArchived As Boolean = False
    Public yArchived As Byte = 0

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Dim sPrepareFullKnowledge As String = ""
        If FullKnowledge Is Nothing = False Then sPrepareFullKnowledge = Base64.EncodeString(BytesToString(FullKnowledge))

        Try
            'attempt to update first
            sSQL = "UPDATE tblPlayerItemIntel SET IntelType = " & CByte(yIntelType) & ", IntelTimeStamp = '" & dtTimeStamp.ToString & _
              "', IntelValue = " & lValue & ", LocX = " & LocX & ", LocZ = " & LocZ & ", EnvirID = " & EnvirID & ", EnvirTypeID = " & _
              EnvirTypeID & ", StatusValue = " & StatusValue & ", bArchived = " & yArchived
            If sPrepareFullKnowledge Is Nothing = False Then sSQL &= ", FullKnowledge = '" & MakeDBStr(sPrepareFullKnowledge) & "'"
			sSQL &= " WHERE ItemID = " & lItemID & " AND ItemTypeID = " & iItemTypeID & " AND PlayerID = " & _
			  lPlayerID & " AND OtherPlayerID = " & lOtherPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                oComm.Dispose()
                oComm = Nothing

				sSQL = "INSERT INTO tblPlayerItemIntel (ItemID, ItemTypeID, PlayerID, OtherPlayerID, IntelType, IntelTimeStamp, " & _
				  "IntelValue, LocX, LocZ, EnvirID, EnvirTypeID, StatusValue, FullKnowledge, bArchived) VALUES (" & _
				  lItemID & ", " & iItemTypeID & ", " & lPlayerID & ", " & lOtherPlayerID & ", " & CByte(yIntelType) & ", '" & dtTimeStamp.ToString & _
				  "', " & lValue & ", " & LocX & ", " & LocZ & ", " & EnvirID & ", " & EnvirTypeID & ", " & StatusValue & ", '"
                If sPrepareFullKnowledge Is Nothing = False Then sSQL &= MakeDBStr(sPrepareFullKnowledge)
                sSQL &= "', " & yArchived & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
                oComm.Dispose()
                oComm = Nothing
            End If

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save PlayerItemIntel. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function GetObjAsString() As Byte()
        Dim yResp(24) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(lItemID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(ObjectType.ePlayerItemIntel).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(iItemTypeID).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOtherPlayerID).CopyTo(yResp, lPos) : lPos += 4
        yResp(lPos) = yIntelType : lPos += 1

        Dim lSeconds As Int32 = CInt(Now.Subtract(dtTimeStamp).TotalSeconds)
        System.BitConverter.GetBytes(lSeconds).CopyTo(yResp, lPos) : lPos += 4
		System.BitConverter.GetBytes(lValue).CopyTo(yResp, lPos) : lPos += 4
		Return yResp
	End Function

	Public Function GetDetails() As Byte()

		Dim lDataLen As Int32 = 12
		If yIntelType >= PlayerItemIntelType.eLocation Then
			lDataLen += 14
			If yIntelType >= PlayerItemIntelType.eStatus Then
				lDataLen += 4
				If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
					lDataLen += FullKnowledge.Length + 4
				End If
			End If
		End If

		Dim yResp(lDataLen - 1) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eGetItemIntelDetail).CopyTo(yResp, lPos) : lPos += 2
		System.BitConverter.GetBytes(lItemID).CopyTo(yResp, lPos) : lPos += 4
		System.BitConverter.GetBytes(iItemTypeID).CopyTo(yResp, lPos) : lPos += 2
		System.BitConverter.GetBytes(lOtherPlayerID).CopyTo(yResp, lPos) : lPos += 4

		If yIntelType >= PlayerItemIntelType.eLocation Then
			System.BitConverter.GetBytes(LocX).CopyTo(yResp, lPos) : lPos += 4
			System.BitConverter.GetBytes(LocZ).CopyTo(yResp, lPos) : lPos += 4
			System.BitConverter.GetBytes(EnvirID).CopyTo(yResp, lPos) : lPos += 4
			System.BitConverter.GetBytes(EnvirTypeID).CopyTo(yResp, lPos) : lPos += 2
			If yIntelType >= PlayerItemIntelType.eStatus Then
				System.BitConverter.GetBytes(StatusValue).CopyTo(yResp, lPos) : lPos += 4
				If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
					System.BitConverter.GetBytes(FullKnowledge.Length).CopyTo(yResp, lPos) : lPos += 4
					FullKnowledge.CopyTo(yResp, lPos) : lPos += FullKnowledge.Length
				End If
			End If
		End If
		Return yResp
	End Function

	''' <summary>
	''' Called only when the player item intel is made... 
	''' </summary>
	''' <remarks></remarks>
	Public Sub PopulateData()
		'if only existance, than no need to go any further
		If yIntelType = PlayerItemIntelType.eExistance Then Return

		'ok, let's get our data...
		Select Case iItemTypeID
			Case ObjectType.eAgent
				Dim oAgent As Agent = GetEpicaAgent(lItemID)
				If oAgent Is Nothing = False Then
					'ok, we have at least location...
					EnvirID = oAgent.lTargetID
					EnvirTypeID = oAgent.iTargetTypeID
					LocX = Int32.MinValue
					LocZ = Int32.MinValue

					'Now, do we have status?
					If yIntelType >= PlayerItemIntelType.eStatus Then
						StatusValue = oAgent.lAgentStatus
					Else : StatusValue = 0
					End If

                    'do we have full knowledge?
                    If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
                        FullKnowledge = oAgent.GetObjAsString()
                    Else
                        FullKnowledge = Nothing
                    End If
				End If
			Case ObjectType.eColony
				Dim oColony As Colony = GetEpicaColony(lItemID)
				If oColony Is Nothing = False Then
					LocX = Int32.MinValue
					LocZ = Int32.MinValue
					Dim oParent As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
					If oParent Is Nothing = False Then
						If oParent.ObjTypeID = ObjectType.eFacility Then
							With CType(oParent, Facility)
								LocX = .LocX
								LocZ = .LocZ
							End With
							oParent = CType(CType(oParent, Facility).ParentObject, Epica_GUID)
							If oParent Is Nothing = False Then
								If oParent.ObjTypeID <> ObjectType.ePlanet AndAlso oParent.ObjTypeID <> ObjectType.eSolarSystem Then
									yIntelType = PlayerItemIntelType.eExistance
									Return
								End If
							End If
						ElseIf oParent.ObjTypeID <> ObjectType.ePlanet AndAlso oParent.ObjTypeID <> ObjectType.eSolarSystem Then
							yIntelType = PlayerItemIntelType.eExistance
							Return
						End If
					Else
						yIntelType = PlayerItemIntelType.eExistance
						Return
					End If
					EnvirID = oParent.ObjectID
					EnvirTypeID = oParent.ObjTypeID

					'Now, check for status
					If yIntelType >= PlayerItemIntelType.eStatus Then
						StatusValue = oColony.Population
					End If

					If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
						'Jobs (by makeup), housing, enlisted, officers, power gen, power need, morale
						Dim oSB As New System.Text.StringBuilder
						With oColony
							oSB.AppendLine("Workforce Jobs: " & .FactoryJobs.ToString("#,##0"))
							oSB.AppendLine("Scientist Jobs: " & .ScientistJobs.ToString("#,##0"))
							oSB.AppendLine("Other Jobs: " & .OtherJobs.ToString("#,##0"))
							oSB.AppendLine("Powered Housing: " & .PoweredHousing.ToString("#,##0"))
							oSB.AppendLine("Unpowered Housing: " & .UnpoweredHousing.ToString("#,##0"))
							oSB.AppendLine("Total Housing: " & (.PoweredHousing + .UnpoweredHousing).ToString("#,##0"))
							oSB.AppendLine("Enlisted: " & .ColonyEnlisted.ToString("#,##0"))
							oSB.AppendLine("Officers: " & .ColonyOfficers.ToString("#,##0"))
							oSB.AppendLine("Power Generated: " & .PowerGeneration.ToString("#,##0"))
							oSB.AppendLine("Power Needed: " & .PowerConsumption.ToString("#,##0"))
							oSB.AppendLine("Morale: " & (.MoraleMultiplier * 100.0F).ToString("##0"))
						End With
						FullKnowledge = System.Text.ASCIIEncoding.ASCII.GetBytes(oSB.ToString)
					Else : FullKnowledge = Nothing
					End If
				End If
			Case ObjectType.eFacility
				Dim oFacility As Facility = GetEpicaFacility(lItemID)
				If oFacility Is Nothing = False Then
					'ok, we have at least location...
					Dim oParent As Epica_GUID = CType(oFacility.ParentObject, Epica_GUID)
					If oParent Is Nothing Then
						yIntelType = PlayerItemIntelType.eExistance
						Return
					End If
					If oParent.ObjTypeID = ObjectType.eFacility Then
						With CType(oParent, Facility)
							LocX = .LocX
							LocZ = .LocZ
						End With
						oParent = CType(CType(oParent, Facility).ParentObject, Epica_GUID)
						If oParent Is Nothing Then
							yIntelType = PlayerItemIntelType.eExistance
							Return
						End If
						If oParent.ObjTypeID <> ObjectType.ePlanet AndAlso oParent.ObjTypeID <> ObjectType.eSolarSystem Then
							yIntelType = PlayerItemIntelType.eExistance
							Return
						End If
						EnvirID = oParent.ObjectID
						EnvirTypeID = oParent.ObjTypeID
					ElseIf oParent.ObjTypeID <> ObjectType.ePlanet AndAlso oParent.ObjTypeID <> ObjectType.eSolarSystem Then
						yIntelType = PlayerItemIntelType.eExistance
						Return
					Else
						EnvirID = oParent.ObjectID
						EnvirTypeID = oParent.ObjTypeID
						LocX = oFacility.LocX
						LocZ = oFacility.LocZ
					End If

					'Now, do we have status?
					If yIntelType >= PlayerItemIntelType.eStatus Then
						StatusValue = CInt(oFacility.Active)
					Else : StatusValue = 0
					End If

                    'do we have full knowledge?
                    If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
                        FullKnowledge = oFacility.GetObjAsString()
                    Else
                        FullKnowledge = Nothing
                    End If
				End If

			Case ObjectType.eMineral
				FullKnowledge = Nothing
                If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
                    Dim oPlayer As Player = GetEpicaPlayer(lOtherPlayerID)
                    If oPlayer Is Nothing = False Then
                        For X As Int32 = 0 To oPlayer.lPlayerMineralUB
                            If oPlayer.oPlayerMinerals(X) Is Nothing = False Then
                                If oPlayer.oPlayerMinerals(X).lMineralID = lItemID Then
                                    FullKnowledge = oPlayer.oPlayerMinerals(X).GetObjAsString
                                End If
                            End If
                        Next X
                    End If
                End If

			Case ObjectType.ePlayer	'playerrel
				Dim oOtherPlayer As Player = GetEpicaPlayer(lOtherPlayerID)
				If oOtherPlayer Is Nothing = False Then
					Dim oRel As PlayerRel = oOtherPlayer.GetPlayerRel(lItemID)
					If oRel Is Nothing = False Then
						StatusValue = oRel.WithThisScore
					End If
					oRel = Nothing
				End If
				oOtherPlayer = Nothing
			Case ObjectType.eTrade
				'amounts...
				Dim oOtherPlayer As Player = GetEpicaPlayer(lOtherPlayerID)
				Dim blExports As Int64 = 0
				Dim blImports As Int64 = 0

				Dim sFinal As String = "Trade Revenue"

				If oOtherPlayer Is Nothing = False Then
					sFinal &= " for " & oOtherPlayer.sPlayerNameProper
					blImports = oOtherPlayer.oBudget.GetTradeIncome(lItemID)
				End If
				oOtherPlayer = GetEpicaPlayer(lItemID)
				If oOtherPlayer Is Nothing = False Then
					sFinal &= " with " & oOtherPlayer.sPlayerNameProper
					blExports = oOtherPlayer.oBudget.GetTradeIncome(lOtherPlayerID)
				End If
				sFinal &= vbCrLf & "Imports: " & blImports.ToString("#,##0")
				sFinal &= vbCrLf & "Exports: " & blExports.ToString("#,##0")
				FullKnowledge = System.Text.ASCIIEncoding.ASCII.GetBytes(sFinal)
			Case ObjectType.eUnit
				Dim oUnit As Unit = GetEpicaUnit(lItemID)
				If oUnit Is Nothing = False Then
					'ok, we have at least location...
					Dim oParent As Epica_GUID = CType(oUnit.ParentObject, Epica_GUID)
					If oParent Is Nothing Then
						yIntelType = PlayerItemIntelType.eExistance
						Return
					End If
					If oParent.ObjTypeID = ObjectType.eFacility Then
						With CType(oParent, Facility)
							LocX = .LocX
							LocZ = .LocZ
						End With
						oParent = CType(CType(oParent, Facility).ParentObject, Epica_GUID)
						If oParent Is Nothing Then
							yIntelType = PlayerItemIntelType.eExistance
							Return
						End If
						If oParent.ObjTypeID <> ObjectType.ePlanet AndAlso oParent.ObjTypeID <> ObjectType.eSolarSystem Then
							yIntelType = PlayerItemIntelType.eExistance
							Return
						End If
						EnvirID = oParent.ObjectID
						EnvirTypeID = oParent.ObjTypeID
					ElseIf oParent.ObjTypeID <> ObjectType.ePlanet AndAlso oParent.ObjTypeID <> ObjectType.eSolarSystem Then
						yIntelType = PlayerItemIntelType.eExistance
						Return
					Else
						EnvirID = oParent.ObjectID
						EnvirTypeID = oParent.ObjTypeID
						LocX = oUnit.LocX
						LocZ = oUnit.LocZ
					End If

					'Now, do we have status?
					If yIntelType >= PlayerItemIntelType.eStatus Then
						StatusValue = oUnit.ExpLevel
					Else : StatusValue = 0
					End If

					'do we have full knowledge?
                    If yIntelType >= PlayerItemIntelType.eFullKnowledge Then
                        FullKnowledge = oUnit.GetObjAsString()
                    Else
                        FullKnowledge = Nothing
                    End If
				End If
			Case Else
		End Select

	End Sub

	Public Function CloneForPlayer(ByRef oNewPlayer As Player) As PlayerItemIntel
		Dim oPII As PlayerItemIntel = oNewPlayer.GetPlayerItemIntel(lItemID, iItemTypeID, lOtherPlayerID)
		Dim bCreated As Boolean = False
		If oPII Is Nothing Then
			oPII = New PlayerItemIntel()
			bCreated = True
		End If
		If oPII Is Nothing = False Then
			'ok, now, overwrite the remainder of our data
			With oPII
				.lItemID = lItemID
				.iItemTypeID = iItemTypeID
				.lOtherPlayerID = lOtherPlayerID
				.lPlayerID = oNewPlayer.ObjectID
				If dtTimeStamp > .dtTimeStamp Then .dtTimeStamp = dtTimeStamp
				.lValue = lValue

				.yIntelType = CType(Math.Max(yIntelType, .yIntelType), PlayerItemIntelType)
				.LocX = LocX
				.LocZ = LocZ
				.EnvirID = EnvirID
				.EnvirTypeID = EnvirTypeID

				If yIntelType >= PlayerItemIntelType.eStatus Then .StatusValue = StatusValue
				If yIntelType >= PlayerItemIntelType.eFullKnowledge AndAlso FullKnowledge Is Nothing = False Then
					ReDim .FullKnowledge(FullKnowledge.GetUpperBound(0))
					Array.Copy(FullKnowledge, 0, .FullKnowledge, 0, .FullKnowledge.Length)
				End If

				If bCreated = True Then
					oNewPlayer.mlItemIntelUB += 1
					ReDim Preserve oNewPlayer.moItemIntel(oNewPlayer.mlItemIntelUB)
					oNewPlayer.moItemIntel(oNewPlayer.mlItemIntelUB) = oPII
				End If

				.SaveObject()
			End With
		End If
		Return oPII
	End Function

	Public Shared Function DeleteAllPlayerItemIntel(ByVal lPlayerID As Int32) As Boolean
		Dim sSQL As String = ""
		Dim oComm As OleDb.OleDbCommand = Nothing
		Dim bResult As Boolean = False

		Try
			sSQL = "DELETE FROM tblPlayerItemIntel WHERE PlayerID = " & lPlayerID
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