Public Class PlayerMineral
    'A relationship table between the player and the mineral... it is similar to mineral
    Public PlayerMineralID As Int32 = -1         'PK

    Public lMineralID As Int32              'quick lookup without having to reference oMineral
    Private mbDiscovered As Boolean = False   'indicates whether the player has discovered this mineral

    Public DiscoveryNumber As Int32 = 1     'this needs to be populated by the calling player object

    'The known mineral property values cannot exceed the mineral's actual property values
    Private moKnownMineralPropertyValues() As MineralPropertyValue
    Private mlKnownMineralPropertyValueUB As Int32 = -1

	Public oPlayer As Player

    'Public bArchived As Boolean = False
    Public yArchived As Byte = 0

    Private mbStringReady As Boolean = False
    Private mySendString() As Byte

    'Public Function GetMineralDetailsMsg() As Byte()
    '    Dim yMsg() As Byte
    '    Dim lPos As Int32

    '    ReDim yMsg(29 + (24 * moKnownMineralPropertyValues.Length))

    '    System.BitConverter.GetBytes(EpicaMessageCode.eMineralDetailsMsg).CopyTo(yMsg, 0)
    '    oMineral.GetGUIDAsString.CopyTo(yMsg, 2)
    '    lPos = 8

    '    If bDiscovered = True Then
    '        oMineral.MineralName.CopyTo(yMsg, lPos)
    '    Else : System.Text.ASCIIEncoding.ASCII.GetBytes("Unknown").CopyTo(yMsg, lPos)
    '    End If
    '    lPos += 20

    '    System.BitConverter.GetBytes(CShort(moKnownMineralPropertyValues.Length)).CopyTo(yMsg, lPos) : lPos += 2

    '    'Now, get our values...
    '    For lProp As Int32 = 0 To mlKnownMineralPropertyValueUB
    '        With moKnownMineralPropertyValues(lProp)
    '            System.BitConverter.GetBytes(.lPropertyID).CopyTo(yMsg, lPos) : lPos += 4

    '            If bDiscovered = True Then
    '                .oProperty.GetValueRangeName(.lValue).CopyTo(yMsg, lPos) : lPos += 20
    '            Else
    '                System.Text.ASCIIEncoding.ASCII.GetBytes("Unknown").CopyTo(yMsg, lPos) : lPos += 20
    '            End If
    '        End With
    '    Next lProp

    '    Return yMsg
    'End Function

    Public Sub SetPropertyValue(ByVal lMinPropValID As Int32, ByVal lPropID As Int32, ByVal lValue As Int32)
        mbStringReady = False
        For X As Int32 = 0 To mlKnownMineralPropertyValueUB
            If moKnownMineralPropertyValues(X).lPropertyID = lPropID Then
                moKnownMineralPropertyValues(X).lValue = lValue
                Return
            End If
        Next

        'Ok, didn't find it
        mlKnownMineralPropertyValueUB += 1
        ReDim Preserve moKnownMineralPropertyValues(mlKnownMineralPropertyValueUB)
        moKnownMineralPropertyValues(mlKnownMineralPropertyValueUB) = New MineralPropertyValue
        With moKnownMineralPropertyValues(mlKnownMineralPropertyValueUB)
            .MineralPropertyValueID = lMinPropValID
            .lPropertyID = lPropID
            .lParentID = Me.PlayerMineralID
            .lValue = lValue
            .oProperty = GetEpicaMineralProperty(lPropID)
        End With
    End Sub

    Public Function GetPropertyValue(ByVal lPropertyID As Int32) As Int32
        For X As Int32 = 0 To mlKnownMineralPropertyValueUB
            If moKnownMineralPropertyValues(X).lPropertyID = lPropertyID Then
                Return moKnownMineralPropertyValues(X).lValue
            End If
        Next
        Return 0
    End Function

	Public Function GetSaveObjectText(ByVal lPlayerID As Int32) As String
		Dim sSQL As String

		If PlayerMineralID = -1 Then
			SaveObject(lPlayerID)
			Return ""
		End If

		Try
			Dim oSB As New System.Text.StringBuilder
			'UPDATE
            sSQL = "UPDATE tblPlayerMineral SET PlayerID = " & lPlayerID & ", MineralID = " & lMineralID & _
              ", bArchived = " & yArchived & ", bDiscovered = "
			If mbDiscovered = True Then sSQL &= "1" Else sSQL &= "0"
			sSQL &= " WHERE PlayerMineralID = " & PlayerMineralID
			oSB.AppendLine(sSQL)

			'Now, save our values...
			For X As Int32 = 0 To mlKnownMineralPropertyValueUB
				moKnownMineralPropertyValues(X).lParentID = Me.PlayerMineralID
				oSB.AppendLine(moKnownMineralPropertyValues(X).GetSaveObjectText(1))
			Next X
			Return oSB.ToString
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & PlayerMineralID & " of type PlayerMineral. Reason: " & Err.Description)
		End Try
		Return ""
	End Function

    Public Function SaveObject(ByVal lPlayerID As Int32) As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If PlayerMineralID = -1 Then
                'INSERT
				sSQL = "INSERT INTO tblPlayerMineral (PlayerID, MineralID, bDiscovered, bArchived) VALUES (" & lPlayerID & _
				  ", " & lMineralID & ", " '& mbDiscovered & ")"
				If mbDiscovered = True Then sSQL &= "1, " Else sSQL &= "0, "
                sSQL &= yArchived & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblPlayerMineral SET PlayerID = " & lPlayerID & ", MineralID = " & lMineralID & _
                  ", bArchived = " & yArchived & ", bDiscovered = "
                If mbDiscovered = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= " WHERE PlayerMineralID = " & PlayerMineralID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If PlayerMineralID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PlayerMineralID) FROM tblPlayerMineral WHERE PlayerID = " & lPlayerID & " AND MineralID = " & lMineralID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    PlayerMineralID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            'Now, save our values...
            For X As Int32 = 0 To mlKnownMineralPropertyValueUB
                moKnownMineralPropertyValues(X).lParentID = Me.PlayerMineralID
                moKnownMineralPropertyValues(X).SaveObject(1)
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & PlayerMineralID & " of type PlayerMineral. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Private moMineral As Mineral = Nothing
    Public ReadOnly Property Mineral() As Mineral
        Get
            If moMineral Is Nothing Then
                moMineral = GetEpicaMineral(lMineralID)
            End If
            Return moMineral
        End Get
    End Property

    Public Property bDiscovered() As Boolean
        Get
            Return mbDiscovered
        End Get
        Set(ByVal value As Boolean)
            mbStringReady = False
            mbDiscovered = value
        End Set
    End Property

    Public Function GetObjAsString() As Byte()
        mbStringReady = False
        If mbStringReady = False Then
            If mbDiscovered = True Then
                Dim lPos As Int32

                ReDim mySendString(28 + ((oPlayer.mlMinPropertyUB + 1) * 45))

                Mineral.GetGUIDAsString.CopyTo(mySendString, 0)
                mySendString(6) = 1
                lPos = 7
                Mineral.MineralName.CopyTo(mySendString, lPos) : lPos += 20
                System.BitConverter.GetBytes(CShort(oPlayer.mlMinPropertyUB + 1)).CopyTo(mySendString, lPos) : lPos += 2

                For X As Int32 = 0 To oPlayer.mlMinPropertyUB

                    With oPlayer.moMinProperties(X)
                        System.BitConverter.GetBytes(.lPropertyID).CopyTo(mySendString, lPos) : lPos += 4
                        .oProperty.GetValueRangeName(Mineral.GetPropertyValue(.lPropertyID)).CopyTo(mySendString, lPos) : lPos += 20


                        If .Discovered > 0 Then
                            For Y As Int32 = 0 To mlKnownMineralPropertyValueUB
                                If moKnownMineralPropertyValues(Y).lPropertyID = .lPropertyID Then
                                    PlayerMineral.GetKnownValueName(Mineral.GetPropertyValue(.lPropertyID), moKnownMineralPropertyValues(Y).lValue).CopyTo(mySendString, lPos)
                                    Exit For
                                End If
                            Next
                        End If
                        lPos += 20

                        'mySendString(lPos) = CByte((Mineral.GetPropertyValue(.lPropertyID) / 255.0F) * 100)
                        mySendString(lPos) = GetClientDisplayedPropertyValue(Mineral.GetPropertyValue(.lPropertyID))
                        lPos += 1
                    End With
                Next X
            Else
                ReDim mySendString(26)
                Mineral.GetGUIDAsString.CopyTo(mySendString, 0)
                mySendString(6) = 0
                System.Text.ASCIIEncoding.ASCII.GetBytes("Unknown Mineral " & DiscoveryNumber).CopyTo(mySendString, 7)
            End If
        End If

        Return mySendString
    End Function

	Public Shared Function GetRandomValueWithinRange(ByVal yVisibleValue As Byte, ByVal yCurrVal As Byte) As Int32
		Dim lMin As Int32 = 1
		Dim lMax As Int32 = 255

		If yVisibleValue = 0 Then
			lMin = 1 : lMax = 22
		ElseIf yVisibleValue = 1 Then
			lMin = 23 : lMax = 45
		ElseIf yVisibleValue = 2 Then
			lMin = 46 : lMax = 68
		ElseIf yVisibleValue = 3 Then
			lMin = 69 : lMax = 91
		ElseIf yVisibleValue = 4 Then
			lMin = 92 : lMax = 115
		ElseIf yVisibleValue = 5 Then
			lMin = 116 : lMax = 139
		ElseIf yVisibleValue = 6 Then
			lMin = 140 : lMax = 163
		ElseIf yVisibleValue = 7 Then
			lMin = 164 : lMax = 186
		ElseIf yVisibleValue = 8 Then
			lMin = 187 : lMax = 209
		ElseIf yVisibleValue = 9 Then
			lMin = 210 : lMax = 232
		ElseIf yVisibleValue = 10 Then
			lMin = 233 : lMax = 255
		End If

		If yCurrVal < yVisibleValue Then
			'reduce lMax
			lMax -= ((lMax - lMin) \ 2)
		ElseIf yCurrVal > yVisibleValue Then
			'increase lmin
			lMin += ((lMax - lMin) \ 2)
		End If

		Return CInt(Rnd() * (lMax - lMin)) + lMin
	End Function
    Public Shared Function GetClientDisplayedPropertyValue(ByVal lValue As Int32) As Byte
        If lValue < 23 Then '11 Then
            Return 0
        ElseIf lValue < 46 Then ' 21 Then
            Return 1
        ElseIf lValue < 69 Then ' 31 Then
            Return 2
        ElseIf lValue < 92 Then ' 41 Then
            Return 3
        ElseIf lValue < 116 Then '51 Then
            Return 4
        ElseIf lValue < 140 Then '81 Then
            Return 5
        ElseIf lValue < 164 Then '101 Then
            Return 6
        ElseIf lValue < 187 Then '126 Then
            Return 7
        ElseIf lValue < 210 Then '151 Then
            Return 8
        ElseIf lValue < 233 Then '201 Then
            Return 9
        Else
            Return 10
        End If
    End Function

    Private Shared msKnown_To_Actual_Name() As String
    Public Shared Function GetKnownValueName(ByVal lActualValue As Int32, ByVal lKnownValue As Int32) As Byte()
        Dim yValue As Byte

        If lActualValue = 0 OrElse lActualValue = lKnownValue Then Return System.Text.ASCIIEncoding.ASCII.GetBytes("a full")

        yValue = CByte(Math.Max(Math.Min(9, (lKnownValue / lActualValue) * 10), 0))

        If msKnown_To_Actual_Name Is Nothing Then
            ReDim msKnown_To_Actual_Name(9)
            msKnown_To_Actual_Name(0) = "no"
            msKnown_To_Actual_Name(1) = "a sliver of"
            msKnown_To_Actual_Name(2) = "very little"
            msKnown_To_Actual_Name(3) = "a small"
            msKnown_To_Actual_Name(4) = "adequate"
            msKnown_To_Actual_Name(5) = "ample"
            msKnown_To_Actual_Name(6) = "a large"
            msKnown_To_Actual_Name(7) = "an expansive"
            msKnown_To_Actual_Name(8) = "a vast"
            msKnown_To_Actual_Name(9) = "an unlimited"
        End If

        Return System.Text.ASCIIEncoding.ASCII.GetBytes(msKnown_To_Actual_Name(yValue))
    End Function

    Public Function GetProductionCost() As ProductionCost
        Dim oProdCost As New ProductionCost

        With oProdCost
            If bDiscovered = False Then
                'Discovery research
                If Me.lMineralID = 41991 Then
                    .ColonistCost = 0
                    .CreditCost = 100000000
                    .EnlistedCost = 0
                    .OfficerCost = 0
                    .PointsRequired = 200000 * Me.Mineral.GetPropertyValue(eMinPropID.Complexity)
                Else
                    .ColonistCost = 0
                    .CreditCost = 10000
                    .EnlistedCost = 0
                    .OfficerCost = 0
                    .PointsRequired = 2000 * Me.Mineral.GetPropertyValue(eMinPropID.Complexity)
                    .PointsRequired \= Me.oPlayer.oSpecials.yDiscoverTime
                End If
            Else
                'Ok, study research
                .ColonistCost = 0
                .CreditCost = 10000
                .EnlistedCost = 0
                .OfficerCost = 0
                .PointsRequired = 600 * Me.Mineral.GetPropertyValue(eMinPropID.Complexity)
            End If

            .ItemCostUB = -1
            ReDim .ItemCosts(.ItemCostUB)
            .AddProductionCostItem(Me.lMineralID, ObjectType.eMineral, 200)
        End With

        Return oProdCost
    End Function

    ''' <summary>
    ''' Returns whether this mineral can be studied further
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ResearchComplete() As Boolean
        Dim X As Int32

        Dim bFurtherStudy As Boolean = False

        mbStringReady = False
 

        'Dim bAdjusted As Boolean = False

        If bDiscovered = False Then
            'ok, discovery research
            bDiscovered = True
            'bAdjusted = True
            Me.oPlayer.TestCustomTitlePermissions_DiscoverMineral()

            'ok, ensure that our discovered mineral properties are the same as the player's known mienral properties
            Dim lPropertyID As Int32
            For X = 0 To oPlayer.mlMinPropertyUB
                lPropertyID = oPlayer.moMinProperties(X).lPropertyID
                Dim lVal As Int32 = Me.Mineral.GetPropertyValue(lPropertyID)

                If lVal <> 0 Then
                    lVal = lVal \ 10I
                    lVal = CInt(Rnd() * lVal) + 1

                    SetPropertyValue(-1, lPropertyID, lVal)
                Else
                    SetPropertyValue(-1, lPropertyID, 0)
                End If
            Next X

            bFurtherStudy = True
        Else
            For lLoopCnt As Int32 = 0 To oPlayer.oSpecials.yStudyMineralEffect
                'Ok, increasing our known property values
                Dim lAmtDisburse As Int32 = 0
                For X = 0 To mlKnownMineralPropertyValueUB
                    Dim lVal As Int32 = moKnownMineralPropertyValues(X).lValue
                    Dim lMinVal As Int32 = Me.Mineral.GetPropertyValue(moKnownMineralPropertyValues(X).lPropertyID)
					Dim lDiff As Int32 = lMinVal - lVal

					If Me.lMineralID = 157 Then
						moKnownMineralPropertyValues(X).lValue = lMinVal
						Continue For
					End If

                    lDiff = Math.Min(lDiff, lMinVal \ 10I)

                    If lDiff > 1 Then
                        lMinVal = CInt(Rnd() * lDiff) '+ 1
                        lAmtDisburse += lMinVal
                        'If lMinVal <> 0 Then bAdjusted = True
                        moKnownMineralPropertyValues(X).lValue += lMinVal
                    Else
                        moKnownMineralPropertyValues(X).lValue = lMinVal
                    End If
                Next X

                'TODO: 20?
                If lAmtDisburse < 20 Then
                    lAmtDisburse = 20 - lAmtDisburse
                    For X = 0 To mlKnownMineralPropertyValueUB
                        Dim lVal As Int32 = moKnownMineralPropertyValues(X).lValue
                        Dim lMinVal As Int32 = Me.Mineral.GetPropertyValue(moKnownMineralPropertyValues(X).lPropertyID)
                        Dim lDiff As Int32 = lMinVal - lVal

                        If lDiff > lAmtDisburse Then
                            moKnownMineralPropertyValues(X).lValue += lAmtDisburse
                            Exit For
                        End If
                    Next X
                End If
            Next lLoopCnt

            'Now, determine if there is any room for improvement
            For X = 0 To mlKnownMineralPropertyValueUB
                Dim lVal As Int32 = moKnownMineralPropertyValues(X).lValue
                Dim lMinVal As Int32 = Me.Mineral.GetPropertyValue(moKnownMineralPropertyValues(X).lPropertyID)
                If lVal <> lMinVal Then
                    bFurtherStudy = True
                    Exit For
                End If
            Next X
        End If

        'If bAdjusted = True Then oPlayer.CheckForNewMineralPropertyTechs(lMineralID, 5)

        'Save our object
        Me.SaveObject(oPlayer.ObjectID)

        mbStringReady = False
		oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)

        Return bFurtherStudy
	End Function

	Public Function DeleteMe() As Boolean
		Dim sSQL As String = ""
		Dim oComm As OleDb.OleDbCommand = Nothing
		Dim bResult As Boolean = False
		Try
			sSQL = "DELETE FROM tblPlayerMineralPropertyValue WHERE PlayerMineralID = " & Me.PlayerMineralID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm.Dispose()
			oComm = Nothing

			sSQL = "DELETE FROM tblPlayerMineral WHERE PlayerMineralID = " & Me.PlayerMineralID
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
