Public Class PlayerMineralProperty

    Public lPlayerID As Int32
    Public lPropertyID As Int32

    Private moPlayer As Player
    Public ReadOnly Property oPlayer() As Player
        Get
            If moPlayer Is Nothing Then moPlayer = GetEpicaPlayer(lPlayerID)
            Return moPlayer
        End Get
    End Property

    Private moProperty As MineralProperty
    Public ReadOnly Property oProperty() As MineralProperty
        Get
            If moProperty Is Nothing Then moProperty = GetEpicaMineralProperty(lPropertyID)
            Return moProperty
        End Get
    End Property

    Private myDiscovered As Byte = 0
    Public Property Discovered() As Byte
        Get
            Return myDiscovered
        End Get
        Set(ByVal value As Byte)
            'is the new value higher than discovered?
            If value > myDiscovered Then
                'Yes, is our current value 0?
                If myDiscovered = 0 Then

                    If oPlayer Is Nothing Then Return
                    'Yes, the new value will be 1 or 2 which means we need to update all known minerals
                    '  to have known property values for this property
                    For X As Int32 = 0 To moPlayer.lPlayerMineralUB
                        'Make sure that the property KNOWN value is 0, if it is not, then it exists and shouldn't be touched
                        If moPlayer.oPlayerMinerals(X).GetPropertyValue(lPropertyID) = 0 Then
                            'Ok, get our Mineral's ACTUAL property value
                            Dim lVal As Int32 = moPlayer.oPlayerMinerals(X).Mineral.GetPropertyValue(lPropertyID)
                            'Is the Value not 0?
                            If lVal <> 0 Then
                                'Ok, let's set its initial KNOWN value
                                lVal = lVal \ 10I
                                lVal = CInt(Rnd() * lVal) + 1
                            End If
                            'And call SetPropertyValue with the new value
                            moPlayer.oPlayerMinerals(X).SetPropertyValue(-1, lPropertyID, lVal)

                            'We'll update the player on the new mineral values...
							moPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(moPlayer.oPlayerMinerals(X), GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns Or AliasingRights.eViewMining)
                        End If

                    Next X
                End If

                'You can ONLY go up in discovery level
                myDiscovered = value
            End If
        End Set
    End Property

    Public Function GetObjAsString() As Byte()
        Dim mySendString(56) As Byte

        oProperty.GetGUIDAsString.CopyTo(mySendString, 0)
        oProperty.MineralPropertyName.CopyTo(mySendString, 6)
        mySendString(56) = myDiscovered

        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand = Nothing

        Try

            'Ok, first, try and update
            sSQL = "UPDATE tblPlayerMineralProperty SET lDiscovered = " & myDiscovered & " WHERE PlayerID = " & Me.lPlayerID & _
               " AND MineralPropertyID = " & Me.lPropertyID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            Dim lRecs As Int32 = oComm.ExecuteNonQuery()

            If lRecs = 0 Then
                'ok, try an insert
                oComm.Dispose()
                oComm = Nothing
                sSQL = "INSERT INTO tblPlayerMineralProperty (PlayerID, MineralPropertyID, lDiscovered) VALUES (" & _
                  Me.lPlayerID & ", " & Me.lPropertyID & ", " & myDiscovered & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
            End If
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save PlayerMineralProperty " & Me.lPlayerID & ", " & Me.lPropertyID & ". Reason: " & Err.Description)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

	End Function

	Public Function GetSaveObjectText() As String
		Return "INSERT INTO tblPlayerMineralProperty (PlayerID, MineralPropertyID, lDiscovered) VALUES (" & _
	Me.lPlayerID & ", " & Me.lPropertyID & ", " & myDiscovered & ")"
	End Function

    Public Sub SetDiscoveredLevel_NoEvents(ByVal yVal As Byte)
        myDiscovered = yVal
    End Sub

End Class
