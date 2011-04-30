''' <summary>
''' Indicates a group of actual technology components that have been built
''' </summary>
''' <remarks></remarks>
Public Class ComponentCache
    Inherits Epica_GUID

    Public ComponentID As Int32
    Public ComponentTypeID As Int16
    Public ComponentOwnerID As Int32    'where to look for this component

    Public ParentObject As Object       'where the cache exists

    Public LocX As Int32
    Public LocZ As Int32

    Private mlQuantity As Int32

    Private mySendString() As Byte

	Public bNeedsSaved As Boolean = False

	Public yCacheTypeID As Byte

    'Helper function
    Private moComponent As Epica_Tech = Nothing
    Public ReadOnly Property GetComponent() As Epica_Tech
        Get
            If moComponent Is Nothing Then
                Dim oPlayer As Player = GetEpicaPlayer(ComponentOwnerID)
                If oPlayer Is Nothing Then oPlayer = goInitialPlayer
                If oPlayer Is Nothing = False Then
                    moComponent = oPlayer.GetTech(ComponentID, ComponentTypeID)
                End If
            End If
            Return moComponent
        End Get
    End Property

    'Let's try this...
    Public Property Quantity() As Int32
        Get
            Return mlQuantity
        End Get
        Set(ByVal value As Int32)
            If value <> mlQuantity Then
                bNeedsSaved = True
                mlQuantity = value
                Me.DataChanged()

                'Now... if the quantity = 0 then we need to remove this cache from the DB
                If mlQuantity < 1 Then
                    mlQuantity = 0

                    'TODO: This sucks... but I have to look through the array for me...
                    For X As Int32 = 0 To glComponentCacheUB
                        If glComponentCacheIdx(X) = Me.ObjectID Then
                            glComponentCacheIdx(X) = -1
                            Exit For
                        End If
                    Next X

                    'Now... delete me
                    DeleteMe()

                    'Now... if this cache is in a planet or system...
                    Dim iTemp As Int16 = CType(Me.ParentObject, Epica_GUID).ObjTypeID
                    If iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
                        Dim yData(8) As Byte
						System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yData, 0)
                        GetGUIDAsString.CopyTo(yData, 2)
                        yData(8) = RemovalType.eObjectDestroyed

                        If iTemp = ObjectType.ePlanet Then
                            CType(ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
                        ElseIf iTemp = ObjectType.eSolarSystem Then
                            CType(ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
                        End If
                    Else
                        'most likely a facility or unit...
                        If iTemp = ObjectType.eFacility OrElse iTemp = ObjectType.eUnit Then
                            'Ok, remove me from their cargo
                            With CType(Me.ParentObject, Epica_Entity)
                                For X As Int32 = 0 To .lCargoUB
                                    If .lCargoIdx(X) = Me.ObjectID AndAlso .oCargoContents(X).ObjTypeID = Me.ObjTypeID Then
                                        .lCargoIdx(X) = -1
                                        .oCargoContents(X) = Nothing
                                        Exit For
                                    End If
                                Next X
                            End With
                        ElseIf iTemp = ObjectType.eColony Then
                            Try
                                With CType(Me.ParentObject, Colony)
                                    For X As Int32 = 0 To .mlComponentCacheUB
                                        If .mlComponentCacheID(X) = Me.ObjectID AndAlso glComponentCacheIdx(.mlComponentCacheIdx(X)) = Me.ObjectID Then
                                            .mlComponentCacheIdx(X) = -1
                                            Exit For
                                        End If
                                    Next X
                                End With
                            Catch
                            End Try
                        End If
                    End If
                End If
            End If
        End Set
    End Property

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
			ReDim mySendString(34)		'0 to 34 = 35 bytes

            Dim lPos As Int32 = 0

            GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            CType(ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            System.BitConverter.GetBytes(LocX).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Quantity).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ComponentID).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(ComponentTypeID).CopyTo(mySendString, lPos) : lPos += 2
			System.BitConverter.GetBytes(ComponentOwnerID).CopyTo(mySendString, lPos) : lPos += 4

			If yCacheTypeID = 0 OrElse yCacheTypeID < 8 Then		'8 is armordebris
				'ok, no component cache is mineable...
				Select Case ComponentTypeID
					Case ObjectType.eArmorTech
						yCacheTypeID = yCacheTypeID Or MineralCacheType.eArmorDebris
					Case ObjectType.eEngineTech
						yCacheTypeID = yCacheTypeID Or MineralCacheType.eEngineDebris
					Case ObjectType.eRadarTech
						yCacheTypeID = yCacheTypeID Or MineralCacheType.eRadarDebris
					Case ObjectType.eShieldTech
						yCacheTypeID = yCacheTypeID Or MineralCacheType.eShieldDebris
					Case ObjectType.eWeaponTech
						yCacheTypeID = yCacheTypeID Or MineralCacheType.eWeaponDebris
				End Select
			End If

			mySendString(lPos) = yCacheTypeID : lPos += 1

            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        'If the parent object is nothing, then this cache no longer exists
        If ParentObject Is Nothing Then Return True

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblComponentCache (ComponentID, ComponentTypeID, ComponentOwnerID, ParentID, " & _
                  "ParentTypeID, LocX, LocZ, Quantity) VALUES (" & ComponentID & ", " & ComponentTypeID & ", " & _
                  ComponentOwnerID & ", " & CType(ParentObject, Epica_GUID).ObjectID & ", " & _
                  CType(ParentObject, Epica_GUID).ObjTypeID & ", " & LocX & ", " & LocZ & ", " & Quantity & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblComponentCache SET ComponentID = " & ComponentID & ", ComponentTypeID = " & _
                  ComponentTypeID & ", ComponentOwnerID = " & ComponentOwnerID & ", ParentID = " & _
                  CType(ParentObject, Epica_GUID).ObjectID & ", ParentTypeID = " & _
                  CType(ParentObject, Epica_GUID).ObjTypeID & ", LocX = " & LocX & ", LocZ = " & LocZ & _
                  ", Quantity = " & Quantity & " WHERE ComponentCacheID = " & Me.ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(ComponentCacheID) FROM tblComponentCache WHERE ComponentID = " & ComponentID & " AND ComponentTypeID = " & ComponentTypeID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            bResult = True
            bNeedsSaved = False
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

	Public Function GetSaveObjectText() As String
		Dim bResult As Boolean = False
		Dim sSQL As String  

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		'If the parent object is nothing, then this cache no longer exists
		If ParentObject Is Nothing Then Return ""
		If ObjectID = -1 Then
			If SaveObject() = False Then Return ""
		End If

		Try
			If ObjectID = -1 Then
				'INSERT
				sSQL = "INSERT INTO tblComponentCache (ComponentID, ComponentTypeID, ComponentOwnerID, ParentID, " & _
				  "ParentTypeID, LocX, LocZ, Quantity) VALUES (" & ComponentID & ", " & ComponentTypeID & ", " & _
				  ComponentOwnerID & ", " & CType(ParentObject, Epica_GUID).ObjectID & ", " & _
				  CType(ParentObject, Epica_GUID).ObjTypeID & ", " & LocX & ", " & LocZ & ", " & Quantity & ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblComponentCache SET ComponentID = " & ComponentID & ", ComponentTypeID = " & _
				  ComponentTypeID & ", ComponentOwnerID = " & ComponentOwnerID & ", ParentID = " & _
				  CType(ParentObject, Epica_GUID).ObjectID & ", ParentTypeID = " & _
				  CType(ParentObject, Epica_GUID).ObjTypeID & ", LocX = " & LocX & ", LocZ = " & LocZ & _
				  ", Quantity = " & Quantity & " WHERE ComponentCacheID = " & Me.ObjectID
			End If
			Return sSQL
			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return ""
	End Function


    Public Sub DeleteMe()
        If Me.ObjectID > 0 Then
            Try
                Dim sSQL As String = "DELETE FROM tblComponentCache WHERE ComponentCacheID = " & Me.ObjectID
                Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm.Dispose()
                oComm = Nothing
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Unable to delete component cache: " & Me.ObjectID)
            End Try
        End If
    End Sub
End Class
