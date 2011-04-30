Public Class AmmunitionCache
    Inherits Epica_GUID

    Private mlQuantity As Int32
    Public oWeaponTech As BaseWeaponTech        'the weapon this ammo is used in (POINTER!!!)    

    Public ParentObject As Object       'where the cache exists

    Public LocX As Int32
    Public LocZ As Int32

    Private mySendString() As Byte
 
    Public Property Quantity() As Int32
        Get
            Return mlQuantity
        End Get
        Set(ByVal value As Int32)
            If value <> mlQuantity Then
                mlQuantity = value
                Me.DataChanged()

                'Now... if the quantity = 0 then we need to remove this cache from the DB
                If mlQuantity < 1 Then

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
                        End If
                    End If
                End If
            End If
        End Set
    End Property

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
            ReDim mySendString(27)      '0 to 27 = 28 bytes

            Dim lPos As Int32 = 0

            GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            CType(ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            System.BitConverter.GetBytes(LocX).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, lPos) : lPos += 4
            System.BitConverter.GetBytes(Quantity).CopyTo(mySendString, lPos) : lPos += 4
            If oWeaponTech Is Nothing = False Then
                System.BitConverter.GetBytes(oWeaponTech.ObjectID).CopyTo(mySendString, lPos) : lPos += 4
            Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, lPos) : lPos += 4
            End If

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
        If oWeaponTech Is Nothing Then Return True
        If mlQuantity < 1 Then Return True

        Try
            If ObjectID = -1 Then
                'INSERT
                With CType(ParentObject, Epica_GUID)
                    sSQL = "INSERT INTO tblAmmoCache (WeaponTechID, Quantity, ParentID, ParentTypeID, LocX, LocZ) VALUES (" & _
                      oWeaponTech.ObjectID & ", " & mlQuantity & ", " & .ObjectID & ", " & .ObjTypeID & ", " & LocX & ", " & LocZ & ")"
                End With
            Else
                'UPDATE
                With CType(ParentObject, Epica_GUID)
                    sSQL = "UPDATE tblAmmoCache SET WeaponTechID = " & oWeaponTech.ObjectID & ", Quantity = " & mlQuantity & _
                      ", ParentID = " & .ObjectID & ", ParentTypeID = " & .ObjTypeID & ", LocX = " & LocX & ", LocZ = " & _
                      LocZ & " WHERE AmmoCacheID = " & Me.ObjectID
                End With
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(AmmoCacheID) FROM tblAmmoCache WHERE WeaponTechID = " & oWeaponTech.ObjectID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub DeleteMe()
        If Me.ObjectID > 0 Then
            Try
                Dim sSQL As String = "DELETE FROM tblAmmoCache WHERE AmmoCacheID = " & Me.ObjectID
                Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm.Dispose()
                oComm = Nothing
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Unable to delete Ammo cache: " & Me.ObjectID)
            End Try
        End If
    End Sub

    'Public Function CacheCargoSpace(ByVal lQty As Int32) As Int32
    '    Dim lResult As Int32 = lQty
    '    If oWeaponTech Is Nothing = False Then
    '        'Ok...
    '        Dim fSize As Single = oWeaponTech.GetAmmoSize()
    '        lResult = CInt(Math.Ceiling(fSize * Quantity))
    '    End If
    '    Return lResult
    'End Function

End Class
