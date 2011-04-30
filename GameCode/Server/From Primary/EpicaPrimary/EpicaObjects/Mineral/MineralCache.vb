Public Class MineralCache
	Inherits Epica_GUID

	Public CacheTypeID As Byte
	Public ParentObject As Object		'where the cache exists
	Public LocX As Int32
	Public LocZ As Int32
	Public oMineral As Mineral
	Private mlQuantity As Int32
	Public Concentration As Int32
	Public OriginalConcentration As Int32

	Public lServerIndex As Int32

	Public BeingMinedBy As Epica_Entity

	Public bNeedsAsync As Boolean = False

	Private mySendString() As Byte

	Private mbIsDeleted As Boolean = False
	Public ReadOnly Property IsDeleted() As Boolean
		Get
			Return mbIsDeleted
		End Get
	End Property

	'Let's try this...
	Public Property Quantity() As Int32
		Get
			Return mlQuantity
		End Get
		Set(ByVal value As Int32)
			If value <> mlQuantity Then
				mlQuantity = value
				Me.DataChanged()

				bNeedsAsync = True

				If Concentration > mlQuantity OrElse CacheTypeID <> MineralCacheType.eMineable Then Concentration = mlQuantity

				'Now... if the quantity = 0 then we need to remove this cache from the DB
				If mlQuantity < 1 Then
                    'MSC - 05/23/08 - we no longer remove the cache, we set its quantity to 0 and leave it for reuse later
                    mlQuantity = 0

					'This is the case only for caches that are NOT mineable
					Dim iTemp As Int16 = -1
					If Me.ParentObject Is Nothing = False Then
						iTemp = CType(Me.ParentObject, Epica_GUID).ObjTypeID
					End If

					Dim bDelete As Boolean = False

					If Me.CacheTypeID = MineralCacheType.eMineable OrElse iTemp = -1 OrElse iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
						bDelete = True
					End If

					'see if we need to find our mineralcache
					If bDelete = True Then
						If Me.lServerIndex = -1 OrElse glMineralCacheIdx(Me.lServerIndex) <> Me.ObjectID Then
							For X As Int32 = 0 To glMineralCacheUB
								If glMineralCacheIdx(X) = Me.ObjectID Then
									Me.lServerIndex = X
									glMineralCacheIdx(X) = -1
									Exit For
								End If
							Next X
						Else : glMineralCacheIdx(Me.lServerIndex) = -1
						End If

						'Now... delete me
						DeleteMe()
					End If

					'Now... if this cache is in a planet or system...
					If Me.ParentObject Is Nothing = False Then
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
							'Else
							''most likely a facility or unit...
							'If iTemp = ObjectType.eFacility OrElse iTemp = ObjectType.eUnit Then
							'	'Ok, remove me from their cargo
							'	With CType(Me.ParentObject, Epica_Entity)
							'		For X As Int32 = 0 To .lCargoUB
							'			If .lCargoIdx(X) = Me.ObjectID AndAlso .oCargoContents(X).ObjTypeID = Me.ObjTypeID Then
							'				.lCargoIdx(X) = -1
							'				.oCargoContents(X) = Nothing
							'				Exit For
							'			End If
							'		Next X
							'	End With
							'End If
						End If
					End If

					If Me.CacheTypeID = MineralCacheType.eMineable Then
						'Test if this planet is depleted
						If iTemp = ObjectType.ePlanet Then
							Dim oParent As Planet = CType(ParentObject, Planet)
							If oParent.ySentGNSLowRes <> 0 AndAlso oParent.ObjectID < 500000000 Then
								Dim blTemp As Int64 = oParent.blOriginalMineralQuantity
								Dim lParentID As Int32 = oParent.ObjectID
								If blTemp > 0 Then
									Dim blCurrent As Int64 = 0
									For X As Int32 = 0 To glMineralCacheUB
										If glMineralCacheIdx(X) <> -1 Then
											Dim oCache As MineralCache = goMineralCache(X)
											If oCache Is Nothing = False AndAlso oCache.CacheTypeID = MineralCacheType.eMineable Then
												With CType(oCache.ParentObject, Epica_GUID)
													If .ObjTypeID = ObjectType.ePlanet AndAlso .ObjectID = lParentID Then
														blCurrent += oCache.Quantity
													End If
												End With
											End If
										End If
									Next X

									If blCurrent / blTemp < 0.25F Then
										oParent.ySentGNSLowRes = 255
										Dim yGNS(98) As Byte
										Dim lPos As Int32 = 0
										System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
										yGNS(lPos) = NewsItemType.ePlanetWideLowResources : lPos += 1
										oParent.ParentSystem.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6

										oParent.PlanetName.CopyTo(yGNS, lPos) : lPos += 20

										Dim oMin As Mineral = GetEpicaMineral(oParent.lPrimaryComposition)
										If oMin Is Nothing = False Then
											oMin.MineralName.CopyTo(yGNS, lPos) : lPos += 20
										Else : StringToBytes("Various Metals").CopyTo(yGNS, lPos) : lPos += 20
										End If
										System.BitConverter.GetBytes(oParent.lPrimaryComposition).CopyTo(yGNS, lPos) : lPos += 4
										System.BitConverter.GetBytes(oParent.OwnerID).CopyTo(yGNS, lPos) : lPos += 4

										If oParent.OwnerID > 0 Then
											Dim oTmpPlayer As Player = GetEpicaPlayer(oParent.OwnerID)
											If oTmpPlayer Is Nothing = False Then
												oTmpPlayer.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
												yGNS(lPos) = oTmpPlayer.yGender : lPos += 1
												oTmpPlayer.EmpireName.CopyTo(yGNS, lPos) : lPos += 20
											Else : lPos += 41
											End If
										Else : lPos += 41
										End If

										goMsgSys.SendToEmailSrvr(yGNS)
									End If
								End If
							End If
						End If

						'Here is where redistribution occurs
                        If oMineral Is Nothing = False AndAlso oMineral.ObjectID <> 41991 Then
                            Dim lID As Int32 = -1
                            Dim iTypeID As Int16 = -1
                            If Me.ParentObject Is Nothing = False Then
                                With CType(Me.ParentObject, Epica_GUID)
                                    lID = .ObjectID : iTypeID = .ObjTypeID
                                End With
                            End If

                            oMineral.RespawnMe(lID, iTypeID)
                        End If
					End If
				End If
			End If
		End Set
	End Property

	Public Function GetObjAsString() As Byte()
		'here we will return the entire object as a string
		If mbStringReady = False Then
			ReDim mySendString(32)		'0 to 32 = 33 bytes

			GetGUIDAsString.CopyTo(mySendString, 0)
			mySendString(6) = CacheTypeID
			CType(ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, 7)
			System.BitConverter.GetBytes(LocX).CopyTo(mySendString, 13)
			System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, 17)
			System.BitConverter.GetBytes(oMineral.ObjectID).CopyTo(mySendString, 21)
			System.BitConverter.GetBytes(Quantity).CopyTo(mySendString, 25)
			System.BitConverter.GetBytes(Concentration).CopyTo(mySendString, 29)

			mbStringReady = True
		End If
		Return mySendString
	End Function

	Public Function GetSaveObjectText() As String
		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
		If ObjectID = -1 Then
			SaveObject()
			Return ""
		End If
		Try
			Return "UPDATE tblMineralCache SET CacheTypeID = " & CacheTypeID & ", ParentID = " & CType(ParentObject, Epica_GUID).ObjectID & _
			", ParentTypeID = " & CType(ParentObject, Epica_GUID).ObjTypeID & ", LocX = " & LocX & ", LocY = " & LocZ & _
			", MineralID = " & oMineral.ObjectID & ", Quantity = " & Quantity & ", Concentration = " & _
			Concentration & ", OriginalConcentration = " & OriginalConcentration & " WHERE CacheID = " & ObjectID
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "GetSaveObjectText of MineralCache: " & ex.Message)
		End Try
		Return ""
	End Function

	Public Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            Dim bRemoval As Boolean = False
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblMineralCache (CacheTypeID, ParentID, ParentTypeID, LocX, LocY, MineralID, " & _
                  "Quantity, Concentration, OriginalConcentration) VALUES (" & CacheTypeID & ", " & _
                  CType(ParentObject, Epica_GUID).ObjectID & ", " & CType(ParentObject, Epica_GUID).ObjTypeID & ", " & LocX & ", " & _
                  LocZ & ", " & oMineral.ObjectID & ", " & Quantity & ", " & Concentration & ", " & OriginalConcentration & " )"
            ElseIf ParentObject Is Nothing = False Then
                'UPDATE
                sSQL = "UPDATE tblMineralCache SET CacheTypeID = " & CacheTypeID & ", ParentID = " & CType(ParentObject, Epica_GUID).ObjectID & _
                  ", ParentTypeID = " & CType(ParentObject, Epica_GUID).ObjTypeID & ", LocX = " & LocX & ", LocY = " & LocZ & _
                  ", MineralID = " & oMineral.ObjectID & ", Quantity = " & Quantity & ", Concentration = " & _
                  Concentration & ", OriginalConcentration = " & OriginalConcentration & " WHERE CacheID = " & ObjectID
            Else
                bRemoval = True
                sSQL = "DELETE FROM tblMineralCache WHERE CacheID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If bRemoval = True Then Return True
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(CacheID) FROM tblMineralCache WHERE CacheTypeID = " & CacheTypeID & " AND MineralID = " & oMineral.ObjectID
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
		bNeedsAsync = False
		Return bResult
	End Function

	Public Sub HandleDepletion(ByVal fFactor As Single)
        If Concentration <> 1 Then
            If OriginalConcentration = 0 Then OriginalConcentration = Concentration
            Dim lVal As Int32 = CInt(Int(Rnd() * Math.Pow(OriginalConcentration, fFactor)))
            If lVal < Concentration Then
                Concentration -= 1
                bNeedsAsync = True
            End If
        End If
	End Sub

	Public Sub DeleteMe()
		mbIsDeleted = True

		If BeingMinedBy Is Nothing = False Then
			BeingMinedBy.lCacheIndex = -1
			BeingMinedBy.lCacheID = -1
		End If

		If Me.ObjectID > 0 Then
			Try
				Dim sSQL As String = "DELETE FROM tblMineralCache WHERE CacheID = " & Me.ObjectID
				Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
				oComm.ExecuteNonQuery()
				oComm.Dispose()
				oComm = Nothing
			Catch ex As Exception
				LogEvent(LogEventType.CriticalError, "Unable to delete mineral cache: " & Me.ObjectID)
			End Try
		End If
		bNeedsAsync = False
	End Sub
End Class
