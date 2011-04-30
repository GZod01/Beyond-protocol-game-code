Option Strict On

'Manages a Sell Order - Someone wishes to sell something they already have
Public Class SellOrder
    'PK BEGIN
    Public TradePostID As Int32         'ID of the facility that is the tradepost (source)
    Public lTradeID As Int32            'ID being sold 
    Public iTradeTypeID As Int16        'TypeID being sold 
	'END OF PK

	Public iExtTypeID As Int16 = 0

    Public blQuantity As Int64          'Quantity being sold or bought
    Public blPrice As Int64             'Price per item

    Public yItemName(19) As Byte 

	Public Const l_MESSAGE_LENGTH As Int32 = 52
    Public Function GetObjAsString() As Byte()

        If TradePost Is Nothing OrElse TradePost.Owner Is Nothing Then Return Nothing

        Dim yMsg(l_MESSAGE_LENGTH - 1) As Byte
        Dim lPos As Int32 = 0

        'PK
        System.BitConverter.GetBytes(TradePostID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lTradeID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(iTradeTypeID).CopyTo(yMsg, lPos) : lPos += 2
        'end pk

        System.BitConverter.GetBytes(blQuantity).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blPrice).CopyTo(yMsg, lPos) : lPos += 8 
        System.BitConverter.GetBytes(GetItemScore).CopyTo(yMsg, lPos) : lPos += 4

		yItemName.CopyTo(yMsg, lPos) : lPos += 20

		System.BitConverter.GetBytes(iExtTypeID).CopyTo(yMsg, lPos) : lPos += 2

        Return yMsg
    End Function

    Private mlItemScore As Int32 = Int32.MinValue
    Public ReadOnly Property GetItemScore() As Int32
		Get
			If mlItemScore = Int32.MinValue Then
				Select Case iTradeTypeID
					Case ObjectType.eMineralCache
						Dim oCache As MineralCache = GetEpicaMineralCache(lTradeID)
						If oCache Is Nothing = False Then mlItemScore = oCache.oMineral.GetMineralScore()
						oCache = Nothing
					Case ObjectType.eMineral
						Dim oMineral As Mineral = GetEpicaMineral(lTradeID)
						If oMineral Is Nothing = False Then mlItemScore = oMineral.GetMineralScore()
						oMineral = Nothing
					Case ObjectType.eUnit
						Dim oUnit As Unit = GetEpicaUnit(lTradeID)
						If oUnit Is Nothing = False Then mlItemScore = oUnit.EntityDef.CombatRating()
						oUnit = Nothing
					Case Is < 0
						For X As Int32 = 0 To TradePost.lCargoUB
							If TradePost.lCargoIdx(X) <> -1 AndAlso TradePost.oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
								With CType(TradePost.oCargoContents(X), ComponentCache)
									If .ComponentID = lTradeID AndAlso .ComponentTypeID = Math.Abs(iTradeTypeID) Then
										mlItemScore = .GetComponent.TechnologyScore
										Exit For
									End If
								End With
							End If
						Next X
					Case Else
						mlItemScore = 0
				End Select
			End If
			Return mlItemScore
		End Get
    End Property

    Private moTradePost As Facility = Nothing
    Public ReadOnly Property TradePost() As Facility
        Get
            If moTradePost Is Nothing Then moTradePost = GetEpicaFacility(TradePostID)
            Return moTradePost
        End Get
    End Property

    ''' <summary>
    ''' Thread safely checks and removes quantity from the param passed returning true if the quantity is available and accounted for.
    ''' </summary>
    ''' <param name="blQty"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PurchaseMe(ByVal blQty As Int64) As Boolean
		Dim bResult As Boolean = False

		If iTradeTypeID = ObjectType.ePlayerIntel OrElse iTradeTypeID = ObjectType.ePlayerItemIntel OrElse iTradeTypeID = ObjectType.ePlayerTechKnowledge Then
			Return True
		End If

        SyncLock Me
            If blQuantity >= blQty Then
                blQuantity -= blQty
                bResult = True
            End If
        End SyncLock
        Return bResult
	End Function

	Public Function GetQuantityInStock() As Int64
		Dim oTP As Facility = TradePost
		If oTP Is Nothing = False Then
            If oTP.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return blQuantity

            Dim oColony As Colony = oTP.ParentColony
            If oColony Is Nothing Then Return 0

            If iTradeTypeID = ObjectType.eUnit Then
                Dim oUnit As Unit = GetEpicaUnit(lTradeID)
                If oUnit Is Nothing = False Then
                    Dim oObj As Epica_GUID = CType(oUnit.ParentObject, Epica_GUID)
                    If oObj Is Nothing = False Then
                        If oObj.ObjTypeID = ObjectType.eFacility Then
                            Dim oFac As Facility = CType(oObj, Facility)
                            If oFac.yProductionType <> ProductionType.eTradePost Then Return 0
                            If oFac.ParentColony Is Nothing = False AndAlso oFac.ParentColony.ObjectID = oColony.ObjectID Then Return 1
                        End If
                    End If
                End If
            ElseIf iTradeTypeID < 0 Then
                'Component Cache
                Dim lQty As Int32 = 0
                If oTP.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                    If oTP Is Nothing = False Then
                        If (oTP.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then

                            For Y As Int32 = 0 To oTP.lCargoUB
                                If oTP.lCargoIdx(Y) <> -1 AndAlso oTP.oCargoContents(Y).ObjTypeID = ObjectType.eComponentCache Then
                                    Dim oCache As ComponentCache = CType(oTP.oCargoContents(Y), ComponentCache)
                                    If oCache.ComponentID = lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(iTradeTypeID) Then
                                        lQty += oCache.Quantity
                                        Exit For
                                    End If
                                End If
                            Next Y
                        End If
                    End If
                Else
                    'For X As Int32 = 0 To oColony.ChildrenUB
                    '    If oColony.lChildrenIdx(X) > -1 Then
                    '        Dim oFac As Facility = oColony.oChildren(X)
                    '        If oFac Is Nothing = False Then
                    '            If (oFac.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then

                    '                For Y As Int32 = 0 To oFac.lCargoUB
                    '                    If oFac.lCargoIdx(Y) <> -1 AndAlso oFac.oCargoContents(Y).ObjTypeID = ObjectType.eComponentCache Then
                    '                        Dim oCache As ComponentCache = CType(oFac.oCargoContents(Y), ComponentCache)
                    '                        If oCache.ComponentID = lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(iTradeTypeID) Then
                    '                            lQty += oCache.Quantity
                    '                            Exit For
                    '                        End If
                    '                    End If
                    '                Next Y
                    '            End If
                    '        End If
                    '    End If
                    'Next X
                    For X As Int32 = 0 To oColony.mlComponentCacheUB
                        If oColony.mlComponentCacheIdx(X) > -1 Then
                            If oColony.mlComponentCacheID(X) = glComponentCacheIdx(oColony.mlComponentCacheIdx(X)) Then
                                Dim oCache As ComponentCache = goComponentCache(oColony.mlComponentCacheIdx(X))
                                If oCache Is Nothing = False Then
                                    If oCache.ComponentID = lTradeID AndAlso oCache.ComponentTypeID = Math.Abs(iTradeTypeID) Then
                                        lQty += oCache.Quantity
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next X
                End If
                
                Return lQty
                
            ElseIf iTradeTypeID = ObjectType.eMineralCache Then
                'Ok, mineral caches
                Dim oCache As MineralCache = GetEpicaMineralCache(lTradeID)
                If oCache Is Nothing = False Then
                    Dim oObj As Epica_GUID = CType(oCache.ParentObject, Epica_GUID)
                    If oObj Is Nothing = False Then
                        If oObj.ObjTypeID = ObjectType.eFacility Then
                            Dim oFac As Facility = CType(oObj, Facility)
                            If oFac.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID OrElse (oFac.ParentColony Is Nothing = False AndAlso oFac.ParentColony.ObjectID = oColony.ObjectID) Then Return oCache.Quantity
                        ElseIf oObj.ObjTypeID = ObjectType.eColony Then
                            Dim oTmpColony As Colony = CType(oObj, Colony)
                            If oTmpColony.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID OrElse (oTmpColony.ObjectID = oColony.ObjectID) Then Return oCache.Quantity
                        End If
                    End If
                End If
            ElseIf iTradeTypeID = ObjectType.ePlayerIntel OrElse iTradeTypeID = ObjectType.ePlayerItemIntel OrElse iTradeTypeID = ObjectType.ePlayerTechKnowledge Then
                Return 1
            Else
                'TODO: what else?
                'for Component Designs...
                ' store the ComponentID in the ObjectID
                ' store the ComponentTypeID as the typeid
                ' store the componentownerid as the extendedid
            End If
        End If
        Return 0
	End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try

			sSQL = "UPDATE tblSellOrder SET Quantity = " & blQuantity & ", Price = " & blPrice & ", ItemName = '" & _
			 MakeDBStr(BytesToString(yItemName)) & "', ExtTypeID = " & iExtTypeID & " WHERE TradePostID = " & _
			 TradePostID & " AND lTradeID = " & lTradeID & " AND iTradeTypeID = " & iTradeTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                oComm = Nothing
				sSQL = "INSERT INTO tblSellOrder (TradePostID, lTradeID, iTradeTypeID, Quantity, Price, ItemName, extTypeID) VALUES (" & _
				  TradePostID & ", " & lTradeID & ", " & iTradeTypeID & ", " & blQuantity & ", " & blPrice & ", '" & _
				  MakeDBStr(BytesToString(yItemName)) & "', " & iExtTypeID & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
            End If 

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object SellOrder. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub DeleteMe()
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            sSQL = "DELETE FROM tblSellOrder WHERE TradePostID = " & TradePostID & " AND lTradeID = " & lTradeID & " AND iTradeTypeID = " & iTradeTypeID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to DELETE object SellOrder. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
    End Sub
End Class
