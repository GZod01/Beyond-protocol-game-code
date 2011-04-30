Option Strict On

Public Class Envir
    'this tracks the environment objects... an environment is any location that a unit can exist
    Inherits Epica_GUID

    Public Enum eyChangePlayerEnvirCode As Byte
        eAddToEnvir = 1
        eRemoveFromEnvir = 0
    End Enum

    Public oGeoObject As Object                 'reference to the geography object that this environment represents

    Public PotentialAggression As Boolean       'this flag indicates whether units exist in this environment that belong to waring factions
    Public PotentialFirstContact As Boolean     'this flag indicates that units exist in this environment that belong to players that do not know each other... this flag is only valid in PLANET environments

    Public oPlayersWhoHaveUnitsHere() As Player
    Public lPlayersWhoHaveUnitsHereCP() As Int32        'command points used for this environment
    Public lPlayersWhoHaveUnitsHereIdx() As Int32
    Public lPlayersWhoHaveUnitsHereUB As Int32 = -1

    Public oPlayersInEnvir() As Player
    Public lPlayersInEnvirIdx() As Int32
    Public lPlayersInEnvirUB As Int32 = -1

    Public lPlayersInEnvirCnt As Int32 = 0      'state machine for determining players in the environment...

#Region "  Do Lookups and such  "
    Private mcolEnvirGrid As New Collection
    Private mlEnvirGridLock(-1) As Int32
    Private Enum elLookupGridType As Int32
        LookupOnly = 0
        AddGrid = 1
        RemoveGrid = 2
        LookupOrAdd = 3
        ClearAndRebuildGrid = 9000
        CheckForExpiration = 9001
        TestForPlayerWithUnitsHere = 9002
    End Enum
    Private Function DoLookupGridFunction(ByVal lType As elLookupGridType, ByVal lGridID As Int32) As EnvirGrid
        Dim bNeedToTestAggression As Boolean = False

        SyncLock mlEnvirGridLock
            If mcolEnvirGrid Is Nothing Then mcolEnvirGrid = New Collection
            Try
                Dim sKey As String = "G" & lGridID
                Select Case lType
                    Case elLookupGridType.LookupOnly
                        If mcolEnvirGrid.Contains(sKey) = True Then
                            Return CType(mcolEnvirGrid(sKey), EnvirGrid)
                        Else : Return Nothing
                        End If
                    Case elLookupGridType.AddGrid
                        If mcolEnvirGrid.Contains(sKey) = True Then Return Nothing

                        Dim oNew As New EnvirGrid
                        oNew.lGridID = lGridID
                        mcolEnvirGrid.Add(oNew, sKey)
                    Case elLookupGridType.RemoveGrid
                        If mcolEnvirGrid.Contains(sKey) = True Then mcolEnvirGrid.Remove(sKey)
                    Case elLookupGridType.LookupOrAdd
                        If mcolEnvirGrid.Contains(sKey) = True Then
                            Return CType(mcolEnvirGrid(sKey), EnvirGrid)
                        Else
                            Dim oNew As New EnvirGrid
                            oNew.lGridID = lGridID
                            mcolEnvirGrid.Add(oNew, sKey)
                            Return CType(mcolEnvirGrid(sKey), EnvirGrid)
                        End If
                    Case elLookupGridType.ClearAndRebuildGrid
                        mcolEnvirGrid = Nothing
                        mcolEnvirGrid = New Collection
                        mcolEnvirGrid.Clear()
                        For X As Int32 = 0 To glEntityUB
                            If glEntityIdx(X) > -1 Then
                                Dim oEnt As Epica_Entity = goEntity(X)
                                If oEnt Is Nothing = False Then
                                    If oEnt.ParentEnvir Is Nothing = False AndAlso oEnt.ParentEnvir.ObjectID = Me.ObjectID AndAlso oEnt.ParentEnvir.ObjTypeID = Me.ObjTypeID Then
                                        oEnt.lGridEntityIdx = Me.oGrid(oEnt.lGridIndex).AddEntity(X)
                                    End If
                                End If
                            End If
                        Next X
                    Case elLookupGridType.CheckForExpiration
                        Dim oRemoveGrid As EnvirGrid = Nothing
                        For Each oGrid As EnvirGrid In mcolEnvirGrid
                            If oGrid.lExpirationCycle < glCurrentCycle Then
                                If oGrid.CheckForExpiration() = True Then
                                    oRemoveGrid = oGrid
                                    Exit For
                                Else
                                    oGrid.lExpirationCycle = Int32.MaxValue
                                End If
                            End If
                        Next
                        If oRemoveGrid Is Nothing = False Then
                            sKey = "G" & oRemoveGrid.lGridID
                            mcolEnvirGrid.Remove(sKey)
                        End If
                        Return oRemoveGrid
                    Case elLookupGridType.TestForPlayerWithUnitsHere
                        For X As Int32 = 0 To lPlayersWhoHaveUnitsHereUB
                            If lPlayersWhoHaveUnitsHereIdx(X) <> -1 Then
                                Dim bFound As Boolean = False
                                If mcolEnvirGrid Is Nothing = False Then
                                    For Each oGrid As EnvirGrid In mcolEnvirGrid
                                        If oGrid Is Nothing Then Continue For
                                        For Y As Int32 = 0 To oGrid.lEntityUB
                                            If oGrid.lEntities(Y) <> -1 Then
                                                If glEntityIdx(oGrid.lEntities(Y)) > 0 Then
                                                    Dim oEntity As Epica_Entity = goEntity(oGrid.lEntities(Y))
                                                    If oEntity Is Nothing = False Then
                                                        If oEntity.Owner.ObjectID = lPlayersWhoHaveUnitsHereIdx(X) Then
                                                            bFound = True
                                                            Exit For
                                                        End If
                                                    Else : oGrid.lEntities(Y) = -1
                                                    End If
                                                Else : oGrid.lEntities(Y) = -1
                                                End If
                                            End If
                                        Next Y
                                    Next
                                End If

                                If bFound = False Then
                                    lPlayersWhoHaveUnitsHereIdx(X) = -1
                                    oPlayersWhoHaveUnitsHere(X) = Nothing
                                    bNeedToTestAggression = True
                                End If
                            End If
                        Next X
                End Select
            Catch
            End Try
        End SyncLock

        If bNeedToTestAggression = True Then
            TestForAggression()
            TestForFirstContact()
        End If

        Return Nothing
    End Function
#End Region

    'Grids now contain our entities
    'Private moGrid() As EnvirGrid
    Public Function oGrid(ByVal lIndex As Int32) As EnvirGrid
        Return DoLookupGridFunction(elLookupGridType.LookupOrAdd, lIndex)

        'Dim lIdx As Int32 = DoLookupGridFunction(0, lIndex, -1)
        'If lIdx > -1 Then Return moGrid(lIdx)

        'If moGrid Is Nothing Then ReDim moGrid(-1)
        'SyncLock moGrid

        '    Dim oNew As New EnvirGrid
        '    oNew.lGridID = lIndex

        '    Dim lUB As Int32 = moGrid.GetUpperBound(0)
        '    lUB += 1
        '    ReDim Preserve moGrid(lUB)
        '    moGrid(lUB) = oNew
        '    lIdx = lUB
        '    DoLookupGridFunction(1, lIndex, lUB)
        'End SyncLock

        'Return moGrid(lIdx)
    End Function
    Public Function GetGridUB() As Int32
        Return lGridUB 'moGrid.GetUpperBound(0)
    End Function
    Public Function GetGridNoAdd(ByVal lIndex As Int32) As EnvirGrid
        Return DoLookupGridFunction(elLookupGridType.LookupOnly, lIndex)
        'Dim lIdx As Int32 = DoLookupGridFunction(0, lIndex, -1)
        'If lIdx = -1 Then Return Nothing
        'If moGrid(lIdx) Is Nothing Then Return Nothing
        'Return moGrid(lIdx)
    End Function

    'These values are used for determining sector changes and are important to be configured properly in SetEnvirGridValues
    Public lGridUB As Int32 = -1        '=(lGridsPerRow * lGridsPerRow) - 1
    Public lEnvirSize As Int32          'the maximum size
    Public lHalfEnvirSize As Int32      'half of lEnvirSize
    Public lGridSquareSize As Int32     'If System = 8000 elseif Planet = 1600
    Public lGridsPerRow As Int32        '=lEnvirSize / lGridSquareSize (NOTE: MUST DIVIDE EVENLY!!!)

    Public lGridIdxAdjust() As Int32    'for grid index lookup optimization
    Public lLeftEdgeGridIdxAdjust() As Int32
    Public lRightEdgeGridIdxAdjust() As Int32

    'These are the extents... set by SetEnvirGridValues
    Public lMinPosX As Int32
    Public lMinPosZ As Int32
    Public lMaxPosX As Int32
    Public lMaxPosZ As Int32
    Public lMapWrapAdjustX As Int32

    Public oCache() As ObjectCache
    Public lCacheIdx() As Int32
    Public lCacheUB As Int32 = -1

    Private myFacBurst() As Byte
    Private mbFacBurstDirty As Boolean = True
    Private myCacheBurst() As Byte
    Private mbCacheBurstDirty As Boolean = True

    Public lOBShots As Int32 = 0
    Public lOBShotCycle As Int32 = 0

    Public bEnvirAtColonyLimit As Boolean = False

    Public Function GenerateBurstMessage(ByVal lForPlayerID As Int32) As Byte()
        'ok, we do this... very... interestingly.
        Dim yReturn() As Byte
        Dim X As Int32
        Dim lPos As Int32
        Dim yTemp() As Byte
        'Dim lGrid As Int32
        Dim lUnitIdx As Int32
        Dim lLenMove As Int32
        Dim lLenNotMove As Int32
        Dim lLen As Int32

        Dim lCurUB As Int32 = -1

        ReDim yTemp(200000)     'initialize the array for 200k bytes

        lPos = 0

        Dim lstIndices As New System.Collections.Generic.List(Of Int32)
        lCurUB = -1
        If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        For X = 0 To lCurUB
            If glEntityIdx(X) > -1 Then
                Dim oEntity As Epica_Entity = goEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.ParentEnvir Is Nothing = False AndAlso oEntity.ParentEnvir.ObjectID = Me.ObjectID AndAlso oEntity.ParentEnvir.ObjTypeID = Me.ObjTypeID Then
                    lstIndices.Add(X)
                End If
            End If
        Next X

        For X = 0 To lstIndices.Count - 1
            lUnitIdx = lstIndices(X)
            If glEntityIdx(lUnitIdx) > 0 Then
                If (goEntity(lUnitIdx).CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 AndAlso goEntity(lUnitIdx).lOwnerID <> lForPlayerID Then Continue For
                If (goEntity(lUnitIdx).CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                    If lLenMove = 0 Then
                        lLenMove = goEntity(lUnitIdx).GetObjectAsSmallString.Length + 2
                    End If
                    lLen = lLenMove
                Else
                    If lLenNotMove = 0 Then
                        lLenNotMove = goEntity(lUnitIdx).GetObjectAsSmallString.Length + 2
                    End If
                    lLen = lLenNotMove
                End If
                'Now, check if we need to increase our cache
                If lPos + lLen + 2 > yTemp.Length Then
                    ReDim Preserve yTemp(yTemp.Length + 200000)
                End If
                System.BitConverter.GetBytes(CShort(lLen)).CopyTo(yTemp, lPos)
                lPos += 2
                'now, add the add object message ID
                System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yTemp, lPos)
                lPos += 2
                'Now, add the unit
                goEntity(lUnitIdx).GetObjectAsSmallString.CopyTo(yTemp, lPos)
                lPos += lLen - 2
            End If
        Next X

        For X = 0 To lstIndices.Count - 1
            lUnitIdx = lstIndices(X)
            If glEntityIdx(lUnitIdx) > 0 Then
                If (goEntity(lUnitIdx).CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 AndAlso goEntity(lUnitIdx).lOwnerID <> lForPlayerID Then Continue For
                If (goEntity(lUnitIdx).CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                    Dim yAddMsg() As Byte = goMsgSys.CreateSetTargetMessage(goEntity(lUnitIdx))
                    If yAddMsg Is Nothing = False Then
                        'Now, check if we need to increase our cache
                        lLen = yAddMsg.Length
                        If lPos + lLen + 2 > yTemp.Length Then
                            ReDim Preserve yTemp(yTemp.Length + 200000)
                        End If
                        System.BitConverter.GetBytes(CShort(lLen)).CopyTo(yTemp, lPos)
                        lPos += 2
                        'Now, add the unit
                        yAddMsg.CopyTo(yTemp, lPos)
                        lPos += lLen '- 2
                    End If
                End If
            End If
        Next X

        lCurUB = -1
        If lCacheIdx Is Nothing = False Then lCurUB = Math.Min(lCacheUB, lCacheIdx.GetUpperBound(0))
        For X = 0 To lCurUB
            If lCacheIdx(X) <> -1 Then
                lLen = ObjectCache.l_CACHE_MSG_LEN + 2
                'Now, check if we need to increase our cache
                If lPos + lLen + 2 > yTemp.Length Then
                    ReDim Preserve yTemp(yTemp.Length + 200000)
                End If
                System.BitConverter.GetBytes(CShort(lLen)).CopyTo(yTemp, lPos)
                lPos += 2
                'now, add the add object message ID
                System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yTemp, lPos)
                lPos += 2
                'Now, add the cache
                oCache(X).GetObjectAsSmallString.CopyTo(yTemp, lPos)
                lPos += lLen - 2
            End If
        Next X

        'If mbFacBurstDirty = False Then
        '    lFinalSize += myFacBurst.Length
        '    myFacBurst.CopyTo(yTemp, lPos)
        '    lPos += myFacBurst.Length
        'Else
        '    lFacSize = lFinalSize
        '    For X = 0 To lFacilityUB
        '        If lFacilityIdx(X) <> -1 Then
        '            oFacility(X).GetObjectAsSmallString.CopyTo(yTemp, lPos)
        '            lPos += 21
        '            lFinalSize += 21
        '        End If
        '    Next X
        '    'Now, fill our myFacBurst portion...
        '    ReDim myFacBurst(lFinalSize - lFacSize - 1)
        '    yTemp.Copy(yTemp, lFacSize, myFacBurst, 0, myFacBurst.Length)
        '    lFacSize = myFacBurst.Length
        '    mbFacBurstDirty = False
        'End If

        'If mbCacheBurstDirty = False Then
        '    lFinalSize += myCacheBurst.Length
        '    myCacheBurst.CopyTo(yTemp, lPos)
        '    lPos += myCacheBurst.Length
        'Else
        '    lCacheSize = lFinalSize
        '    For X = 0 To lCacheUB
        '        If lCacheIdx(X) <> -1 Then
        '            oCache(X).GetObjectAsSmallString.CopyTo(yTemp, lPos)
        '            lPos += 19
        '            lFinalSize += 19
        '        End If
        '    Next X
        '    'Now, fill our myCacheBurst
        '    ReDim myCacheBurst(lFinalSize - lCacheSize - 1)
        '    yTemp.Copy(yTemp, lCacheSize, myCacheBurst, 0, myCacheBurst.Length)
        '    lCacheSize = myCacheBurst.Length
        '    mbCacheBurstDirty = False
        'End If

        'Now, check any players that are in an Iron Curtain state...
        'If Me.ObjTypeID = ObjectType.ePlanet Then
        'Dim lCurUB As Int32 = Math.Min(glPlayerUB, glPlayerIdx.GetUpperBound(0))
        lCurUB = -1
        If glPlayerIdx Is Nothing = False Then lCurUB = Math.Min(glPlayerUB, glPlayerIdx.GetUpperBound(0))
        For X = 0 To lCurUB
            If glPlayerIdx(X) <> -1 Then
                If goPlayers(X).lIronCurtainPlanetID = Me.ObjectID OrElse goPlayers(X).lIronCurtainPlanetID = Int32.MinValue Then
                    'ok, got one...
                    Dim yNewMsg(10) As Byte
                    Dim lNewMsgPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetIronCurtain).CopyTo(yNewMsg, lNewMsgPos) : lNewMsgPos += 2
                    System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yNewMsg, lNewMsgPos) : lNewMsgPos += 4
                    System.BitConverter.GetBytes(goPlayers(X).ObjectID).CopyTo(yNewMsg, lNewMsgPos) : lNewMsgPos += 4
                    yNewMsg(lNewMsgPos) = 1 : lNewMsgPos += 1

                    lLen = yNewMsg.Length
                    If lPos + lLen + 2 > yTemp.Length Then
                        ReDim Preserve yTemp(yTemp.Length + 200000)
                    End If
                    System.BitConverter.GetBytes(CShort(lLen)).CopyTo(yTemp, lPos)
                    lPos += 2
                    yNewMsg.CopyTo(yTemp, lPos)
                    lPos += lLen
                End If
            End If
        Next X
        'End If

        'Now, the final message we want them to receive is the response message
        If yTemp.Length < lPos + 4 Then
            ReDim Preserve yTemp(yTemp.Length + 10)
        End If
        System.BitConverter.GetBytes(CShort(2)).CopyTo(yTemp, lPos)
        lPos += 2
        System.BitConverter.GetBytes(GlobalMessageCode.eBurstEnvironmentResponse).CopyTo(yTemp, lPos)
        lPos += 2

        ReDim yReturn(lPos - 1)
        Array.Copy(yTemp, 0, yReturn, 0, lPos - 1)

        'Disable this when we are in go-live
        'gfrmDisplayForm.AddEventLine("Burst Msg Len for " & Me.ObjectID & ", " & Me.ObjTypeID & ": " & yReturn.Length)

        Return yReturn

        'oDomainSocket.SendLenAppendedData(yFinal)

        'System.BitConverter.GetBytes(lFinalSize).CopyTo(yTemp, 2)

        'now... do our deal
        'ReDim yReturn(lFinalSize - 1)
        'ReDim yReturn(lFinalSize + 5)
        'yTemp.Copy(yTemp, yReturn, lFinalSize)
        'yTemp.Copy(yTemp, yReturn, lFinalSize + 6)
        'Erase yTemp
        'Return yReturn

    End Function

	'NOTE: it is best to use this sub as opposed to doing the work yourself
	Public Sub AddEntity(ByVal oNewEntity As Epica_Entity)
		Dim X As Int32
		Dim lIdx As Int32
		Dim bFound As Boolean

		'Ok, determine the entity's grid
		Dim lGridID As Int32
		Dim lTemp As Int32
		Dim lTemp2 As Int32

		Dim lSmallPerRow As Int32
		'If ObjTypeID = ObjectType.ePlanet Then
		'    lSmallPerRow = gl_PLANET_SMALL_PER_ROW
		'Else
		'    lSmallPerRow = gl_SYSTEM_SMALL_PER_ROW
		'End If
		lSmallPerRow = gl_SMALL_PER_ROW

        'Dim bDone As Boolean = False
        'While bDone = False
        '	bDone = True
        '	If oNewEntity.LocX < Me.lMinPosX Then
        '		If Me.ObjTypeID = ObjectType.ePlanet Then oNewEntity.LocX += Me.lMapWrapAdjustX Else oNewEntity.LocX = Me.lMinPosX + 100
        '		bDone = False
        '	ElseIf oNewEntity.LocX > Me.lMaxPosX Then
        '		If Me.ObjTypeID = ObjectType.ePlanet Then oNewEntity.LocX -= Me.lMapWrapAdjustX Else oNewEntity.LocX = Me.lMaxPosX - 100
        '		bDone = False
        '	End If
        '      End While
        If oNewEntity.LocX < Me.lMinPosX Then oNewEntity.LocX = Me.lMinPosX
        If oNewEntity.LocX > Me.lMaxPosX Then oNewEntity.LocX = Me.lMaxPosX
		If oNewEntity.LocZ < Me.lMinPosZ Then oNewEntity.LocZ = Me.lMinPosZ
        If oNewEntity.LocZ > Me.lMaxPosZ Then oNewEntity.LocZ = Me.lMaxPosZ

        If oNewEntity.DestX > Me.lMaxPosX Then oNewEntity.DestX = Me.lMaxPosX
        If oNewEntity.DestX < Me.lMinPosX Then oNewEntity.DestX = Me.lMinPosX

		'MSC - 08/08/07 - this may have caused issues... made it exactly like the EngineCode.HandleMovement version
		'  Left changed lines as remarked out

		lTemp = oNewEntity.LocZ + lHalfEnvirSize
		'lGridID = CInt(Math.Floor(lTemp / lGridSquareSize)) * lGridsPerRow
		lGridID = (lTemp \ lGridSquareSize) * lGridsPerRow
		'lTemp -= CInt((lGridID / lGridsPerRow) * lGridSquareSize)
		lTemp -= ((lGridID \ lGridsPerRow) * lGridSquareSize)
		'oNewEntity.lSmallSectorID = CInt(Math.Floor(lTemp / gl_SMALL_GRID_SQUARE_SIZE)) * lSmallPerRow
		oNewEntity.lSmallSectorID = ((lTemp \ gl_SMALL_GRID_SQUARE_SIZE) * lSmallPerRow)
		'lTemp -= CInt((oNewEntity.lSmallSectorID / lSmallPerRow) * gl_SMALL_GRID_SQUARE_SIZE) 
		lTemp -= ((oNewEntity.lSmallSectorID \ lSmallPerRow) * gl_SMALL_GRID_SQUARE_SIZE)
		'oNewEntity.lTinyZ = CInt(Math.Floor(lTemp / gl_FINAL_GRID_SQUARE_SIZE))
		oNewEntity.lTinyZ = lTemp \ gl_FINAL_GRID_SQUARE_SIZE

		lTemp2 = oNewEntity.LocX + lHalfEnvirSize
		'lTemp = CInt(Math.Floor(lTemp2 / lGridSquareSize))
		lTemp = lTemp2 \ lGridSquareSize
		lGridID += lTemp
		lTemp2 -= (lTemp * lGridSquareSize)
		'lTemp = CInt(Math.Floor(lTemp2 / gl_SMALL_GRID_SQUARE_SIZE))
		lTemp = lTemp2 \ gl_SMALL_GRID_SQUARE_SIZE
		oNewEntity.lSmallSectorID += lTemp
		lTemp2 -= (lTemp * gl_SMALL_GRID_SQUARE_SIZE)
		'oNewEntity.lTinyX = CInt(Math.Floor(lTemp2 / gl_FINAL_GRID_SQUARE_SIZE))
		oNewEntity.lTinyX = lTemp2 \ gl_FINAL_GRID_SQUARE_SIZE

		'Now that we have the grid id, add the entity
		oNewEntity.lGridEntityIdx = oGrid(lGridID).AddEntity(oNewEntity.ServerIndex)
		oNewEntity.lGridIndex = lGridID

		'now, check for new players
		lIdx = -1
		bFound = False
		For X = 0 To lPlayersWhoHaveUnitsHereUB
			If lPlayersWhoHaveUnitsHereIdx(X) = oNewEntity.Owner.ObjectID Then
				bFound = True
				lIdx = X
				Exit For
			ElseIf lIdx = -1 AndAlso lPlayersWhoHaveUnitsHereIdx(X) = -1 Then
				lIdx = X
			End If
		Next X

		If bFound = False Then
			If lIdx = -1 Then
				lPlayersWhoHaveUnitsHereUB += 1
				ReDim Preserve oPlayersWhoHaveUnitsHere(lPlayersWhoHaveUnitsHereUB)
				ReDim Preserve lPlayersWhoHaveUnitsHereIdx(lPlayersWhoHaveUnitsHereUB)
				ReDim Preserve lPlayersWhoHaveUnitsHereCP(lPlayersWhoHaveUnitsHereUB)
				lIdx = lPlayersWhoHaveUnitsHereUB
			End If
			lPlayersWhoHaveUnitsHereIdx(lIdx) = oNewEntity.Owner.ObjectID
			oPlayersWhoHaveUnitsHere(lIdx) = oNewEntity.Owner

			If PotentialAggression = False Then TestForAggression()
			If PotentialFirstContact = False Then TestForFirstContact()
		End If

        'By now, lIdx should be SOMETHING
        If oNewEntity.ObjTypeID = ObjectType.eUnit Then lPlayersWhoHaveUnitsHereCP(lIdx) += oNewEntity.CPUsage + oNewEntity.Owner.BadWarDecCPIncrease

	End Sub

    Public Sub RemoveEntity(ByVal lEntityIdx As Int32, ByVal yRemovalType As RemovalType, ByVal bSendClientsUpdate As Boolean, ByVal bSendToPrimary As Boolean, ByVal lKilledByID As Int32)
        'Dim oTmpUnit As Epica_Entity

        'Set the grid square entity index
        If goEntity(lEntityIdx) Is Nothing = False Then
            Dim oEntity As Epica_Entity = goEntity(lEntityIdx)

            If oEntity Is Nothing = True Then Return
            With oEntity

                RemoveLookupEntity(.ObjectID, .ObjTypeID)

#If EXTENSIVELOGGING = 1 Then
                Try
                    gfrmDisplayForm.AddEventLine("Envir.RemoveEntity: " & .ObjectID & ", " & .ObjTypeID & " from " & .ParentEnvir.ObjectID & ", " & .ParentEnvir.ObjTypeID & " at " & .LocX & ", " & .LocZ & " for " & yRemovalType.ToString)
                Catch
                End Try
#End If

                Dim lTryCount As Int32 = 0
                Dim oEnvirGrid As EnvirGrid = Nothing
                While oEnvirGrid Is Nothing
                    oEnvirGrid = oGrid(.lGridIndex)
                    lTryCount += 1
                    If lTryCount > 100 Then Exit While
                    If oEnvirGrid Is Nothing Then
                        Threading.Thread.Sleep(1)
                    End If
                End While
                'Just assume oEnvirGrid is something
                lTryCount = 0
                While oEnvirGrid.lEntities Is Nothing OrElse oEnvirGrid.lEntities.GetUpperBound(0) < .lGridEntityIdx
                    lTryCount += 1
                    If lTryCount > 100 Then Exit While
                    Threading.Thread.Sleep(1)
                End While

                oGrid(.lGridIndex).lEntities(.lGridEntityIdx) = -1
                TestForPlayerWithUnitsHere()

                For X As Int32 = 0 To lPlayersWhoHaveUnitsHereUB
                    If lPlayersWhoHaveUnitsHereIdx(X) = .Owner.ObjectID Then
                        lPlayersWhoHaveUnitsHereCP(X) -= (.CPUsage + .Owner.BadWarDecCPIncrease)
                        Exit For
                    End If
                Next X

                ''Remove targets and Targeted Bys
                .ClearTargetLists(False, -1)
                'Dim lTmpTargetByUB As Int32 = -1
                'If .lTargetedByIdx Is Nothing = False Then
                '	lTmpTargetByUB = Math.Min(.lTargetedByUB, .lTargetedByIdx.GetUpperBound(0))
                'End If
                'For lTemp As Int32 = 0 To lTmpTargetByUB
                '	If .lTargetedByIdx(lTemp) <> -1 Then
                '		If glEntityIdx(.lTargetedByIdx(lTemp)) <> -1 Then
                '			oTmpUnit = goEntity(.lTargetedByIdx(lTemp))
                '			If oTmpUnit Is Nothing = False Then
                '				If oTmpUnit.lPrimaryTargetServerIdx = .ServerIndex Then
                '					'If (oTmpUnit.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then oTmpUnit.CurrentStatus -= elUnitStatus.eUnitEngaged
                '					If (oTmpUnit.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eUnitEngaged
                '					oTmpUnit.lPrimaryTargetServerIdx = -1
                '					oTmpUnit.bForceAggressionTest = True
                '					SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 
                '				End If
                '				If oTmpUnit.lTargetsServerIdx(0) = .ServerIndex Then
                '					oTmpUnit.lTargetsServerIdx(0) = -1
                '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0 Then
                '						'oTmpUnit.CurrentStatus -= elUnitStatus.eSide1HasTarget
                '						oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide1HasTarget
                '					End If
                '					oTmpUnit.bForceAggressionTest = True
                '					SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                '				End If
                '				If oTmpUnit.lTargetsServerIdx(1) = .ServerIndex Then
                '					oTmpUnit.lTargetsServerIdx(1) = -1
                '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0 Then
                '						'oTmpUnit.CurrentStatus -= elUnitStatus.eSide2HasTarget
                '						oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide2HasTarget
                '					End If
                '					oTmpUnit.bForceAggressionTest = True
                '					SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                '				End If
                '				If oTmpUnit.lTargetsServerIdx(2) = .ServerIndex Then
                '					oTmpUnit.lTargetsServerIdx(2) = -1
                '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0 Then
                '						'oTmpUnit.CurrentStatus -= elUnitStatus.eSide3HasTarget
                '						oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide3HasTarget
                '					End If
                '					oTmpUnit.bForceAggressionTest = True
                '					SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                '				End If
                '				If oTmpUnit.lTargetsServerIdx(3) = .ServerIndex Then
                '					oTmpUnit.lTargetsServerIdx(3) = -1
                '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0 Then
                '						'oTmpUnit.CurrentStatus -= elUnitStatus.eSide4HasTarget
                '						oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Xor elUnitStatus.eSide4HasTarget
                '					End If
                '					oTmpUnit.bForceAggressionTest = True
                '					SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                '				End If
                '			End If
                '		End If
                '	End If
                'Next lTemp

                If bSendClientsUpdate = True OrElse bSendToPrimary = True Then
                    Dim yOutMsg() As Byte = goMsgSys.CreateRemoveObjectMsg(.ObjectID, .ObjTypeID, yRemovalType, .LocX, .LocZ, .CurrentStatus, lKilledByID)
                    If bSendClientsUpdate = True AndAlso lPlayersInEnvirCnt > 0 Then goMsgSys.BroadcastToEnvironmentClients(yOutMsg, .ParentEnvir)
                    If bSendToPrimary = True Then goMsgSys.SendToPrimary(yOutMsg)
                End If
            End With
        End If
        'Remove the entity from our global array
        glEntityIdx(lEntityIdx) = -1
        goEntity(lEntityIdx) = Nothing
    End Sub

	Public Sub TestForAggression()
		Dim X As Int32
		Dim Y As Int32
		Dim yRelID As Byte

		'Now, check rels for possible aggression if we don't already have it
		PotentialAggression = False

		For X = 0 To lPlayersWhoHaveUnitsHereUB
			If lPlayersWhoHaveUnitsHereIdx(X) <> -1 Then

				If lPlayersWhoHaveUnitsHereIdx(X) = gl_HARDCODE_PIRATE_PLAYER_ID Then
					PotentialAggression = True
					Return
				End If

				Dim oPlayer As Player = oPlayersWhoHaveUnitsHere(X)
				If oPlayer Is Nothing = False Then

					For Y = 0 To lPlayersWhoHaveUnitsHereUB
						If lPlayersWhoHaveUnitsHereIdx(Y) <> -1 AndAlso X <> Y Then
							'yRelID = oPlayersWhoHaveUnitsHere(X).GetPlayerRelScore(lPlayersWhoHaveUnitsHereIdx(Y))
							yRelID = oPlayer.GetPlayerRelScore(lPlayersWhoHaveUnitsHereIdx(Y), False, -1)
							If yRelID <= elRelTypes.eWar Then
								PotentialAggression = True
								Return
							End If
						End If
					Next Y

				End If
			End If
		Next X

	End Sub

	Public Sub TestForFirstContact()
		Dim X As Int32
		Dim Y As Int32

		'Now, check rels for possible aggression if we don't already have it
		PotentialFirstContact = False

		For X = 0 To lPlayersWhoHaveUnitsHereUB
			If lPlayersWhoHaveUnitsHereIdx(X) <> -1 Then

				If lPlayersWhoHaveUnitsHereIdx(X) = gl_HARDCODE_PIRATE_PLAYER_ID Then Continue For

				Dim oPlayer As Player = oPlayersWhoHaveUnitsHere(X)
				If oPlayer Is Nothing = False Then

					For Y = 0 To lPlayersWhoHaveUnitsHereUB
						If lPlayersWhoHaveUnitsHereIdx(Y) <> -1 AndAlso X <> Y Then
							If oPlayer.HasPlayerRelationship(lPlayersWhoHaveUnitsHereIdx(Y)) = False Then
								PotentialFirstContact = True
								Return
							End If
						End If
					Next Y
				End If
			End If
		Next X
	End Sub

	Protected Overrides Sub Finalize()
		oGeoObject = Nothing
		MyBase.Finalize()
	End Sub

	Public Sub SetEnvirGridValues()
		'Dim X As Int32

		'NOTE: this assumes that the oGeoObject is set correctly... it should be either a system or a planet
		If oGeoObject Is Nothing = False Then
			If CType(oGeoObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
				'Planet object
				lEnvirSize = CType(oGeoObject, Planet).GetCellSpacing() * TerrainClass.Width
				'lGridSquareSize = 1600
				lMapWrapAdjustX = CType(oGeoObject, Planet).GetCellSpacing * (TerrainClass.Width - 1)
			Else
				'System Object
				lEnvirSize = 10000000
				'lGridSquareSize = 8000
			End If
			lGridSquareSize = 8000

			'lHalfEnvirSize = CInt(lEnvirSize / 2)
			lHalfEnvirSize = lEnvirSize \ 2
			'lGridsPerRow = CInt(lEnvirSize / lGridSquareSize)
			lGridsPerRow = lEnvirSize \ lGridSquareSize

			lGridUB = (lGridsPerRow * lGridsPerRow) - 1

			'set our extents
			lMinPosX = -lHalfEnvirSize + 1
			lMinPosZ = -lHalfEnvirSize + 1
			lMaxPosX = lHalfEnvirSize - 1
			lMaxPosZ = lHalfEnvirSize - 1

			'Initialize our grid array
            'ReDim moGrid(lGridUB)

			ReDim lGridIdxAdjust(8)
			lGridIdxAdjust(0) = 0
			lGridIdxAdjust(1) = -lGridsPerRow - 1
			lGridIdxAdjust(2) = -lGridsPerRow
			lGridIdxAdjust(3) = -lGridsPerRow + 1
			lGridIdxAdjust(4) = -1
			lGridIdxAdjust(5) = 1
			lGridIdxAdjust(6) = lGridsPerRow - 1
			lGridIdxAdjust(7) = lGridsPerRow
			lGridIdxAdjust(8) = lGridsPerRow + 1

			'I am on the left edge of the map... (only for planets)
            ReDim lLeftEdgeGridIdxAdjust(5) '8
			lLeftEdgeGridIdxAdjust(0) = 0
            'lLeftEdgeGridIdxAdjust(1) = -1					'upleft
            lLeftEdgeGridIdxAdjust(1) = -lGridsPerRow       'up
            lLeftEdgeGridIdxAdjust(2) = -lGridsPerRow + 1   'upright
            'lLeftEdgeGridIdxAdjust(4) = lGridsPerRow - 1	'left
            lLeftEdgeGridIdxAdjust(3) = 1                   'right
            'lLeftEdgeGridIdxAdjust(6) = (lGridsPerRow * 2) - 1	'lower left
            lLeftEdgeGridIdxAdjust(4) = lGridsPerRow        'down
            lLeftEdgeGridIdxAdjust(5) = lGridsPerRow + 1    'down right

			'use this on the right edge of the map (planets only)
            ReDim lRightEdgeGridIdxAdjust(5)    '8
			lRightEdgeGridIdxAdjust(0) = 0
			lRightEdgeGridIdxAdjust(1) = -lGridsPerRow - 1			 'upleft
			lRightEdgeGridIdxAdjust(2) = -lGridsPerRow				 'up
            'lRightEdgeGridIdxAdjust(3) = -(lGridsPerRow * 2 - 1)	 'upright
            lRightEdgeGridIdxAdjust(3) = -1                          'left
            'lRightEdgeGridIdxAdjust(5) = -(lGridsPerRow - 1)		 'right
            lRightEdgeGridIdxAdjust(4) = lGridsPerRow - 1            'lower left
            lRightEdgeGridIdxAdjust(5) = lGridsPerRow                'down
            'lRightEdgeGridIdxAdjust(8) = 1							 'down right

		End If
	End Sub

	Public Sub TestForPlayerWithUnitsHere()
        DoLookupGridFunction(elLookupGridType.TestForPlayerWithUnitsHere, 0)
    End Sub

	Public Function GetEnvirSpeedMod(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iAngle As Int16) As Single
		If Me.ObjTypeID = ObjectType.ePlanet Then
			If oGeoObject Is Nothing = False Then
				Return CType(oGeoObject, Planet).GetSpeedMod(lLocX, lLocZ, iAngle)
			Else : Return 1.0F
			End If
		Else
			'Ok, must be solar system

			'TODO: Here is where you would put the nebula speed modifier, for example return 0.8

			Return 1.0F
		End If
	End Function

	Public Sub DoCommandPointUpdates()
		For Y As Int32 = 0 To lPlayersInEnvirUB
			If lPlayersInEnvirIdx(Y) > 0 Then

				Dim lTmpPlayerID As Int32 = oPlayersInEnvir(Y).ObjectID
				If oPlayersInEnvir(Y).lAliasingPlayerID <> -1 Then lTmpPlayerID = oPlayersInEnvir(Y).lAliasingPlayerID

				'Ok, let's send the update
				For Z As Int32 = 0 To lPlayersWhoHaveUnitsHereUB
					If lPlayersWhoHaveUnitsHereIdx(Z) = lTmpPlayerID Then ' lPlayersInEnvirIdx(Y) Then
                        If oPlayersInEnvir(Y).oSocket Is Nothing = False Then
                            Dim yMsg(11) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateCommandPoints).CopyTo(yMsg, 0)
                            GetGUIDAsString.CopyTo(yMsg, 2)
                            System.BitConverter.GetBytes(lPlayersWhoHaveUnitsHereCP(Z)).CopyTo(yMsg, 8)

                            If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eUpdateCommandPoints, MsgMonitor.eMM_AppType.ClientConnection, yMsg.Length, lPlayersInEnvirIdx(Y))
                            oPlayersInEnvir(Y).oSocket.SendData(yMsg)

                            'Else
                            'lPlayersInEnvirIdx(Y) = -1
                            'lPlayersInEnvirCnt -= 1
                        End If

						Exit For
					End If
				Next Z

			End If
		Next Y
    End Sub

    Public Sub ResetAllPlayerCP()
        For X As Int32 = 0 To lPlayersWhoHaveUnitsHereUB
            Try
                If lPlayersWhoHaveUnitsHereIdx(X) > -1 Then
                    lPlayersWhoHaveUnitsHereCP(X) = 0
                    Exit For
                End If
            Catch
            End Try
        Next X
    End Sub

    Public Sub AdjustPlayerCommandPoints(ByVal lPlayerID As Int32, ByVal lByAmt As Int32)
        'Reset our CP Usages
        For X As Int32 = 0 To lPlayersWhoHaveUnitsHereUB
            If lPlayersWhoHaveUnitsHereIdx(X) = lPlayerID Then
                lPlayersWhoHaveUnitsHereCP(X) += lByAmt
                Exit For
            End If
        Next X
    End Sub

	Public Sub ClearEnvirVariables()
		On Error Resume Next
		oGeoObject = Nothing
		oPlayersWhoHaveUnitsHere = Nothing
		lPlayersWhoHaveUnitsHereIdx = Nothing
		lPlayersWhoHaveUnitsHereCP = Nothing
		oPlayersInEnvir = Nothing
		lPlayersInEnvirIdx = Nothing

        If mcolEnvirGrid Is Nothing = False Then mcolEnvirGrid.Clear()
        mcolEnvirGrid = Nothing

		lGridIdxAdjust = Nothing
		lLeftEdgeGridIdxAdjust = Nothing
		lRightEdgeGridIdxAdjust = Nothing
		oCache = Nothing
		lCacheIdx = Nothing
		myFacBurst = Nothing
		myCacheBurst = Nothing
	End Sub

	Public Function PlaceMineralCache() As System.Drawing.Point
        If Me.ObjTypeID = ObjectType.ePlanet Then
            Dim oRnd As New System.Random()

            With CType(oGeoObject, Planet)
                Dim lCnt As Int32 = 0
                Dim bDone As Boolean = False
                Dim lWaterHeight As Int32 = .WaterHeight
                Dim lTmpExt As Int32 = CInt(lEnvirSize * 0.95F)
                Dim lHlfTmpExt As Int32 = lTmpExt \ 2
                While bDone = False
                    lCnt += 1
                    Dim lX As Int32 = oRnd.Next(-lHlfTmpExt, lHlfTmpExt)
                    'Dim lX As Int32 = CInt(Rnd() * lTmpExt) - lHlfTmpExt
                    'Dim lZ As Int32 = CInt(Rnd() * lTmpExt) - lHlfTmpExt
                    Dim lZ As Int32 = oRnd.Next(-lHlfTmpExt, lHlfTmpExt)
                    If lCnt > 1000 Then Return New Point(lX, lZ)

                    'mineral caches cannot be placed in the water on lava and acid planets...
                    If .MapTypeID = PlanetType.eGeoPlastic OrElse .MapTypeID = PlanetType.eAcidic Then
                        Dim lHt As Int32 = .GetHeightAtPoint(lX, lZ, False)
                        If lHt < lWaterHeight Then Continue While
                    End If

                    If .TerrainGradeBuildable(lX, lZ) = True Then
                        Return New Point(lX, lZ)
                    End If
                End While
            End With
        Else
            Return New Point(CInt(Rnd() * 10000000) - 5000000, CInt(Rnd() * 10000000) - 5000000)
        End If
	End Function

    'Public Sub FillGridUsageValues(ByRef lTotalGrids As Int32, ByRef lUnusedGrids As Int32, ByRef lUsedGrids As Int32, ByRef lInUseGrids As Int32, ByRef lOldGrids As Int32)
    '	lTotalGrids += (moGrid.GetUpperBound(0) + 1)
    '	For Y As Int32 = 0 To moGrid.GetUpperBound(0)
    '		If moGrid(Y) Is Nothing = False Then
    '			If moGrid(Y).lEntityUB = -1 Then
    '				lUnusedGrids += 1
    '			Else
    '				lUsedGrids += 1

    '				Dim bFound As Boolean = False
    '				Dim lCurUB As Int32 = -1
    '				If moGrid(Y).lEntities Is Nothing = False Then
    '					lCurUB = Math.Min(moGrid(Y).lEntityUB, moGrid(Y).lEntities.GetUpperBound(0))
    '				End If
    '				For Z As Int32 = 0 To lCurUB
    '					If moGrid(Y).lEntities(Z) > -1 Then
    '						lInUseGrids += 1
    '						bFound = True
    '						Exit For
    '					End If
    '				Next Z
    '				If bFound = False Then lOldGrids += 1
    '			End If
    '		Else : lTotalGrids -= 1
    '		End If
    '	Next Y
    '   End Sub

    Public Sub ClearPlayersCP(ByVal lPlayerID As Int32)
        For X As Int32 = 0 To lPlayersWhoHaveUnitsHereUB
            Try
                If lPlayersWhoHaveUnitsHereIdx(X) = lPlayerID Then
                    lPlayersWhoHaveUnitsHereCP(X) = 0
                    Exit For
                End If
            Catch
            End Try
        Next X
        
    End Sub

    Private mlLockMe(-1) As Int32
    'Public Sub AddPlayerToEnvir(ByRef oPlayer As Player)

    '    If mlLockMe Is Nothing Then ReDim mlLockMe(-1)

    '    SyncLock mlLockMe
    '        Dim bFound As Boolean = False
    '        Dim lIdx As Int32 = -1
    '        For Y As Int32 = 0 To lPlayersInEnvirUB
    '            If lPlayersInEnvirIdx(Y) = oPlayer.ObjectID Then
    '                If oPlayersInEnvir(Y) Is Nothing Then oPlayersInEnvir(Y) = oPlayer
    '                Return
    '            ElseIf lIdx = -1 AndAlso lPlayersInEnvirIdx(Y) = -1 Then
    '                lIdx = Y
    '            End If
    '        Next Y

    '        If bFound = False Then
    '            If lIdx = -1 Then
    '                lPlayersInEnvirUB += 1
    '                lIdx = lPlayersInEnvirUB
    '                ReDim Preserve lPlayersInEnvirIdx(lIdx)
    '                ReDim Preserve oPlayersInEnvir(lIdx)
    '            End If
    '            lPlayersInEnvirIdx(lIdx) = oPlayer.ObjectID
    '            oPlayersInEnvir(lIdx) = oPlayer
    '        End If
    '    End SyncLock
    'End Sub
    Private Shared mbDoPlayerInEnvirCntDec As Boolean = False
    Public Sub DoEnvirPlayerChange(ByVal yType As eyChangePlayerEnvirCode, ByRef oPlayer As Player)
        If mlLockMe Is Nothing Then ReDim mlLockMe(-1)
        SyncLock mlLockMe
            If yType = eyChangePlayerEnvirCode.eRemoveFromEnvir Then       'remove
                Dim lOtherUB As Int32 = -1
                Dim lPlayerID As Int32 = oPlayer.ObjectID
                If lPlayersInEnvirIdx Is Nothing = False Then lOtherUB = Math.Min(lPlayersInEnvirUB, lPlayersInEnvirIdx.GetUpperBound(0))
                For Y As Int32 = 0 To lOtherUB
                    If lPlayersInEnvirIdx(Y) = lPlayerID Then
                        lPlayersInEnvirIdx(Y) = -1
                        oPlayersInEnvir(Y) = Nothing

                        If mbDoPlayerInEnvirCntDec = True Then
                            lPlayersInEnvirCnt -= 1
                            If lPlayersInEnvirCnt < 0 Then lPlayersInEnvirCnt = 0
                        End If
                    End If
                Next Y
            Else                'add

                'Dim bFound As Boolean = False
                Dim lIdx As Int32 = -1

                For Y As Int32 = 0 To lPlayersInEnvirUB
                    If lPlayersInEnvirIdx(Y) = oPlayer.ObjectID Then
                        'bFound = True
                        If oPlayersInEnvir(Y) Is Nothing Then oPlayersInEnvir(Y) = oPlayer
                        Return
                    ElseIf lIdx = -1 AndAlso lPlayersInEnvirIdx(Y) = -1 Then
                        lIdx = Y
                    End If
                Next Y

                'If bFound = False Then
                If lIdx = -1 Then
                    lPlayersInEnvirUB += 1
                    lIdx = lPlayersInEnvirUB
                    ReDim Preserve lPlayersInEnvirIdx(lIdx)
                    ReDim Preserve oPlayersInEnvir(lIdx)
                End If
                lPlayersInEnvirIdx(lIdx) = oPlayer.ObjectID
                oPlayersInEnvir(lIdx) = oPlayer
                'End If

            End If
        End SyncLock

    End Sub

    Public Function GetPlayerInEnvironment(ByVal lPlayerID As Int32) As Player
        Dim lUB As Int32 = -1
        If oPlayersInEnvir Is Nothing = False Then lUB = Math.Min(lPlayersInEnvirUB, oPlayersInEnvir.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If lPlayersInEnvirIdx(X) = lPlayerID AndAlso oPlayersInEnvir(X) Is Nothing = False Then Return oPlayersInEnvir(X)
        Next X
        Return Nothing
    End Function

    Private Shared mbDoCheckForGridExpiration As Boolean = True
    Public Function CheckForGridExpiration() As Boolean
        If mcolEnvirGrid Is Nothing Then Return False
        If mbDoCheckForGridExpiration = False Then Return False

        Try
            Dim oGrid As EnvirGrid = DoLookupGridFunction(elLookupGridType.CheckForExpiration, 0)
            If oGrid Is Nothing = False Then
                Return True
            End If
        Catch
            Try
                DoLookupGridFunction(elLookupGridType.ClearAndRebuildGrid, 0)
            Catch
            End Try
        End Try

        Return False
    End Function

#Region " War Trackers "
	Public oWars() As WarTracker
	Public lWarIdx() As Int32		'the player ID
	Public lWarUB As Int32 = -1

	Public Function GetOrAddWarTracker(ByVal lPlayerID As Int32) As WarTracker
		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To lWarUB
			If lWarIdx(X) = lPlayerID Then
				Return oWars(X)
			ElseIf lIdx = -1 AndAlso lWarIdx(X) = -1 Then
				lIdx = X
			End If
		Next X

		If lIdx = -1 Then
			lWarUB += 1
			ReDim Preserve oWars(lWarUB)
			ReDim Preserve lWarIdx(lWarUB)
			lIdx = lWarUB
		End If
		oWars(lIdx) = New WarTracker
		lWarIdx(lIdx) = lPlayerID
		oWars(lIdx).lPlayerID = lPlayerID
		oWars(lIdx).lStart = GetDateAsNumber(Now)
		oWars(lIdx).lPreviousMsgSend = glCurrentCycle
		Return oWars(lIdx)
	End Function
#End Region
End Class

Public Class EnvirGrid
    'Ok... we contain indices to units that reside here...
	Public lEntities() As Int32
	Public lEntityUB As Int32 = -1

    Public lMissiles() As Int32
    Public lMissileUB As Int32 = -1

    Public lGridID As Int32

    Public lExpirationCycle As Int32 = Int32.MaxValue

    Public Function AddEntity(ByVal lEntityIdx As Int32) As Int32  'returns the index
        Dim X As Int32

        Dim lIdx As Int32 = -1

        lExpirationCycle = Int32.MaxValue

        For X = 0 To lEntityUB
            If lEntities(X) = lEntityIdx Then
                Return X
            ElseIf lIdx = -1 AndAlso lEntities(X) = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx <> -1 Then
            lEntities(lIdx) = lEntityIdx
            Return lIdx
        End If

        lEntityUB += 1
        ReDim Preserve lEntities(lEntityUB)
        lEntities(lEntityUB) = lEntityIdx
        Return lEntityUB
    End Function

    Public Sub RemoveEntity(ByVal lIndex As Int32)
        'index is the INDEX in the entity array, not the ID of the entity
        lEntities(lIndex) = -1
        lExpirationCycle = glCurrentCycle + 150
    End Sub

    Public Function AddMissile(ByVal lItemIdx As Int32) As Int32
        Dim lIdx As Int32 = -1

        lExpirationCycle = Int32.MaxValue

        For X As Int32 = 0 To lMissileUB
            If lMissiles(X) = lItemIdx Then
                Return X
            ElseIf lIdx = -1 AndAlso lMissiles(X) = -1 Then
                lIdx = X
            End If
        Next X

        If lIdx <> -1 Then
            lMissiles(lIdx) = lItemIdx
            Return lIdx
        End If

        lMissileUB += 1
        ReDim Preserve lMissiles(lMissileUB)
        lMissiles(lMissileUB) = lItemIdx
        Return lMissileUB
    End Function

    Public Sub RemoveMissile(ByVal lArrayIdx As Int32)
        lExpirationCycle = glCurrentCycle + 150
        If lMissiles Is Nothing Then Return
        If lArrayIdx > -1 AndAlso lArrayIdx <= lMissiles.GetUpperBound(0) Then lMissiles(lArrayIdx) = -1
    End Sub

    Public Function CheckForExpiration() As Boolean
        Try
            For X As Int32 = 0 To lEntityUB
                If lEntities(X) > -1 Then Return False
            Next X
            For X As Int32 = 0 To lMissileUB
                If lMissiles(X) > -1 Then Return False
            Next X
            Return True
        Catch
        End Try
        Return False
    End Function
End Class

Public Class StarType
    Public StarTypeID As Byte
    Public StarTypeAttrs As Int32       'bit-wise attributes
    Public StarRadius As Int32
    Public HeatIndex As Byte

End Class

Public Class Galaxy
    Inherits Epica_GUID

    Public GalaxyName As String

    Public moSystems() As SolarSystem
    Public mlSystemUB As Int32 = -1
    'TODO: Include Nebulae

    Public CurrentSystemIdx As Int32 = -1

    Public Sub AddSystem(ByVal oSystem As SolarSystem)
        mlSystemUB += 1
        ReDim Preserve moSystems(mlSystemUB)
        moSystems(mlSystemUB) = oSystem
    End Sub
End Class

Public Class SolarSystem
    Inherits Epica_GUID

    Public SystemName As String
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32

    Public EnvirIdx As Int32 = -1       'index in the goEnvirs array

    Public StarType1Idx As Int32 = -1   'index in the StarType array
    Public StarType2Idx As Int32 = -1   'index in the StarType array
    Public StarType3Idx As Int32 = -1   'index in the StarType array

    Public SystemType As Byte = 0

    Public FleetJumpPointX As Int32
    Public FleetJumpPointZ As Int32

    Public moPlanets() As Planet
    Public PlanetUB As Int32 = -1

    Public moWormholes() As Wormhole
    Public mlWHX() As Int32
    Public mlWHZ() As Int32
    Public WormholeUB As Int32 = -1

    Public Sub AddPlanet(ByVal oPlanet As Planet)
        PlanetUB += 1
        ReDim Preserve moPlanets(PlanetUB)
        moPlanets(PlanetUB) = oPlanet
    End Sub

    Public Sub AddWormhole(ByRef oWormhole As Wormhole)
        For X As Int32 = 0 To WormholeUB
            If moWormholes(X) Is Nothing = False AndAlso moWormholes(X).ObjectID = oWormhole.ObjectID Then
                moWormholes(X) = oWormhole
                Return
            End If
        Next X
        WormholeUB += 1
        ReDim Preserve moWormholes(WormholeUB)
        ReDim Preserve mlWHX(WormholeUB)
        ReDim Preserve mlWHZ(WormholeUB)
        moWormholes(WormholeUB) = oWormhole
        If oWormhole.System1 Is Nothing = False AndAlso oWormhole.System1.ObjectID = Me.ObjectID Then
            mlWHX(WormholeUB) = oWormhole.LocX1
            mlWHZ(WormholeUB) = oWormhole.LocY1
        Else
            mlWHX(WormholeUB) = oWormhole.LocX2
            mlWHZ(WormholeUB) = oWormhole.LocY2
        End If
    End Sub

    'Private Const mf_Wormhole_Detect_Radius As Single = (gl_FINAL_GRID_SQUARE_SIZE * 128) * (gl_FINAL_GRID_SQUARE_SIZE * 128)
    'Public Function RoutePassesWormhole(ByVal lSX As Int32, ByVal lSZ As Int32, ByVal lEX As Int32, ByVal lEZ As Int32) As Boolean
    '    For X As Int32 = 0 To WormholeUB
    '        If moWormholes(X) Is Nothing = False Then
    '            With moWormholes(X)
    '                Dim lWX As Int32
    '                Dim lWZ As Int32
    '                If .System1.ObjectID = Me.ObjectID Then
    '                    lWX = .LocX1
    '                    lWZ = .LocY1
    '                Else
    '                    lWX = .LocX2
    '                    lWZ = .LocY2
    '                End If
    '                Dim lX1 As Int32 = lSX - lWX
    '                Dim lY1 As Int32 = lSZ - lWZ
    '                Dim lX2 As Int32 = lEX - lWX
    '                Dim lY2 As Int32 = lEZ - lWZ

    '                lWX = lX2 - lX1
    '                lWZ = lY2 - lY1
    '                Dim fX As Single = lWX * lWX
    '                Dim fZ As Single = lWZ * lWZ
    '                Dim dDist As Double = Math.Sqrt(fX + fZ)
    '                Dim fD As Single = (lX1 * lY2) - (lX2 * lY1)
    '                Dim dDiscriminant As Double = mf_Wormhole_Detect_Radius * (dDist * dDist) - (fD * fD)
    '                If dDiscriminant >= 0 Then Return True
    '            End With
    '        End If
    '    Next X
    '    Return False
    'End Function

    'We square the proximity Test dist to remove the need to sqrt our distance value
	Private Const mf_ProximityTestDist As Single = ((gl_FINAL_GRID_SQUARE_SIZE * 255) * (gl_FINAL_GRID_SQUARE_SIZE * 255)) * 10
    Public Sub HandleWormholeProximityTest(ByRef oEntity As Epica_Entity)
        For X As Int32 = 0 To WormholeUB
            Try
                With moWormholes(X)

                    If (Me.ObjectID = .System1.ObjectID AndAlso (.WormholeFlags And elWormholeFlag.eSystem1Detectable) <> 0) OrElse (Me.ObjectID = .System2.ObjectID AndAlso (.WormholeFlags And elWormholeFlag.eSystem2Detectable) <> 0) Then
                        Dim lWHX As Int32 = oEntity.LocX - mlWHX(X)
                        Dim lWHZ As Int32 = oEntity.LocZ - mlWHZ(X)
                        If Math.Abs(lWHX) > 100000 OrElse Math.Abs(lWHZ) > 100000 Then Continue For
                        Dim fX As Double = CDbl(lWHX) * CDbl(lWHX)
                        Dim fZ As Double = CDbl(lWHZ) * CDbl(lWHZ)
                        Dim fDist As Double = fX + fZ
                        If fDist < mf_ProximityTestDist Then
                            oEntity.Owner.HandleCheckFirstContactWithWormhole(moWormholes(X), Me.ObjectID)
                            'moWormholes(X).StartCycle = 1

                            'If Me.ObjectID = .System1.ObjectID Then
                            '    .WormholeFlags = .WormholeFlags Or elWormholeFlag.eSystem2Detectable
                            'Else
                            '    .WormholeFlags = .WormholeFlags Or elWormholeFlag.eSystem1Detectable
                            'End If
                        End If
                    End If

                End With
            Catch
            End Try
        Next X
    End Sub
End Class


Public Enum elWormholeFlag As int32
    eSystem1Detectable = 1
    eSystem2Detectable = 2
    eSystem1Jumpable = 4
    eSystem2Jumpable = 8
End Enum
Public Class Wormhole
    Inherits Epica_GUID

    Public System1 As SolarSystem
    Public System2 As SolarSystem
    Public LocX1 As Int32
    Public LocY1 As Int32
    Public LocX2 As Int32
    Public LocY2 As Int32

    Public StartCycle As Int32
    'Public EndCycle As Int32
    Public WormholeFlags As Int32

    'The worm hole is NEVER sent from the region server!!!
End Class

Public Class Planet
	Inherits Epica_GUID

	Public Shared moTutorialPlanet As Planet = Nothing


    Private Const ml_START_LOC_GRID_WH As Int32 = 12
    Private Const ml_MAX_START_LOC_FAC_COUNT As Int32 = 30

    Public PlanetName As String
    Public MapTypeID As Byte        'the planet typeid
    Public PlanetSizeID As Byte     'used for determining map size, 0-tiny, 1-small, 2-medium, 3-large, 4-huge. Maybe able to remove and base off of Radius
    Public PlanetRadius As Int16    'might be able to remove sizeID and base the map size off of radius
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32
    Public Vegetation As Byte
    Public Atmosphere As Byte
    Public Hydrosphere As Byte
    Public Gravity As Byte
    Public SurfaceTemperature As Byte
    Public RotationDelay As Int16   'cycles between incrementing rotation angle

    Public ParentSystem As SolarSystem

    Private moTerrain As TerrainClass
    'Private miSpeedMod(,,) As Int16
    Private mlFullSizeWH As Int32
    Private mlDoubleWH As Int32
    Private mf1OverCell As Single

    Public AxisAngle As Int32       'axis angle (yaw)
    Public RotateAngle As Int32     'rotation angle

    Public EnvirIdx As Int32 = -1       'index in the goEnvirs array

    Public myStartLocGrid() As Byte      'Indicates that the grid location is usable or not

    Public Sub New(ByVal lID As Int32, ByVal ySizeID As Byte, ByVal yMapTypeID As Byte)
        ObjectID = lID
        PlanetSizeID = ySizeID
        MapTypeID = yMapTypeID

        moTerrain = New TerrainClass(lID)
        If PlanetSizeID = 0 Then moTerrain.ml_Y_Mult = 9.0F
        moTerrain.MapType = yMapTypeID

        Select Case PlanetSizeID
            Case 0 : moTerrain.CellSpacing = gl_TINY_PLANET_CELL_SPACING
            Case 1 : moTerrain.CellSpacing = gl_SMALL_PLANET_CELL_SPACING
            Case 2 : moTerrain.CellSpacing = gl_MEDIUM_PLANET_CELL_SPACING
            Case 3 : moTerrain.CellSpacing = gl_LARGE_PLANET_CELL_SPACING  '540 '550
            Case 4 : moTerrain.CellSpacing = gl_HUGE_PLANET_CELL_SPACING  '740 '700
        End Select
    End Sub

    'NOTE: Changed Singles to Int32s
	Public Function GetHeightAtPoint(ByVal lX As Int32, ByVal lZ As Int32, ByVal bWaterHeightMin As Boolean) As Int32
		Return moTerrain.GetHeightAtLocation(lX, lZ, bWaterHeightMin)
	End Function

    Public Function HasLineOfSight(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32, ByVal lZ2 As Int32, ByVal lAttackerHeight As Int32) As Boolean
        Return moTerrain.HasLineOfSight(lX1, lY1, lZ1, lX2, lY2, lZ2, lAttackerHeight)
    End Function

    Public Sub PopulateData()
        'this sub takes the place of everything that I removed from Client, we need to do everything to get it ready...
        moTerrain.PopulateData()
        'FillSpeedModArray()
        mlFullSizeWH = moTerrain.CellSpacing * TerrainClass.Width      'NOTE: we assume Width = Height here
        mlDoubleWH = TerrainClass.Width * 2                            'NOTE: we assume Width = Height here
        mf1OverCell = 1.0F / moTerrain.CellSpacing

        SetupStartLocGrid()
	End Sub

	Public Sub PopulateInstanceData()
		'this sub takes the place of everything that I removed from Client, we need to do everything to get it ready...
		moTerrain.PopulateInstanceData()
        'FillSpeedModArray()
		mlFullSizeWH = moTerrain.CellSpacing * TerrainClass.Width	   'NOTE: we assume Width = Height here
        mlDoubleWH = TerrainClass.Width * 2                            'NOTE: we assume Width = Height here
        mf1OverCell = 1.0F / moTerrain.CellSpacing
	End Sub

    Public Function GetCellSpacing() As Int32
        Return moTerrain.CellSpacing
    End Function

    'Private Sub FillSpeedModArray()
    '    'Call this when terrain is ready...

    '    ReDim miSpeedMod((TerrainClass.Width * 2) - 1, (TerrainClass.Height * 2) - 1, 7)
    '    'Ok, now, fill it up
    '    Dim X As Int32
    '    Dim Y As Int32
    '    Dim lHalfCell As Int32 = CInt(moTerrain.CellSpacing / 2)

    '    Dim fTX As Single
    '    Dim fTY As Single

    '    Dim vTemp As Vector3

    '    'For each quad...
    '    For Y = 0 To (TerrainClass.Height * 2) - 1
    '        fTY = Y / 2.0F
    '        For X = 0 To (TerrainClass.Width * 2) - 1
    '            fTX = X / 2.0F

    '            'Ok, get our value
    '            vTemp = moTerrain.GetTerrainNormalEx(fTX, fTY)

    '            ''Right
    '            'miSpeedMod(X, Y, 0) = (1 + vTemp.X) * 100
    '            ''Up Right
    '            'miSpeedMod(X, Y, 1) = (1 + (vTemp.Z + vTemp.X)) * 100
    '            ''Up
    '            'miSpeedMod(X, Y, 2) = (1 + vTemp.Z) * 100
    '            ''Up Left
    '            'miSpeedMod(X, Y, 3) = (1 + (vTemp.Z + -vTemp.X)) * 100
    '            ''Left
    '            'miSpeedMod(X, Y, 4) = (1 + (-vTemp.X)) * 100
    '            ''Down
    '            'miSpeedMod(X, Y, 6) = (1 + (-vTemp.Z)) * 100
    '            ''Down Left
    '            'miSpeedMod(X, Y, 5) = (1 + (-vTemp.Z + -vTemp.X)) * 100
    '            ''Down Right
    '            'miSpeedMod(X, Y, 7) = (1 + (-vTemp.Z + vTemp.X)) * 100
    '            'Right
    '            miSpeedMod(X, Y, 0) = CShort((1 + vTemp.X) * 100)
    '            'Up Right
    '            miSpeedMod(X, Y, 7) = CShort((1 + (vTemp.Z + vTemp.X)) * 100)
    '            'Up
    '            miSpeedMod(X, Y, 6) = CShort((1 + vTemp.Z) * 100)
    '            'Up Left
    '            miSpeedMod(X, Y, 5) = CShort((1 + (vTemp.Z + -vTemp.X)) * 100)
    '            'Left
    '            miSpeedMod(X, Y, 4) = CShort((1 + (-vTemp.X)) * 100)
    '            'Down Left
    '            miSpeedMod(X, Y, 3) = CShort((1 + (-vTemp.Z + -vTemp.X)) * 100)
    '            'Down
    '            miSpeedMod(X, Y, 2) = CShort((1 + (-vTemp.Z)) * 100)
    '            'Down Right
    '            miSpeedMod(X, Y, 1) = CShort((1 + (-vTemp.Z + vTemp.X)) * 100)
    '        Next X
    '    Next Y
    'End Sub

    Public Function GetSpeedMod(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iAngle As Int16) As Single
		Dim lTX As Int32 = CInt((2.0F * lLocX + mlFullSizeWH) / moTerrain.CellSpacing)
		Dim lTZ As Int32 = CInt((2.0F * lLocZ + mlFullSizeWH) / moTerrain.CellSpacing)
        Dim lTA As Int32

        'Ensure angle is in valid range
        If iAngle < 0 Then
            iAngle += 3600S
        ElseIf iAngle > 3600 Then
            iAngle -= 3600S
        End If
		lTA = CInt(iAngle \ 450S)

        If (lTX < 0 OrElse lTX > mlDoubleWH - 1) OrElse (lTZ < 0 OrElse lTZ > mlDoubleWH - 1) Then
            'Range is outside bounds...
            Return 1.0F
        Else
            'Return (miSpeedMod(lTX, lTZ, lTA) * 0.01F) ' / 100.0F)
            Dim fTX As Single = lTX / 2.0F
            Dim fTZ As Single = lTZ / 2.0F
            Dim vTemp As Vector3 = moTerrain.GetTerrainNormalEx(fTX, fTZ)

            Select Case lTA
                Case 0
                    Return (1.0F + vTemp.X)
                Case 7
                    Return (1.0F + (vTemp.Z + vTemp.X))
                Case 6
                    Return (1.0F + vTemp.Z)
                Case 5
                    Return (1.0F + (vTemp.Z + -vTemp.X))
                Case 4
                    Return (1.0F + -vTemp.X)
                Case 3
                    Return (1.0F + (-vTemp.Z + -vTemp.X))
                Case 2
                    Return (1.0F + -vTemp.Z)
                Case 1
                    Return (1.0F + (-vTemp.Z + vTemp.X))
            End Select
        End If
    End Function

    Public ReadOnly Property WaterHeight() As Int32
        Get
            Return CInt(moTerrain.ml_Y_Mult * moTerrain.WaterHeight)
        End Get
    End Property

    Public Function TerrainGradeBuildable(ByVal lLocX As Int32, ByVal lLocZ As Int32) As Boolean
        'Need to convert locX to Vert
		Dim fX As Single = CSng(CSng(lLocX) / moTerrain.CellSpacing) + (TerrainClass.Width \ 2)
		Dim fZ As Single = CSng(CSng(lLocZ) / moTerrain.CellSpacing) + (TerrainClass.Height \ 2)
        Dim vecTemp As Vector3 = moTerrain.GetTerrainNormalEx(fX, fZ)
        Return Not (Math.Abs(vecTemp.X) > 0.4F OrElse Math.Abs(vecTemp.Z) > 0.4F)
    End Function

    Private Sub SetupStartLocGrid()

        ReDim myStartLocGrid(ml_START_LOC_GRID_WH * ml_START_LOC_GRID_WH - 1)

        Dim lVertsPerGrid As Int32 = TerrainClass.Width \ ml_START_LOC_GRID_WH
        Dim lMinVal As Int32 = moTerrain.CellSpacing \ 4
        Dim lHalfWidth As Int32 = TerrainClass.Width \ 2
        Dim lHalfHeight As Int32 = TerrainClass.Height \ 2

        Dim lWaterHeightValue As Int32 = CInt(moTerrain.WaterHeight * moTerrain.ml_Y_Mult)

        For lGridY As Int32 = 0 To ml_START_LOC_GRID_WH - 1
            For lGridX As Int32 = 0 To ml_START_LOC_GRID_WH - 1

                Dim lGoodPoints As Int32 = 0
                Dim lBadPoints As Int32 = 0

                'Ok, now go through the verts in this square starting at 1 and ending on lvertspergrid - 2
                For Y As Int32 = (lGridY * lVertsPerGrid) + 1 To ((lGridY + 1) * lVertsPerGrid) - 2
                    For X As Int32 = (lGridX * lVertsPerGrid) + 1 To ((lGridX + 1) * lVertsPerGrid) - 2
                        'Now, for this vert... get its location
                        Dim lVertLocX As Int32 = (X - lHalfWidth) * moTerrain.CellSpacing
                        Dim lVertLocY As Int32 = (Y - lHalfHeight) * moTerrain.CellSpacing

                        'Now, 4 tests... 
                        Dim lLocX As Int32 = lVertLocX - lMinVal
                        Dim lLocY As Int32 = lVertLocY - lMinVal

                        'up left
						If moTerrain.GetHeightAtLocation(lLocX, lLocY, False) > lWaterHeightValue Then
							If TerrainGradeBuildable(lLocX, lLocY) = True Then
								lGoodPoints += 1
							Else : lBadPoints += 1
							End If
						Else : lBadPoints += 1
						End If
                        'up right
                        lLocX = lVertLocX + lMinVal
                        lLocY = lVertLocY - lMinVal
						If moTerrain.GetHeightAtLocation(lLocX, lLocY, False) > lWaterHeightValue Then
							If TerrainGradeBuildable(lLocX, lLocY) = True Then
								lGoodPoints += 1
							Else : lBadPoints += 1
							End If
						Else : lBadPoints += 1
						End If
                        'down left
                        lLocX = lVertLocX - lMinVal
                        lLocY = lVertLocY + lMinVal
						If moTerrain.GetHeightAtLocation(lLocX, lLocY, False) > lWaterHeightValue Then
							If TerrainGradeBuildable(lLocX, lLocY) = True Then
								lGoodPoints += 1
							Else : lBadPoints += 1
							End If
						Else : lBadPoints += 1
						End If
                        'down right
                        lLocX = lVertLocX + lMinVal
                        lLocY = lVertLocY + lMinVal
						If moTerrain.GetHeightAtLocation(lLocX, lLocY, False) > lWaterHeightValue Then
							If TerrainGradeBuildable(lLocX, lLocY) = True Then
								lGoodPoints += 1
							Else : lBadPoints += 1
							End If
						Else : lBadPoints += 1
						End If

                    Next X
                Next Y

                'Now, determine if this is a good spot... this is based on whether the Bad Points exceed 1/3 of the total points
                If lGoodPoints >= 2 * lBadPoints Then
                    myStartLocGrid(lGridY * ml_START_LOC_GRID_WH + lGridX) = 255        'it is a good spot
                Else : myStartLocGrid(lGridY * ml_START_LOC_GRID_WH + lGridX) = 0       'it is a bad spot
                End If

            Next lGridX
        Next lGridY
    End Sub

    Public Function GetPlayerStartLocation() As Point
        'Ok, determine how many start locs we have available
        If myStartLocGrid Is Nothing Then Return Point.Empty

        Dim lIdx As Int32
        Dim lLocX As Int32
        Dim lLocY As Int32

        Dim bProceed As Boolean = False

        Randomize(Val(Now.ToString("MMddHHmm")))

        Dim lMarkSize As Int32 = 1
        If Me.PlanetSizeID = 0 Then lMarkSize = 2

        'Before continuing, check if there are any facilities within this square
        Dim oEnvir As Envir = Nothing
        For X As Int32 = 0 To glEnvirUB
            If goEnvirs(X).ObjectID = Me.ObjectID AndAlso goEnvirs(X).ObjTypeID = Me.ObjTypeID Then
                oEnvir = goEnvirs(X)
                Exit For
            End If
        Next X
        If oEnvir Is Nothing Then Return Point.Empty

        While bProceed = False
            bProceed = True

            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To myStartLocGrid.GetUpperBound(0)
                If myStartLocGrid(X) <> 0 Then lCnt += 1
            Next X

            If lCnt = 0 Then Return Point.Empty

            'Now, get our index
            lCnt = CInt(Math.Floor(Rnd() * lCnt))

            lIdx = -1
            For X As Int32 = 0 To myStartLocGrid.GetUpperBound(0)
                If myStartLocGrid(X) <> 0 Then
                    lCnt -= 1
                    If lCnt = 0 Then
                        lIdx = X
                        Exit For
                    End If
                End If
            Next X

            If lIdx = -1 Then Return Point.Empty

            lLocY = lIdx \ ml_START_LOC_GRID_WH
            lLocX = lIdx - (lLocY * ml_START_LOC_GRID_WH)


            For Y As Int32 = lLocY - lMarkSize To lLocY + lMarkSize
                For X As Int32 = lLocX - lMarkSize To lLocX + lMarkSize
                    'This handles Map Wrapping
                    Dim lTmpLocX As Int32 = X
                    If lTmpLocX < 0 Then
                        lTmpLocX = 0 ' ml_START_LOC_GRID_WH
                    ElseIf lTmpLocX > ml_START_LOC_GRID_WH - 1 Then
                        lTmpLocX = ml_START_LOC_GRID_WH - 1
                    End If
                    Dim lTmpLocY As Int32 = Y
                    If lTmpLocY < 0 Then
                        lTmpLocY = 0
                    ElseIf lTmpLocY > ml_START_LOC_GRID_WH - 1 Then
                        lTmpLocY = ml_START_LOC_GRID_WH - 1
                    End If

                    Dim lTmpIdx As Int32 = lTmpLocY * ml_START_LOC_GRID_WH + lTmpLocX
                    If lTmpIdx > -1 AndAlso lTmpIdx <= myStartLocGrid.GetUpperBound(0) Then
                        If myStartLocGrid(lTmpIdx) = 0 Then
                            bProceed = False
                            Continue While
                        End If
                    End If
                Next X
            Next Y

            If oEnvir Is Nothing = False Then
                Dim lMult As Int32 = oEnvir.lEnvirSize \ ml_START_LOC_GRID_WH
                Dim lTrueLocX As Int32 = (lLocX - (ml_START_LOC_GRID_WH \ 2)) * lMult
                Dim lTrueLocY As Int32 = (lLocY - (ml_START_LOC_GRID_WH \ 2)) * lMult

                Dim lGridX As Int32 = (lTrueLocX + oEnvir.lHalfEnvirSize) \ oEnvir.lGridSquareSize
                Dim lGridY As Int32 = (lTrueLocY + oEnvir.lHalfEnvirSize) \ oEnvir.lGridSquareSize
                Dim lGridIdx As Int32 = lGridY * oEnvir.lGridsPerRow + lGridX

                Dim lFacCnt As Int32 = 0

                If lGridIdx > -1 AndAlso lGridIdx <= oEnvir.lGridUB Then
                    Dim oTmpGrid As EnvirGrid = oEnvir.GetGridNoAdd(lGridIdx)
                    If oTmpGrid Is Nothing = False Then
                        For X As Int32 = 0 To oTmpGrid.lEntityUB
                            Dim lTemp As Int32 = oTmpGrid.lEntities(X)
                            If lTemp <> -1 AndAlso glEntityIdx(lTemp) > 0 Then
                                If goEntity(lTemp).ObjTypeID = ObjectType.eFacility Then
                                    'ok, facility is here...
                                    lFacCnt += 1
                                End If
                            End If
                        Next X
                    End If
                End If

                If lFacCnt > ml_MAX_START_LOC_FAC_COUNT AndAlso (Me.ParentSystem Is Nothing OrElse Me.ParentSystem.ObjectID <> 36) Then
                    'set this spot as marked
                    myStartLocGrid(lIdx) = 0
                    'cannot proceed, find a new spot
                    bProceed = False
                End If
            End If
        End While

        'Mark this square and the 8 squares around it as UNUSEABLE... (or if it is a tiny, the squares 2 squares away
        For Y As Int32 = lLocY - lMarkSize To lLocY + lMarkSize
            For X As Int32 = lLocX - lMarkSize To lLocX + lMarkSize
                'This handles Map Wrapping
                Dim lTmpLocX As Int32 = X
                If lTmpLocX < 0 Then
                    lTmpLocX = 0 ' ml_START_LOC_GRID_WH
                ElseIf lTmpLocX > ml_START_LOC_GRID_WH - 1 Then
                    lTmpLocX = ml_START_LOC_GRID_WH - 1
                End If
                Dim lTmpLocY As Int32 = Y
                If lTmpLocY < 0 Then
                    lTmpLocY = 0
                ElseIf lTmpLocY > ml_START_LOC_GRID_WH - 1 Then
                    lTmpLocY = ml_START_LOC_GRID_WH - 1
                End If

                Dim lTmpIdx As Int32 = lTmpLocY * ml_START_LOC_GRID_WH + lTmpLocX
                If lTmpIdx > -1 AndAlso lTmpIdx <= myStartLocGrid.GetUpperBound(0) Then myStartLocGrid(lTmpIdx) = 0
            Next X
        Next Y

        'Now, find a suitable place within this square... so adjust LocX and LocY as actual coordinates
        Dim lVertsPerGrid As Int32 = (TerrainClass.Width \ ml_START_LOC_GRID_WH)
        Dim lHalfVertsPerGrid As Int32 = lVertsPerGrid \ 2
        lLocX *= lVertsPerGrid
        lLocY *= lVertsPerGrid
        lLocX += lHalfVertsPerGrid
        lLocY += lHalfVertsPerGrid

        Dim bGood As Boolean = False

        'reusing marksize
        lMarkSize = 0
        While bGood = False AndAlso lMarkSize < lHalfVertsPerGrid
            For Y As Int32 = lLocY - lMarkSize To lLocY + lMarkSize
                For X As Int32 = lLocX - lMarkSize To lLocX + lMarkSize
                    If moTerrain.GetVertexHeight(X, Y) > moTerrain.WaterHeight Then
                        bGood = True
                        lLocX = X
                        lLocY = Y
                        Exit While
                    End If
                Next X
            Next Y
            lMarkSize += 1
        End While

        If bGood = False Then Return Point.Empty

        'Ok, now... lLocX and lLocY indicate our STARTING vertex... so transform that to world coordinates
        lLocX -= (TerrainClass.Width \ 2)
        lLocY -= (TerrainClass.Height \ 2)
        lLocX *= moTerrain.CellSpacing
        lLocY *= moTerrain.CellSpacing

        'And return that point
        Return New System.Drawing.Point(lLocX, lLocY)
    End Function

    Public Sub SetPlayerStartLocationMarked(ByVal lLocX As Int32, ByVal lLocZ As Int32)
        'ok, coordinates passed in our actual location... convert that to Vert
        Dim lVertX As Int32 = lLocX \ moTerrain.CellSpacing
        Dim lVertZ As Int32 = lLocZ \ moTerrain.CellSpacing
        lVertX += (TerrainClass.Width \ 2)
        lVertZ += (TerrainClass.Height \ 2)

        'Now, get the grid loc
        Dim lVertsPerGrid As Int32 = (TerrainClass.Width \ ml_START_LOC_GRID_WH)
        Dim lGridX As Int32 = lVertX \ lVertsPerGrid
        Dim lGridZ As Int32 = lVertZ \ lVertsPerGrid

        Dim lMarkSize As Int32 = 1
        If Me.PlanetSizeID = 0 Then lMarkSize = 2
        For Y As Int32 = lGridZ - lMarkSize To lGridZ + lMarkSize
            For X As Int32 = lGridX - lMarkSize To lGridX + lMarkSize
                'This handles Map Wrapping
                Dim lTmpLocX As Int32 = X
                If lTmpLocX < 0 Then
                    lTmpLocX += ml_START_LOC_GRID_WH
                ElseIf lTmpLocX > ml_START_LOC_GRID_WH - 1 Then
                    lTmpLocX -= ml_START_LOC_GRID_WH
                End If

                Dim lTmpIdx As Int32 = Y * ml_START_LOC_GRID_WH + lTmpLocX
                If lTmpIdx > -1 AndAlso lTmpIdx <= myStartLocGrid.GetUpperBound(0) Then myStartLocGrid(lTmpIdx) = 0
            Next X
        Next Y
    End Sub

    Public Sub SetPirateStartLocationMarked(ByVal lLocX As Int32, ByVal lLocZ As Int32)
        If lLocX = Int32.MinValue OrElse lLocZ = Int32.MinValue Then Return

        'ok, coordinates passed in our actual location... convert that to Vert
        Dim lVertX As Int32 = lLocX \ moTerrain.CellSpacing
        Dim lVertZ As Int32 = lLocZ \ moTerrain.CellSpacing
        lVertX += (TerrainClass.Width \ 2)
        lVertZ += (TerrainClass.Height \ 2)

        'Now, get the grid loc
        Dim lVertsPerGrid As Int32 = (TerrainClass.Width \ ml_START_LOC_GRID_WH)
        Dim lGridX As Int32 = lVertX \ lVertsPerGrid
        Dim lGridZ As Int32 = lVertZ \ lVertsPerGrid

        Dim lTmpIdx As Int32 = lGridZ * ml_START_LOC_GRID_WH + lGridX
        If lTmpIdx > -1 AndAlso lTmpIdx <= myStartLocGrid.GetUpperBound(0) Then myStartLocGrid(lTmpIdx) = 0
    End Sub

    Public Function GetPirateStartLocation(ByVal lPlayerID As Int32, ByVal lPlayerStartLocX As Int32, ByVal lPlayerStartLocZ As Int32) As Point
        Dim lVertX As Int32 = lPlayerStartLocX \ moTerrain.CellSpacing
        Dim lVertZ As Int32 = lPlayerStartLocZ \ moTerrain.CellSpacing
        lVertX += (TerrainClass.Width \ 2)
        lVertZ += (TerrainClass.Height \ 2)
        Dim lVertsPerGrid As Int32 = (TerrainClass.Width \ ml_START_LOC_GRID_WH)
        Dim lGridX As Int32 = lVertX \ lVertsPerGrid
        Dim lGridZ As Int32 = lVertZ \ lVertsPerGrid

        'Now, find the player's command center
        Dim bUsePlayerStartLoc As Boolean = False
        'If the player's CC is not in the same start point square as the player's start point square or any of the adjacent squares, 
		'then the point is the player's start point 
		Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
            If glEntityIdx(X) > 0 AndAlso goEntity(X).yProductionType = ProductionType.eCommandCenterSpecial Then
                If goEntity(X).Owner.ObjectID = lPlayerID Then
                    Dim lFacVertX As Int32 = goEntity(X).LocX \ moTerrain.CellSpacing
                    Dim lFacVertZ As Int32 = goEntity(X).LocZ \ moTerrain.CellSpacing
                    lFacVertX += (TerrainClass.Width \ 2)
                    lFacVertZ += (TerrainClass.Height \ 2)
                    Dim lFacGridX As Int32 = lFacVertX \ lVertsPerGrid
                    Dim lFacGridZ As Int32 = lFacVertZ \ lVertsPerGrid

                    bUsePlayerStartLoc = Math.Abs(lFacGridX - lGridX) > 1 OrElse Math.Abs(lFacGridZ - lGridZ) > 1
                    Exit For
                End If
            End If
		Next X

        If bUsePlayerStartLoc = True Then Return New Point(lPlayerStartLocX, lPlayerStartLocZ)

        'Else: Any square adjacent to an Adjacent square of the player's start point. 
        'The square cannot be already marked as used.
        Dim lIndexList() As Int32 = Nothing
        Dim lIndexListUB As Int32 = -1

        Dim lTmpGZ As Int32 = 0
        If lGridZ > 1 Then
            lTmpGZ = lGridZ - 2
            For X As Int32 = -2 To 2
                Dim lTmpGX As Int32 = lGridX + X
                'This handles Map Wrapping
                If lTmpGX < 0 Then
                    lTmpGX += ml_START_LOC_GRID_WH
                ElseIf lTmpGX > ml_START_LOC_GRID_WH - 1 Then
                    lTmpGX -= ml_START_LOC_GRID_WH
                End If

                Dim lTmpIdx As Int32 = lTmpGZ * ml_START_LOC_GRID_WH + lTmpGX
                If lTmpIdx > -1 AndAlso lTmpIdx <= myStartLocGrid.GetUpperBound(0) Then
                    If myStartLocGrid(lTmpIdx) <> 0 Then
                        lIndexListUB += 1
                        ReDim Preserve lIndexList(lIndexListUB)
                        lIndexList(lIndexListUB) = lTmpIdx
                    End If
                End If
            Next X
        End If
        For lTmpGZ = lGridZ - 1 To lGridZ + 1
            If lTmpGZ > -1 AndAlso lTmpGZ < ml_START_LOC_GRID_WH Then
                For X As Int32 = -2 To 2 Step 4
                    Dim lTmpGX As Int32 = lGridX + X
                    'This handles Map Wrapping
                    If lTmpGX < 0 Then
                        lTmpGX += ml_START_LOC_GRID_WH
                    ElseIf lTmpGX > ml_START_LOC_GRID_WH - 1 Then
                        lTmpGX -= ml_START_LOC_GRID_WH
                    End If

                    Dim lTmpIdx As Int32 = lTmpGZ * ml_START_LOC_GRID_WH + lTmpGX
                    If lTmpIdx > -1 AndAlso lTmpIdx <= myStartLocGrid.GetUpperBound(0) Then
                        If myStartLocGrid(lTmpIdx) <> 0 Then
                            lIndexListUB += 1
                            ReDim Preserve lIndexList(lIndexListUB)
                            lIndexList(lIndexListUB) = lTmpIdx
                        End If
                    End If
                Next X
            End If
        Next lTmpGZ

        lTmpGZ = lGridZ + 2
        If lTmpGZ > -1 AndAlso lTmpGZ < ml_START_LOC_GRID_WH Then
            For X As Int32 = -2 To 2
                Dim lTmpGX As Int32 = lGridX + X
                'This handles Map Wrapping
                If lTmpGX < 0 Then
                    lTmpGX += ml_START_LOC_GRID_WH
                ElseIf lTmpGX > ml_START_LOC_GRID_WH - 1 Then
                    lTmpGX -= ml_START_LOC_GRID_WH
                End If

                Dim lTmpIdx As Int32 = lTmpGZ * ml_START_LOC_GRID_WH + lTmpGX
                If lTmpIdx > -1 AndAlso lTmpIdx <= myStartLocGrid.GetUpperBound(0) Then
                    If myStartLocGrid(lTmpIdx) <> 0 Then
                        lIndexListUB += 1
                        ReDim Preserve lIndexList(lIndexListUB)
                        lIndexList(lIndexListUB) = lTmpIdx
                    End If
                End If
            Next X
        End If

        Dim lPirateStartIndex As Int32 = -1

        If lIndexListUB = -1 Then
            'Ok, we'll just choose a spot at random... this is copied right out of player start loc
            Dim lIdx As Int32
            Dim lLocX As Int32
            Dim lLocY As Int32

            Dim bProceed As Boolean = False

            While bProceed = False
                bProceed = True

                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To myStartLocGrid.GetUpperBound(0)
                    If myStartLocGrid(X) <> 0 Then lCnt += 1
                Next X

                If lCnt = 0 Then Return Point.Empty

                'Now, get our index
                lCnt = CInt(Math.Floor(Rnd() * lCnt))

                lIdx = -1
                For X As Int32 = 0 To myStartLocGrid.GetUpperBound(0)
                    If myStartLocGrid(X) <> 0 Then
                        lCnt -= 1
                        If lCnt = 0 Then
                            lIdx = X
                            Exit For
                        End If
                    End If
                Next X

                If lIdx = -1 Then Return Point.Empty

                lLocY = lIdx \ ml_START_LOC_GRID_WH
                lLocX = lIdx - (lLocY * ml_START_LOC_GRID_WH)

                'Before continuing, check if there are any facilities within this square
                Dim oEnvir As Envir = Nothing
                For X As Int32 = 0 To glEnvirUB
                    If goEnvirs(X).ObjectID = Me.ObjectID AndAlso goEnvirs(X).ObjTypeID = Me.ObjTypeID Then
                        oEnvir = goEnvirs(X)
                        Exit For
                    End If
                Next X
                If oEnvir Is Nothing = False Then
                    Dim lMult As Int32 = oEnvir.lEnvirSize \ ml_START_LOC_GRID_WH
                    Dim lTrueLocX As Int32 = (lLocX - (ml_START_LOC_GRID_WH \ 2)) * lMult
                    Dim lTrueLocY As Int32 = (lLocY - (ml_START_LOC_GRID_WH \ 2)) * lMult

                    Dim lTmpGridX As Int32 = (lTrueLocX + oEnvir.lHalfEnvirSize) \ oEnvir.lGridSquareSize
                    Dim lTmpGridY As Int32 = (lTrueLocY + oEnvir.lHalfEnvirSize) \ oEnvir.lGridSquareSize
                    Dim lGridIdx As Int32 = lTmpGridY * oEnvir.lGridsPerRow + lTmpGridX

                    Dim lFacCnt As Int32 = 0

                    If lGridIdx > -1 AndAlso lGridIdx <= oEnvir.lGridUB Then
                        Dim oTmpGrid As EnvirGrid = oEnvir.GetGridNoAdd(lGridIdx)
                        If oTmpGrid Is Nothing = False Then
                            For X As Int32 = 0 To oTmpGrid.lEntityUB
                                Dim lTemp As Int32 = oTmpGrid.lEntities(X)
                                If lTemp <> -1 AndAlso glEntityIdx(lTemp) > 0 Then
                                    If goEntity(lTemp).ObjTypeID = ObjectType.eFacility Then
                                        'ok, facility is here...
                                        lFacCnt += 1
                                    End If
                                End If
                            Next X
                        End If
                    End If

                    If lFacCnt > ml_MAX_START_LOC_FAC_COUNT Then
                        'set this spot as marked
                        myStartLocGrid(lIdx) = 0
                        'cannot proceed, find a new spot
                        bProceed = False
                    End If
                End If
            End While

            lPirateStartIndex = lIdx
        Else
            'Ok, randomize in the list
            Dim lRndVal As Int32 = CInt(Rnd() * lIndexListUB)
            lPirateStartIndex = lIndexList(lRndVal)
        End If

        If lPirateStartIndex = -1 Then Return Point.Empty

        'Once the pirate start loc is chosen, it is marked as used (only that square)
        myStartLocGrid(lPirateStartIndex) = 0

        'Now, return just the closest thing
        Dim ptResult As Point
        ptResult.Y = lPirateStartIndex \ ml_START_LOC_GRID_WH
        ptResult.X = lPirateStartIndex - (ptResult.Y * ml_START_LOC_GRID_WH)
        ptResult.X *= lVertsPerGrid
        ptResult.Y *= lVertsPerGrid

        ptResult.X -= (TerrainClass.Width \ 2)
        ptResult.Y -= (TerrainClass.Height \ 2)

        ptResult.X *= moTerrain.CellSpacing
        ptResult.Y *= moTerrain.CellSpacing

        Return ptResult

    End Function

    Public Function PlacePirateFacilities(ByRef yData() As Byte) As Byte()
        '- MsgCode, EnvirGUID(6), StartLoc (8), PlayerID (4), ItmCnt (4), Items...
        Dim lPos As Int32 = 8
        Dim lSX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lSZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += 4
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim yResp(23 + (lCnt * 20)) As Byte
        Dim lRespPos As Int32 = 0

        Array.Copy(yData, 0, yResp, 0, 24)
        lRespPos += 24

        'Now, get the grid loc
        Dim lVertsPerGrid As Int32 = (TerrainClass.Width \ ml_START_LOC_GRID_WH)
        Dim lGridX As Int32 = (lSX + ((TerrainClass.Width \ 2) * moTerrain.CellSpacing)) \ (lVertsPerGrid * moTerrain.CellSpacing)
        Dim lGridZ As Int32 = (lSZ + ((TerrainClass.Width \ 2) * moTerrain.CellSpacing)) \ (lVertsPerGrid * moTerrain.CellSpacing)

        'Now, determine our range... do this by multiplying verts per grid...
        lGridX *= lVertsPerGrid
        lGridZ *= lVertsPerGrid
        'GridX/Z is in Verts now
        lGridX -= (TerrainClass.Width \ 2)
        lGridZ -= (TerrainClass.Height \ 2)
        'now, multiply by cellspacing
        lGridX *= moTerrain.CellSpacing
        lGridZ *= moTerrain.CellSpacing

        'Now, lGridX/Z is our upper left corner loc
        Dim rcFullExtent As Rectangle
        rcFullExtent.X = lGridX
        rcFullExtent.Y = lGridZ
        rcFullExtent.Width = (lVertsPerGrid * moTerrain.CellSpacing)
        rcFullExtent.Height = rcFullExtent.Width
        Dim rcHalfExtent As Rectangle
        rcHalfExtent.X = lGridX + (rcFullExtent.Width \ 4)
        rcHalfExtent.Y = lGridZ + (rcFullExtent.Height \ 4)
        rcHalfExtent.Width = rcFullExtent.Width \ 2
        rcHalfExtent.Height = rcFullExtent.Height \ 2

        Dim lRealWaterHeight As Int32 = Me.WaterHeight

        'Now, each item is:
        'ObjectGUID (6)
        'DefGUID (6)
        Dim rcUsed(lCnt - 1) As Rectangle
        For X As Int32 = 0 To lCnt - 1
            Array.Copy(yData, lPos, yResp, lRespPos, 6) : lPos += 6 : lRespPos += 6

            Dim lDefID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iDefTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim lLocX As Int32 = Int32.MinValue
            Dim lLocZ As Int32 = Int32.MinValue

            If iDefTypeID = ObjectType.eUnitDef Then
                lLocX = lSX
                lLocZ = lSZ
                rcUsed(X) = Rectangle.Empty
            Else
                Dim oDef As Epica_Entity_Def = Nothing
                For Y As Int32 = 0 To glEntityDefUB
                    If glEntityDefIdx(Y) = lDefID AndAlso goEntityDefs(Y).ObjTypeID = iDefTypeID Then
                        oDef = goEntityDefs(Y)
                        Exit For
                    End If
                Next Y

                If oDef Is Nothing = False Then
                    'Now, figure out where this item will go...
                    Dim lHalfRectSize As Int32 = ((oDef.lModelRangeOffset * gl_FINAL_GRID_SQUARE_SIZE) + 1) \ 2
                    If oDef.ProductionTypeID <> ProductionType.eNoProduction Then
                        'use half extent
                        Dim bGood As Boolean = False
                        Dim lAttempts As Int32 = 0

                        While bGood = False
                            lAttempts += 1
                            If lAttempts > 10 Then Exit While
                            lLocX = CInt(Rnd() * rcHalfExtent.Width) + rcHalfExtent.X
							lLocZ = CInt(Rnd() * rcHalfExtent.Height) + rcHalfExtent.Y

							'Dim lHalfVal As Int32 = oDef.lModelSizeXZ \ 2
							'Dim lMinHt As Int32 = Int32.MaxValue
							'Dim lMaxHt As Int32 = Int32.MinValue
							'With moTerrain
							'	Dim LocX As Int32 = CInt(lLocX)
							'	Dim LocZ As Int32 = CInt(lLocZ)
							'	Dim lTemp As Int32 = CInt(.GetHeightAtLocation(LocX - lHalfVal, LocZ - lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX + lHalfVal, LocZ - lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX - lHalfVal, LocZ + lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX + lHalfVal, LocZ + lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX, LocZ))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'End With
							'If lMaxHt - lMinHt > 400 Then		'a little more leeway (200 is what it normally)
							'	bGood = False
							'	Continue While
							'End If
							'If lMinHt < lRealWaterHeight Then
							'	bGood = False
							'	Continue While
							'End If

							If moTerrain.GetHeightAtLocation(lLocX, lLocZ, False) > lRealWaterHeight AndAlso TerrainGradeBuildable(lLocX, lLocZ) = True Then
								Dim rcTemp As Rectangle = New Rectangle(lLocX - lHalfRectSize, lLocZ - lHalfRectSize, lHalfRectSize * 2, lHalfRectSize * 2)
								bGood = True
								For Y As Int32 = 0 To X - 1
									If rcUsed(Y).IsEmpty = False AndAlso rcUsed(Y).IntersectsWith(rcTemp) = True Then
										bGood = False
										Exit For
									End If
								Next Y
								If bGood = True Then rcUsed(X) = rcTemp
							End If
							
                        End While

                        While bGood = False
                            lAttempts += 1
                            If lAttempts > 20 Then Exit While
                            lLocX = CInt(Rnd() * rcFullExtent.Width) + rcFullExtent.X
                            lLocZ = CInt(Rnd() * rcFullExtent.Height) + rcFullExtent.Y

							If moTerrain.GetHeightAtLocation(lLocX, lLocZ, False) > lRealWaterHeight AndAlso TerrainGradeBuildable(lLocX, lLocZ) = True Then
								Dim rcTemp As Rectangle = New Rectangle(lLocX - lHalfRectSize, lLocZ - lHalfRectSize, lHalfRectSize * 2, lHalfRectSize * 2)
								bGood = True
								For Y As Int32 = 0 To X - 1
									If rcUsed(Y).IsEmpty = False AndAlso rcTemp.IntersectsWith(rcTemp) = True Then
										bGood = False
										Exit For
									End If
								Next Y
								If bGood = True Then rcUsed(X) = rcTemp
							End If
                        End While
                    Else
                        'use full extent and prefer locations outside of half extent
                        Dim bGood As Boolean = False
                        Dim lAttempts As Int32 = 0
                        Dim lQuarter As Int32 = rcHalfExtent.Width \ 2
                        While bGood = False
                            lAttempts += 1
                            If lAttempts > 10 Then Exit While
                            lLocX = CInt(Rnd() * rcHalfExtent.Width) '+ rcHalfExtent.X
                            lLocZ = CInt(Rnd() * rcHalfExtent.Height) '+ rcHalfExtent.Y
                            If lLocX < lQuarter Then lLocX = rcFullExtent.X + lLocX Else lLocX = rcFullExtent.X + rcFullExtent.Width - lLocX
							If lLocZ < lQuarter Then lLocZ = rcFullExtent.Y + lLocZ Else lLocZ = rcFullExtent.Y + rcFullExtent.Width - lLocZ

							'Dim lHalfVal As Int32 = oDef.lModelSizeXZ \ 2
							'Dim lMinHt As Int32 = Int32.MaxValue
							'Dim lMaxHt As Int32 = Int32.MinValue
							'With moTerrain
							'	Dim LocX As Int32 = CInt(lLocX)
							'	Dim LocZ As Int32 = CInt(lLocZ)
							'	Dim lTemp As Int32 = CInt(.GetHeightAtLocation(LocX - lHalfVal, LocZ - lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX + lHalfVal, LocZ - lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX - lHalfVal, LocZ + lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX + lHalfVal, LocZ + lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX, LocZ))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'End With
							'If lMaxHt - lMinHt > 400 Then		'a little more leeway (200 is what it normally)
							'	bGood = False
							'	Continue While
							'End If
							'If lMinHt < lRealWaterHeight Then
							'	bGood = False
							'	Continue While
							'End If

							If moTerrain.GetHeightAtLocation(lLocX, lLocZ, False) > lRealWaterHeight AndAlso TerrainGradeBuildable(lLocX, lLocZ) = True Then
								Dim rcTemp As Rectangle = New Rectangle(lLocX - lHalfRectSize, lLocZ - lHalfRectSize, lHalfRectSize * 2, lHalfRectSize * 2)
								bGood = True
								For Y As Int32 = 0 To X - 1
									If rcUsed(Y).IsEmpty = False AndAlso rcTemp.IntersectsWith(rcTemp) = True Then
										bGood = False
										Exit For
									End If
								Next Y
								If bGood = True Then rcUsed(X) = rcTemp
							End If
                        End While

                        While bGood = False
                            lAttempts += 1
                            If lAttempts > 20 Then Exit While
                            lLocX = CInt(Rnd() * rcFullExtent.Width) + rcFullExtent.X
							lLocZ = CInt(Rnd() * rcFullExtent.Height) + rcFullExtent.Y

							'Dim lHalfVal As Int32 = oDef.lModelSizeXZ \ 2
							'Dim lMinHt As Int32 = Int32.MaxValue
							'Dim lMaxHt As Int32 = Int32.MinValue
							'With moTerrain
							'	Dim LocX As Int32 = CInt(lLocX)
							'	Dim LocZ As Int32 = CInt(lLocZ)
							'	Dim lTemp As Int32 = CInt(.GetHeightAtLocation(LocX - lHalfVal, LocZ - lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX + lHalfVal, LocZ - lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX - lHalfVal, LocZ + lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX + lHalfVal, LocZ + lHalfVal))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'	lTemp = CInt(.GetHeightAtLocation(LocX, LocZ))
							'	If lTemp > lMaxHt Then lMaxHt = lTemp
							'	If lTemp < lMinHt Then lMinHt = lTemp
							'End With
							'If lMaxHt - lMinHt > 400 Then		'a little more leeway (200 is what it normally)
							'	bGood = False
							'	Continue While
							'End If
							'If lMinHt < lRealWaterHeight Then
							'	bGood = False
							'	Continue While
							'End If

							If moTerrain.GetHeightAtLocation(lLocX, lLocZ, False) > lRealWaterHeight AndAlso TerrainGradeBuildable(lLocX, lLocZ) = True Then
								Dim rcTemp As Rectangle = New Rectangle(lLocX - lHalfRectSize, lLocZ - lHalfRectSize, lHalfRectSize * 2, lHalfRectSize * 2)
								bGood = True
								For Y As Int32 = 0 To X - 1
									If rcUsed(Y).IsEmpty = False AndAlso rcTemp.IntersectsWith(rcTemp) = True Then
										bGood = False
										Exit For
									End If
								Next Y
								If bGood = True Then rcUsed(X) = rcTemp
							End If
                        End While
                    End If
                End If
			End If

			If moTerrain.GetHeightAtLocation(lLocX, lLocZ, False) < lRealWaterHeight Then
				lLocX = Int32.MinValue
				lLocZ = Int32.MinValue
            End If

            System.BitConverter.GetBytes(lLocX).CopyTo(yResp, lRespPos) : lRespPos += 4
            System.BitConverter.GetBytes(lLocZ).CopyTo(yResp, lRespPos) : lRespPos += 4
        Next X

        Return yResp
    End Function
End Class

Public Class TerrainClass

#Region "Constant Expressions"
    Public Const Width As Int32 = 240 '256
    Public Const Height As Int32 = 240 '256
    Private Const mlPASSES As Int32 = 5
    Private Const mlQuads As Int32 = 24 '8      '8x8 quad = 64 quads total, this is to segment the total vertex buffer further
    Private Const mlVertsPerQuad As Int32 = CInt(Width / mlQuads)
    Private Const mlHalfHeight As Int32 = CInt(Height / 2)
    Private Const mlHalfWidth As Int32 = CInt(Width / 2)
    Private Const VertsTotal As Int32 = Width * Height
    Private Const QuadsX As Int32 = Width - 1
    Private Const QuadsZ As Int32 = Height - 1
    Private Const TrisX As Int32 = CInt(QuadsX / 2)
#End Region

    Public CellSpacing As Int32 = 200
    Public ml_Y_Mult As Single = 15.0       'TODO: does this need to be a single???
    Public MapType As Int32
    Public WaterHeight As Byte      'height of the water level

#Region "Private Variables"
    'Our heightmap array, we need to keep this for... getting heights at locations
    Private HeightMap() As Byte
    'Indicates if the Heightmap has been generated yet
    Private mbHMReady As Boolean = False
    'the Parent Planet ID... set in New()
    Private mlSeed As Int32

    Private HMNormals() As Vector3
#End Region

    Public Sub New(ByVal lSeed As Int32)
        mlSeed = lSeed
    End Sub

    Public Function GetVertexHeight(ByVal lVertX As Int32, ByVal lVertY As Int32) As Int32
        Dim lIdx As Int32 = lVertY * Width + lVertX
        If lIdx > -1 AndAlso lIdx < HeightMap.Length Then Return CInt(HeightMap(lIdx)) Else Return 0
    End Function

    'NOTE: I changed X and Z to Int32 for performance... also, this returned a Single before, now it is an Int32
	Public Function GetHeightAtLocation(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal bWaterHeightMin As Boolean) As Int32
		Dim fTX As Single	'translated X 
		Dim fTZ As Single	'translated Z 
		Dim lCol As Int32
		Dim lRow As Int32

		Dim yA As Int32
		Dim yB As Int32
		Dim yC As Int32
		Dim yD As Int32

		Dim lIdx As Int32

		fTX = mlHalfWidth + (CSng(lLocX) / CellSpacing)
		fTZ = mlHalfHeight + (CSng(lLocZ) / CellSpacing)

		lCol = CInt(Math.Floor(fTX))
		lRow = CInt(Math.Floor(fTZ))
		fTX -= lCol
		fTZ -= lRow

		lIdx = (lRow * Width) + lCol
		If lIdx < 0 OrElse lIdx > HeightMap.Length - 1 Then yA = 0 Else yA = HeightMap(lIdx)
		lIdx = (lRow * Width) + lCol + 1
		If lIdx < 0 OrElse lIdx > HeightMap.Length - 1 Then yB = 0 Else yB = HeightMap(lIdx)
		lIdx = ((lRow + 1) * Width) + lCol
		If lIdx < 0 OrElse lIdx > HeightMap.Length - 1 Then yC = 0 Else yC = HeightMap(lIdx)
		lIdx = ((lRow + 1) * Width) + lCol + 1
		If lIdx < 0 OrElse lIdx > HeightMap.Length - 1 Then yD = 0 Else yD = HeightMap(lIdx)

		Dim fV1 As Single = yA * (1 - fTX) + yB * fTX
		Dim fV2 As Single = yC * (1 - fTX) + yD * fTX
		If bWaterHeightMin = True Then
			Return CInt(Math.Max((fV1 * (1 - fTZ) + fV2 * fTZ) * ml_Y_Mult, WaterHeight * ml_Y_Mult))
		Else : Return CInt((fV1 * (1 - fTZ) + fV2 * fTZ) * ml_Y_Mult)
		End If
	End Function

    'NOTE: Changed Singles to Int32
    Public Function HasLineOfSight(ByVal lX1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lZ2 As Int32, ByVal lAttackerHeight As Int32) As Boolean
        Dim lY1 As Int32
        Dim lY2 As Int32

		lY1 = GetHeightAtLocation(lX1, lZ1, True)
		lY2 = GetHeightAtLocation(lX2, lZ2, True)

        Return HasLineOfSight(lX1, lY1, lZ1, lX2, lY2, lZ2, lAttackerHeight)
    End Function

    'NOTE: Changed Singles to Int32
    Public Function HasLineOfSight(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32, ByVal lZ2 As Int32, ByVal lAttackerHt As Int32) As Boolean
        Dim lPt1Col As Int32
        Dim lPt1Row As Int32

        Dim lPt2Col As Int32
        Dim lPt2Row As Int32

        Dim fTmp As Single

        Dim lStep As Int32
        Dim lStepCnt As Int32

        Dim fXMod As Single
        Dim fZMod As Single

        Dim fX As Single
        Dim fZ As Single

        Dim fMaxHt As Single

        Dim lHtDiff As Int32

        Dim lAttackerHeight As Int32

        Dim lIdx As Int32

		fTmp = mlHalfWidth + (CSng(lX1) / CellSpacing)
        lPt1Col = CInt(Math.Floor(fTmp))
		fTmp = mlHalfHeight + (CSng(lZ1) / CellSpacing)
        lPt1Row = CInt(Math.Floor(fTmp))

		fTmp = mlHalfWidth + (CSng(lX2) / CellSpacing)
        lPt2Col = CInt(Math.Floor(fTmp))
		fTmp = mlHalfHeight + (CSng(lZ2) / CellSpacing)
        lPt2Row = CInt(Math.Floor(fTmp))

        lAttackerHeight = lY1 + lAttackerHt

        If lY1 < lY2 Then
            'in this case, we need to do an angle algorithm, but to reduce computations... we do it bro style
            fXMod = lPt2Col - lPt1Col
            fZMod = lPt2Row - lPt1Row
            lStepCnt = CInt(Math.Max(Math.Abs(fXMod), Math.Abs(fZMod)))
            fXMod /= lStepCnt
            fZMod /= lStepCnt

            fX = lPt1Col
            fZ = lPt1Row

            lHtDiff = lY2 - lY1
            For lStep = 0 To lStepCnt - 1
				fMaxHt = ((CSng(lStep) / lStepCnt) * lHtDiff) + lAttackerHeight
                lIdx = CInt((fZ * Width) + fX)
                If lIdx < 0 OrElse lIdx > HeightMap.GetUpperBound(0) OrElse HeightMap(lIdx) * ml_Y_Mult > fMaxHt Then
                    Return False
                End If
                fX += fXMod
                fZ += fZMod
            Next lStep
        Else
            'in this case, we need to do the reverse
            fXMod = lPt1Col - lPt2Col
            fZMod = lPt1Row - lPt2Row
            lStepCnt = CInt(Math.Max(Math.Abs(fXMod), Math.Abs(fZMod)))
            fXMod /= lStepCnt
            fZMod /= lStepCnt

            fX = lPt2Col
            fZ = lPt2Row

            lHtDiff = lY1 - lY2
            For lStep = 0 To lStepCnt - 1
				fMaxHt = lAttackerHeight - ((CSng(lStep) / lStepCnt) * lHtDiff)
                lIdx = CInt((fZ * Width) + fX)
                If HeightMap(lIdx) * ml_Y_Mult > fMaxHt Then
                    Return False
                End If
                fX += fXMod
                fZ += fZMod
            Next lStep
        End If
        Return True
    End Function

    Public Sub CleanResources()
        'A call to this sub means that the program wants to release all but the bare essential data...
        mbHMReady = False
        Erase HeightMap
    End Sub

    Private Shared mlPercLookup() As Int32 = Nothing
    Private Shared Sub SetupPercLookup()
        If mlPercLookup Is Nothing = False Then Return

        ReDim mlPercLookup(100)
        mlPercLookup(0) = 235
        mlPercLookup(1) = 576
        mlPercLookup(2) = 1152
        mlPercLookup(3) = 1728
        mlPercLookup(4) = 2304
        mlPercLookup(5) = 2880
        mlPercLookup(6) = 3456
        mlPercLookup(7) = 4032
        mlPercLookup(8) = 4608
        mlPercLookup(9) = 5184
        mlPercLookup(10) = 5760
        mlPercLookup(11) = 6336
        mlPercLookup(12) = 6912
        mlPercLookup(13) = 7488
        mlPercLookup(14) = 8064
        mlPercLookup(15) = 8640
        mlPercLookup(16) = 9216
        mlPercLookup(17) = 9792
        mlPercLookup(18) = 10368
        mlPercLookup(19) = 10944
        mlPercLookup(20) = 11520
        mlPercLookup(21) = 12096
        mlPercLookup(22) = 12672
        mlPercLookup(23) = 13248
        mlPercLookup(24) = 13824
        mlPercLookup(25) = 14400
        mlPercLookup(26) = 14976
        mlPercLookup(27) = 15552
        mlPercLookup(28) = 16128
        mlPercLookup(29) = 16704
        mlPercLookup(30) = 17280
        mlPercLookup(31) = 17856
        mlPercLookup(32) = 18432
        mlPercLookup(33) = 19008
        mlPercLookup(34) = 19584
        mlPercLookup(35) = 20160
        mlPercLookup(36) = 20736
        mlPercLookup(37) = 21312
        mlPercLookup(38) = 21888
        mlPercLookup(39) = 22464
        mlPercLookup(40) = 23040
        mlPercLookup(41) = 23616
        mlPercLookup(42) = 24192
        mlPercLookup(43) = 24768
        mlPercLookup(44) = 25344
        mlPercLookup(45) = 25920
        mlPercLookup(46) = 26496
        mlPercLookup(47) = 27072
        mlPercLookup(48) = 27648
        mlPercLookup(49) = 28224
        mlPercLookup(50) = 28800
        mlPercLookup(51) = 29376
        mlPercLookup(52) = 29952
        mlPercLookup(53) = 30528
        mlPercLookup(54) = 31104
        mlPercLookup(55) = 31680
        mlPercLookup(56) = 32256
        mlPercLookup(57) = 32832
        mlPercLookup(58) = 33408
        mlPercLookup(59) = 33984
        mlPercLookup(60) = 34560
        mlPercLookup(61) = 35136
        mlPercLookup(62) = 35712
        mlPercLookup(63) = 36288
        mlPercLookup(64) = 36864
        mlPercLookup(65) = 37440
        mlPercLookup(66) = 38016
        mlPercLookup(67) = 38592
        mlPercLookup(68) = 39168
        mlPercLookup(69) = 39744
        mlPercLookup(70) = 40320
        mlPercLookup(71) = 40896
        mlPercLookup(72) = 41472
        mlPercLookup(73) = 42048
        mlPercLookup(74) = 42624
        mlPercLookup(75) = 43200
        mlPercLookup(76) = 43776
        mlPercLookup(77) = 44352
        mlPercLookup(78) = 44928
        mlPercLookup(79) = 45504
        mlPercLookup(80) = 46080
        mlPercLookup(81) = 46656
        mlPercLookup(82) = 47232
        mlPercLookup(83) = 47808
        mlPercLookup(84) = 48384
        mlPercLookup(85) = 48960
        mlPercLookup(86) = 49536
        mlPercLookup(87) = 50112
        mlPercLookup(88) = 50688
        mlPercLookup(89) = 51264
        mlPercLookup(90) = 51840
        mlPercLookup(91) = 52416
        mlPercLookup(92) = 52992
        mlPercLookup(93) = 53568
        mlPercLookup(94) = 54144
        mlPercLookup(95) = 54720
        mlPercLookup(96) = 55296
        mlPercLookup(97) = 55872
        mlPercLookup(98) = 56448
        mlPercLookup(99) = 57024
        mlPercLookup(100) = 57600
    End Sub

    Private Sub GenerateTerrain(ByVal lSeed As Int32)
        SetupPercLookup()

        Dim lWidthSpans() As Int32
        Dim lHeightSpans() As Int32
        Dim X As Int32
        Dim lVal As Int32
        Dim lWaterPerc As Int32
        Dim bDone As Boolean
        Dim bLandmassLoop As Boolean
        Dim lStartX As Int32
        Dim lStartY As Int32
        Dim Y As Int32
        Dim lIdx As Int32
        Dim lPass As Int32

        Dim lSubXMax As Int32
        Dim lSubYMax As Int32
        Dim lSubX As Int32
        Dim lSubY As Int32

        Dim lTemp As Int32

        Dim lTotalLand As Int32

        Dim HM_Type() As Byte

        Call Rnd(-1)
        Randomize(lSeed)

        'fill our spans
        ReDim lWidthSpans(mlPASSES - 1)
        For X = 0 To mlPASSES - 1
            lVal = CInt(2 ^ X)
            If lVal > (Width / mlPASSES) Then
                lVal = lWidthSpans(X - 1) + CInt((Width / mlPASSES) * (X / mlPASSES))
            End If
            lWidthSpans(X) = lVal
        Next X
        ReDim lHeightSpans(mlPASSES - 1)
        For X = 0 To mlPASSES - 1
            lVal = CInt(2 ^ X)
            If lVal > (Height / mlPASSES) Then
                lVal = lHeightSpans(X - 1) + CInt((Height / mlPASSES) * (X / mlPASSES))
            End If
            lHeightSpans(X) = lVal
        Next X

        ReDim HeightMap(Width * Height)
        ReDim HM_Type(Width * Height)

        'Get our map type
        'MapType = -1
        'lTemp = Int(Rnd() * 100) + 1
        'Select Case lTemp
        '    Case Is < me_Planet_Type.ePT_Acid
        '        MapType = me_Planet_Type.ePT_Acid
        '    Case Is < me_Planet_Type.ePT_Barren
        '        MapType = me_Planet_Type.ePT_Barren
        '    Case Is < me_Planet_Type.ePT_GeoPlastic
        '        MapType = me_Planet_Type.ePT_GeoPlastic
        '    Case Is < me_Planet_Type.ePT_Desert
        '        MapType = me_Planet_Type.ePT_Desert
        '    Case Is < me_Planet_Type.ePT_Adaptable
        '        MapType = me_Planet_Type.ePT_Adaptable
        '    Case Is < me_Planet_Type.ePT_Tundra
        '        MapType = me_Planet_Type.ePT_Tundra
        '    Case Is < me_Planet_Type.ePT_Terran
        '        MapType = me_Planet_Type.ePT_Terran
        '    Case Else
        '        MapType = me_Planet_Type.ePT_Waterworld
        'End Select

        lTemp = CInt(Int(Rnd() * 100) + 1)


        'Get our water percentage based on map type
        lWaterPerc = GetWaterPerc(MapType)

        'Now, set our values to water unless there is none
        If lWaterPerc = 0 Then
            'set em all to land
            WaterHeight = 0
            For X = 0 To HeightMap.Length - 1
                HeightMap(X) = 100
                HM_Type(X) = 1
            Next X
        Else
            For X = 0 To HeightMap.Length - 1
                HeightMap(X) = 0
                HM_Type(X) = 0
            Next X

            'Base the waterheight off of the waterperc, if our water perc is practically nothing, then
            '  we don't want waterheight to be 256 :)
            'WaterHeight = CByte(Int(Rnd() * (lWaterPerc + 20)) + 1)

            If MapType = PlanetType.eTerran Then
                WaterHeight = 70
            ElseIf MapType = PlanetType.eWaterWorld Then
                WaterHeight = 160
            Else : WaterHeight = 40
            End If

            'Generate landmasses
            bDone = False
            While bDone = False
                lStartX = CInt(Int(Rnd() * Width) + 1)
                lStartY = CInt(Int(Rnd() * Height) + 1)
                X = lStartX
                Y = lStartY

                bLandmassLoop = False
                While bLandmassLoop = False
                    'Get our next movement...
                    Select Case Int(Rnd() * 8) + 1
                        Case 1 : Y -= 1
                        Case 2 : Y -= 1 : X += 1
                        Case 3 : X += 1
                        Case 4 : Y += 1 : X += 1
                        Case 5 : Y += 1
                        Case 6 : Y += 1 : X -= 1
                        Case 7 : X -= 1
                        Case 8 : X -= 1 : Y -= 1
                    End Select
                    'Validate ranges
                    If X < 0 Then X = Width - 1
                    If X >= Width Then X = 0
                    If Y < 0 Then Y = 1
                    If Y >= Height Then Y = Height - 2

                    lIdx = (Y * Width) + X
                    If HM_Type(lIdx) = 0 Then
                        HeightMap(lIdx) = CByte(WaterHeight + 32)      'to ensure that we are way above water
                        HM_Type(lIdx) = 1
                        lTotalLand += 1
                    End If

                    'Ok, this fixes a very nasty rounding bug that occurs on 32-bit systems.
                    ' basically, it would calculate the equation as being equal because one value
                    ' would be stored in memory identically to another. Therefore, we do some voodoo here
                    ' to see if the rounding error is happening....
                    'Dim dblVal1 As Double = (CDbl(lTotalLand) / CDbl(Width * Height)) 'get our errored value
                    'Dim lCompareVal As Int32 = 0
                    'If (dblVal1 * 100) > (100 - lWaterPerc) - 4 Then                ' are we close enough to test, if not, don't waste time
                    '    Dim sPerc1 As String = dblVal1.ToString()                   'store the errored value as a string
                    '    If sPerc1.Contains("E") = True Then                         'if the string contains an E, we're not even close
                    '        dblVal1 = 0
                    '        lCompareVal = 0
                    '    Else
                    '        sPerc1 = Mid$(sPerc1, 1, 5)                             'Ok, lop off some digits
                    '        lCompareVal = CInt(Val(sPerc1) * 100)                             'now, multiply the results accordingly and we have a valid test
                    '    End If
                    'Else : lCompareVal = CInt(dblVal1 * 100)
                    'End If
                    Dim lLookupIdx As Int32 = 100 - lWaterPerc

                    If X = lStartX AndAlso Y = lStartY Then
                        bLandmassLoop = True
                    ElseIf lTotalLand >= mlPercLookup(lLookupIdx) Then
                        'ElseIf lCompareVal >= 100 - lWaterPerc Then 'ElseIf CInt(Math.Floor((lTotalLand / (Width * Height)) * 100)) >= 100 - lWaterPerc Then
                        bLandmassLoop = True
                        bDone = True
                        'goUILib.AddNotification(Me.mlSeed & ": " & lTotalLand, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                    'MSC - 10/10/08 - remarked this out and added the above to change the double to an int32
                    'Dim dblVal1 As Double = (CDbl(lTotalLand) / CDbl(Width * Height)) 'get our errored value
                    'If (dblVal1 * 100) > (100 - lWaterPerc) - 2 Then                ' are we close enough to test, if not, don't waste time
                    '    Dim sPerc1 As String = dblVal1.ToString()                   'store the errored value as a string
                    '    If sPerc1.Contains("E") = True Then                         'if the string contains an E, we're not even close
                    '        dblVal1 = 0
                    '    Else
                    '        sPerc1 = Mid$(sPerc1, 1, 5)                             'Ok, lop off some digits
                    '        dblVal1 = Val(sPerc1) * 100                             'now, multiply the results accordingly and we have a valid test
                    '    End If
                    'End If

                    'If X = lStartX AndAlso Y = lStartY Then
                    '    bLandmassLoop = True
                    'ElseIf dblVal1 >= 100 - lWaterPerc Then 'ElseIf CInt(Math.Floor((lTotalLand / (Width * Height)) * 100)) >= 100 - lWaterPerc Then
                    '    bLandmassLoop = True
                    '    bDone = True
                    '    'goUILib.AddNotification(Me.mlSeed & ": " & lTotalLand, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    'End If
                End While
            End While
        End If

        'Generate the terrain on the map...
        For lPass = mlPASSES - 1 To 0 Step -1
            For X = 0 To Width - 1 Step lWidthSpans(lPass)
                For Y = 0 To Height - 1 Step lHeightSpans(lPass)
                    lVal = CInt(Int(Rnd() * 255) + 1)

                    lSubXMax = X + lWidthSpans(lPass) - 1
                    If lSubXMax > Width Then lSubXMax = Width - 1
                    lSubYMax = Y + lHeightSpans(lPass) - 1
                    If lSubYMax > Height Then lSubYMax = Height - 1

                    For lSubX = X To lSubXMax
                        For lSubY = Y To lSubYMax
                            lIdx = lSubY * Width + lSubX
                            If HM_Type(lIdx) = 0 Then
                                lTemp = CInt((lVal / 255) * WaterHeight)
                            Else
                                lTemp = lVal
                            End If
                            HeightMap(lIdx) = SmoothTerrainVal(lSubX, lSubY, lTemp)
                        Next lSubY
                    Next lSubX
                Next Y
            Next X
        Next lPass

        'Now, assuming we are lunar...
        If MapType = PlanetType.eBarren Then
            DoLunar()
        End If

        'Now, apply final filter to smooth everything over
        For lPass = 1 To 0 Step -1
            For X = 0 To Width - 1 Step lWidthSpans(lPass)
                For Y = 0 To Height - 1 Step lHeightSpans(lPass)
                    lSubXMax = X + lWidthSpans(lPass) - 1
                    If lSubXMax > Width Then lSubXMax = Width - 1
                    lSubYMax = Y + lHeightSpans(lPass) - 1
                    If lSubYMax > Height Then lSubYMax = Height - 1

                    For lSubX = X To lSubXMax
                        For lSubY = Y To lSubYMax
                            lIdx = lSubY * Width + lSubX

                            lTemp = HeightMap(lIdx)
                            HeightMap(lIdx) = SmoothTerrainVal(lSubX, lSubY, lTemp)
                        Next lSubY
                    Next lSubX
                Next Y
            Next X
        Next lPass

        If MapType = PlanetType.eGeoPlastic OrElse MapType = PlanetType.eWaterWorld Then
            CreateVolcanoes()
        End If

        'Now, accentuate
        lTemp = 0 : lVal = 0
        For X = 0 To Width - 1
            For Y = 0 To Height - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lTemp Then lTemp = HeightMap(lIdx)
                If HeightMap(lIdx) > WaterHeight Then lVal += HeightMap(lIdx)
            Next Y
        Next X
        lVal = lVal \ (Width * Height) 'CInt(lVal / (Width * Height))
        'lVal = lVal + CInt((lTemp - lVal) / 1.8)        'what is 1.8?
        lVal += CInt((lTemp - lVal) / 1.8F)
        For X = 0 To Width - 1
            For Y = 0 To Height - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lVal Then HeightMap(lIdx) = CByte(HeightMap(lIdx) * (255 / lTemp))
            Next Y
        Next X

        If MapType = PlanetType.eBarren Then
            For Y = 0 To Height - 1
                For X = 0 To Width - 1
                    lIdx = Y * Width + X
                    lTemp = HeightMap(lIdx)
                    lTemp += (CInt(Rnd() * 20) - 10)
                    If lTemp < 0 Then lTemp = 0
                    If lTemp > 255 Then lTemp = 255
                    HeightMap(lIdx) = CByte(lTemp)
                Next X
            Next Y
            MaximizeTerrain()
        ElseIf MapType = PlanetType.eDesert Then
            DoDesert()

            'and an additional smooth over
            For X = 0 To Width - 1
                For Y = 0 To Height - 1
                    lIdx = Y * Width + X
                    HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
                Next Y
            Next X
        ElseIf MapType = PlanetType.eAcidic Then
            DoAcidPlateaus()
        ElseIf MapType = PlanetType.eTerran Then
            DoPeakAccents()
        End If

        'One last soften...
        For X = 0 To Width - 1
            For Y = 0 To Height - 1
                lIdx = Y * Width + X
                HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
            Next Y
        Next X

        If MapType = PlanetType.eGeoPlastic Then
            DoGeoPlastic()
        ElseIf MapType = PlanetType.eAcidic Then
            DoAcidic()
        ElseIf MapType = PlanetType.eWaterWorld Then
            If mptVolcanoes Is Nothing = False Then
                For X = 0 To mptVolcanoes.GetUpperBound(0)
                    lIdx = mptVolcanoes(X).Y * Width + mptVolcanoes(X).X
                    HeightMap(lIdx) = 0
                Next
            End If
            MaximizeTerrain()
        ElseIf MapType = PlanetType.eTerran Then
            MaximizeTerrain()
        End If

        'To ensure map wrapping lines up perfectly
        For Y = 0 To TerrainClass.Height - 1
            lIdx = (Y * TerrainClass.Width)
            Dim lOtherIdx As Int32 = (Y * TerrainClass.Width) + (TerrainClass.Width - 1)
            HeightMap(lOtherIdx) = HeightMap(lIdx)
        Next Y

        mbHMReady = True

    End Sub
    Private mptVolcanoes() As Point
    Private Sub CreateVolcanoes()
        Dim lCnt As Int32 '= CInt(Rnd() * 5) + 1

        If MapType = PlanetType.eWaterWorld Then
            If Rnd() * 100 > 60 Then Return

            ReDim mptVolcanoes(0)

            Dim lYRng As Int32 = Height \ 2
            Dim lYOffset As Int32 = Height \ 4
            Dim lIdx As Int32

            Dim lVolX As Int32
            Dim lVolY As Int32

            'Get our max height
            Dim lVolcanoVal As Int32
            For Y As Int32 = 0 To Height - 1
                For X As Int32 = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > lVolcanoVal Then
                        lVolcanoVal = HeightMap(lIdx)
                        lVolX = X
                        lVolY = Y
                    End If
                Next X
            Next Y
            lVolcanoVal += 5
            If lVolcanoVal > 255 Then lVolcanoVal = 255

            mptVolcanoes(0).X = lVolX
            mptVolcanoes(0).Y = lVolY

            Dim lSize As Int32 = CInt(Rnd() * 4) + 6
            For lLocY As Int32 = lVolY - lSize To lVolY + lSize
                For lLocX As Int32 = lVolX - lSize To lVolX + lSize

                    Dim lTempX As Int32 = lLocX
                    Dim lTempY As Int32 = lLocY
                    If lTempX > Width - 1 Then lTempX -= Width
                    If lTempX < 0 Then lTempX += Width
                    If lTempY > Height - 1 Then lTempY -= Height
                    If lTempY < 0 Then lTempY += Height

                    lIdx = lTempY * Width + lTempX

                    If lLocX <> lVolX OrElse lLocY <> lVolY Then
                        If HeightMap(lIdx) > WaterHeight Then
                            If (Math.Abs(lLocX - lVolX) = 1 AndAlso Math.Abs(lLocY - lVolY) < 2) OrElse (Math.Abs(lLocY - lVolY) = 1 AndAlso Math.Abs(lLocX - lVolX) < 2) Then
                                HeightMap(lIdx) = CByte(lVolcanoVal)
                            Else
                                Dim lXVal As Int32 = lLocX - lVolX
                                lXVal *= lXVal
                                Dim lYVal As Int32 = lLocY - lVolY
                                lYVal *= lYVal

                                Dim fTotalVal As Single = CSng(Math.Sqrt(lXVal + lYVal))
                                Dim lOriginal As Int32 = HeightMap(lIdx)
                                Dim lVal As Int32 = WaterHeight + CInt((lVolcanoVal - WaterHeight) * (1.0F - (fTotalVal / (lSize * 3))))

                                If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                If lVal > lVolcanoVal Then lVal = lVolcanoVal
                                If lVal < lOriginal Then lVal = lOriginal

                                HeightMap(lIdx) = CByte(lVal)
                            End If
                        End If
                    Else : HeightMap(lIdx) = CByte(lVolcanoVal)
                    End If
                Next lLocX
            Next lLocY

        Else
            lCnt = CInt(Rnd() * 5) + 1
            ReDim mptVolcanoes(lCnt - 1)

            Dim lYRng As Int32 = Height \ 2
            Dim lYOffset As Int32 = Height \ 4
            Dim lIdx As Int32

            'Get our max height
            Dim lVolcanoVal As Int32
            For Y As Int32 = 0 To Height - 1
                For X As Int32 = 0 To Width - 1
                    lIdx = Y * Width + X
                    If HeightMap(lIdx) > lVolcanoVal Then lVolcanoVal = HeightMap(lIdx)
                Next X
            Next Y
            lVolcanoVal += 5
            If lVolcanoVal > 255 Then lVolcanoVal = 255

            While lCnt > 0
                Dim Y As Int32 = CInt(Rnd() * lYRng) + lYOffset
                Dim X As Int32 = CInt(Rnd() * (Width - 1))

                If HeightMap(Y * Width + X) > WaterHeight Then
                    mptVolcanoes(lCnt - 1).X = X
                    mptVolcanoes(lCnt - 1).Y = Y
                    lCnt -= 1

                    Dim lSize As Int32 = CInt(Rnd() * 4) + 6
                    For lLocY As Int32 = Y - lSize To Y + lSize
                        For lLocX As Int32 = X - lSize To X + lSize

                            Dim lTempX As Int32 = lLocX
                            Dim lTempY As Int32 = lLocY
                            If lTempX > Width - 1 Then lTempX -= Width
                            If lTempX < 0 Then lTempX += Width
                            If lTempY > Height - 1 Then lTempY -= Height
                            If lTempY < 0 Then lTempY += Height

                            lIdx = lTempY * Width + lTempX

                            If lLocX <> X OrElse lLocY <> Y Then
                                If HeightMap(lIdx) > WaterHeight Then
                                    If (Math.Abs(lLocX - X) = 1 AndAlso Math.Abs(lLocY - Y) < 2) OrElse (Math.Abs(lLocY - Y) = 1 AndAlso Math.Abs(lLocX - X) < 2) Then
                                        HeightMap(lIdx) = CByte(lVolcanoVal)
                                    Else
                                        Dim lXVal As Int32 = lLocX - X
                                        lXVal *= lXVal
                                        Dim lYVal As Int32 = lLocY - Y
                                        lYVal *= lYVal

                                        Dim fTotalVal As Single = CSng(Math.Sqrt(lXVal + lYVal))
                                        Dim lOriginal As Int32 = HeightMap(lIdx)
                                        Dim lVal As Int32 = WaterHeight + CInt((lVolcanoVal - WaterHeight) * (1.0F - (fTotalVal / (lSize * 3))))

                                        If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                        If lVal > lVolcanoVal Then lVal = lVolcanoVal
                                        If lVal < lOriginal Then lVal = lOriginal

                                        HeightMap(lIdx) = CByte(lVal)
                                    End If
                                End If
                            Else : HeightMap(lIdx) = CByte(lVolcanoVal)
                            End If
                        Next lLocX
                    Next lLocY
                End If

            End While
        End If

    End Sub

    Private Sub DoPeakAccents()
        Dim X As Int32
        Dim Y As Int32
        Dim lIdx As Int32
        Dim ptPeaks() As Point = Nothing
        Dim lPeakHt() As Int32 = Nothing
        Dim lPeakUB As Int32 = -1
        Dim lMinPeakHt As Int32 = 0

        For Y = 0 To Height - 1
            For X = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lMinPeakHt Then lMinPeakHt = HeightMap(lIdx)
            Next X
        Next Y
        lMinPeakHt -= CInt((Rnd() * 15) + 5)

        For Y = 0 To Height - 1
            For X = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lMinPeakHt Then
                    Dim bSkip As Boolean = False
                    For lTmpIdx As Int32 = 0 To lPeakUB
                        If Math.Abs(ptPeaks(lTmpIdx).X - X) < 3 AndAlso Math.Abs(ptPeaks(lTmpIdx).Y - Y) < 3 Then
                            If lPeakHt(lTmpIdx) < HeightMap(lIdx) Then
                                ptPeaks(lTmpIdx).X = X
                                ptPeaks(lTmpIdx).Y = Y
                                lPeakHt(lTmpIdx) = HeightMap(lIdx)
                            End If
                            bSkip = True
                        End If
                    Next lTmpIdx

                    If bSkip = False Then
                        lPeakUB += 1
                        ReDim Preserve ptPeaks(lPeakUB)
                        ReDim Preserve lPeakHt(lPeakUB)
                        ptPeaks(lPeakUB).X = X
                        ptPeaks(lPeakUB).Y = Y
                        lPeakHt(lPeakUB) = HeightMap(lIdx)
                    End If
                End If
            Next X
        Next Y

        Dim mfMults(5) As Single
        mfMults(0) = 1.0F
        mfMults(1) = 0.667F
        mfMults(2) = 0.475F
        mfMults(3) = 0.365F
        mfMults(4) = 0.304F
        mfMults(5) = 0.276F

        'Now, check our UB
        If lPeakUB <> -1 Then
            For lPkIdx As Int32 = 0 To lPeakUB

                Dim lVolX As Int32 = ptPeaks(lPkIdx).X
                Dim lVolY As Int32 = ptPeaks(lPkIdx).Y

                For lLocY As Int32 = lVolY - 6 To lVolY + 6
                    For lLocX As Int32 = lVolX - 6 To lVolX + 6

                        Dim lTempX As Int32 = lLocX
                        Dim lTempY As Int32 = lLocY
                        If lTempX > Width - 1 Then lTempX -= Width
                        If lTempX < 0 Then lTempX += Width
                        If lTempY > Height - 1 Then lTempY -= Height
                        If lTempY < 0 Then lTempY += Height

                        lIdx = lTempY * Width + lTempX

                        If lLocX <> lVolX OrElse lLocY <> lVolY Then
                            If HeightMap(lIdx) > WaterHeight Then
                                If (Math.Abs(lLocX - lVolX) = 1 AndAlso Math.Abs(lLocY - lVolY) < 2) OrElse (Math.Abs(lLocY - lVolY) = 1 AndAlso Math.Abs(lLocX - lVolX) < 2) Then
                                    HeightMap(lIdx) = 255
                                Else
                                    Dim lXVal As Int32 = lLocX - lVolX
                                    lXVal *= lXVal
                                    Dim lYVal As Int32 = lLocY - lVolY
                                    lYVal *= lYVal

                                    Dim fTotalVal As Single = lXVal + lYVal 'CSng(Math.Sqrt(lXVal + lYVal))
                                    Dim lOriginal As Int32 = HeightMap(lIdx)
                                    Dim lVal As Int32 = WaterHeight + CInt((255 - WaterHeight) * (1.0F - (fTotalVal / 18.0F)))
                                    If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                    If lVal > 255 Then lVal = 255
                                    If lVal < lOriginal \ 2 Then lVal = lOriginal \ 2

                                    HeightMap(lIdx) = CByte(lVal)
                                End If
                            End If
                        Else : HeightMap(lIdx) = CByte(255)
                        End If
                    Next lLocX
                Next lLocY
            Next lPkIdx

        End If
    End Sub

    Private Sub DoGeoPlastic()
        'GeoPlastic - 
        Dim lIdx As Int32
        Dim lTemp As Int32
        Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

        'short dropoff cliffs where rivers of lava flow. 
        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight + 10 Then
                    lTemp = CInt(HeightMap(lIdx)) + 10
                Else : lTemp = 0
                End If
                If lTemp > 255 Then lTemp = 255
                HeightMap(lIdx) = CByte(lTemp)
            Next X
        Next Y

        'Magma regions (water's beach area). 
        'Mountainous regions with no sharp inclines (rolling but with peaks). 

        '"Rips" in the ground for steam vents.
        Dim lCnt As Int32 = CInt(Rnd() * 10) + 15
        While lCnt <> 0
            Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
            Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight AndAlso HeightMap(lIdx) < CInt(lAboveWaterHeight * 0.6F) Then
                Dim lSize As Int32 = CInt(Rnd() * 10) + 10
                Dim lDirX As Int32
                Dim lDirZ As Int32

                If Rnd() * 100 < 50 Then lDirX = -1 Else lDirX = 1
                If Rnd() * 100 < 50 Then lDirZ = -1 Else lDirZ = 1

                lCnt -= 1

                While lSize > 0
                    If Rnd() * 100 < 50 Then
                        lSX += lDirX
                    Else : lSZ += lDirZ
                    End If

                    If lSX < 0 Then
                        lSX += Width
                    ElseIf lSX > Width - 1 Then
                        lSX -= Width
                    End If
                    If lSZ < 0 Then
                        lSZ += Height
                    ElseIf lSZ > Height - 1 Then
                        lSZ -= Height
                    End If

                    lIdx = lSZ * Width + lSX

                    HeightMap(lIdx) = 0

                    lSize -= 1
                End While 

            End If
        End While


        'Reset the volcano center's to 0
        If mptVolcanoes Is Nothing = False Then
            For X As Int32 = 0 To mptVolcanoes.GetUpperBound(0)
                lIdx = mptVolcanoes(X).Y * Width + mptVolcanoes(X).X
                HeightMap(lIdx) = 0
            Next
        End If

        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) < WaterHeight Then
                    If Rnd() * 100 < 1 Then
                        lIdx = 1        'needed to do this to force the compiler not to wax this method... we need to call rnd to sync with client
                    End If
                End If
            Next X
        Next Y

    End Sub

    Private Sub DoAcidPlateaus()
        'Ok, at this point, everything is accentuated... we are almost done
        Dim fPlateauStrength() As Single
        ReDim fPlateauStrength((Width * Height) - 1)

        'Ok, now... let's create our plateaus
        Dim lCnt As Int32 = CInt(Rnd() * 3) + 3
        Dim lIdx As Int32 ''

        While lCnt > 0
			Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
			Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight Then
                lCnt -= 1

                Select Case CInt(Rnd() * 100)
                    Case Is < 50
                        'EAST / WEST
                        'determine which side is the high side
                        If CInt(Rnd() * 100) < 50 Then
                            'East
                            Dim lEndVal As Int32 = Math.Max(lSX - (CInt(Rnd() * 30) + 30), 0)
                            For X As Int32 = lSX To lEndVal Step -1    '0 step -1

                                Dim fHtMult As Single = 0
                                If lSX - lEndVal <> 0 Then fHtMult = CSng((X - lEndVal) / (lSX - lEndVal))

                                'For Y As Int32 = 0 To Height - 1
                                For Y As Int32 = Math.Max(lSZ - CInt(Rnd() * 30 + 15), 0) To Math.Min(Height - 1, lSZ + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next Y
                            Next X
                        Else
                            'West
                            Dim lEndVal As Int32 = Math.Min(Width - 1, lSX + (CInt(Rnd() * 30) + 30))
                            For X As Int32 = lSX To lEndVal 'Width - 1

                                Dim fHtMult As Single = 0
                                If lEndVal - lSX <> 0 Then fHtMult = CSng((X - lSX) / (lEndVal - lSX))

                                'For Y As Int32 = 0 To Height - 1
                                For Y As Int32 = Math.Max(lSZ - CInt(Rnd() * 30 + 15), 0) To Math.Min(Height - 1, lSZ + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next Y
                            Next X
                        End If
                    Case Else ' Is < 40
                        'NORTH / SOUTH
                        'determine which side is the high side
                        If CInt(Rnd() * 100) < 50 Then
                            'NORTH
                            Dim lEndVal As Int32 = Math.Max(lSZ - (CInt(Rnd() * 30) + 30), 0)
                            For Y As Int32 = lSZ To lEndVal Step -1 '0 Step -1
                                Dim fHtMult As Single = 0
                                If lSZ - lEndVal <> 0 Then fHtMult = CSng((Y - lEndVal) / (lSZ - lEndVal))
                                'For X As Int32 = 0 To Width - 1
                                For X As Int32 = Math.Max(lSX - CInt(Rnd() * 30 + 15), 0) To Math.Min(Width - 1, lSX + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next X
                            Next Y
                        Else
                            'SOUTH
                            Dim lEndVal As Int32 = Math.Min(Height - 1, lSZ + (CInt(Rnd() * 30) + 30))
                            For Y As Int32 = lSZ To lEndVal 'Height - 1
                                Dim fHtMult As Single = 0
                                If lEndVal - lSZ <> 0 Then fHtMult = CSng((Y - lSZ) / (lEndVal - lSZ))

                                'For X As Int32 = 0 To Width - 1
                                For X As Int32 = Math.Max(lSX - CInt(Rnd() * 30 + 15), 0) To Math.Min(Width - 1, lSX + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next X
                            Next Y
                        End If
                End Select

            End If
        End While

        'Now, go through all of our values
        Dim lTmpValWH As Int32 = 255 - WaterHeight
        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                Dim lVal As Int32 = HeightMap(lIdx)

                If lVal > WaterHeight Then
                    Dim fVal As Single = fPlateauStrength(lIdx) * lTmpValWH

                    'Now... I want to affect this by fVal
                    fVal = ((lVal + fVal) / 2.0F) + WaterHeight
                    If fVal > 255 Then
                        HeightMap(lIdx) = 255
                    ElseIf fVal < WaterHeight + 1 Then
                        HeightMap(lIdx) = CByte(WaterHeight + 1)
                    Else
                        HeightMap(lIdx) = CByte(fVal)
                    End If

                End If
            Next X
        Next Y

        MaximizeTerrain()

    End Sub

    Private Sub MaximizeTerrain()
        Dim lIdx As Int32

        Dim lMinVal As Int32 = 255
        Dim lMaxVal As Int32 = 0

        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight Then
                    If HeightMap(lIdx) > lMaxVal Then lMaxVal = HeightMap(lIdx)
                    If HeightMap(lIdx) < lMinVal Then lMinVal = HeightMap(lIdx)
                End If
            Next X
        Next Y

        Dim lDiff As Int32 = lMaxVal - lMinVal
        Dim lDesiredDiff As Int32 = 255 - (WaterHeight + 1)

        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight Then
                    Dim lVal As Int32 = HeightMap(lIdx)
                    lVal = CInt(((lVal - lMinVal) / lDiff) * lDesiredDiff) + WaterHeight + 1
                    If lVal < WaterHeight + 1 Then lVal = WaterHeight + 1
                    If lVal > 255 Then lVal = 255
                    HeightMap(lIdx) = CByte(lVal)
                End If
            Next X
        Next Y
    End Sub

    Private Sub DoAcidic()
        Dim lIdx As Int32
        Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

        'Deep chasms with acidic rivers flowing through them. 
        Dim lCnt As Int32 = CInt(Rnd() * 10) + 15
        While lCnt <> 0
            Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
            Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight AndAlso HeightMap(lIdx) < CInt(lAboveWaterHeight * 0.6F) Then
                Dim lSize As Int32 = CInt(Rnd() * 10) + 10
                Dim lDirX As Int32
                Dim lDirZ As Int32

                If Rnd() * 100 < 50 Then lDirX = -1 Else lDirX = 1
                If Rnd() * 100 < 50 Then lDirZ = -1 Else lDirZ = 1

                lCnt -= 1

                While lSize > 0

                    Dim bXChg As Boolean = False
                    If Rnd() * 100 < 50 Then
                        lSX += lDirX
                        bXChg = True
                    Else : lSZ += lDirZ
                    End If


                    If lSX < 0 Then
                        lSX += Width
                    ElseIf lSX > Width - 1 Then
                        lSX -= Width
                    End If
                    If lSZ < 0 Then
                        lSZ += Height
                    ElseIf lSZ > Height - 1 Then
                        lSZ -= Height
                    End If

                    lIdx = lSZ * Width + lSX
                    HeightMap(lIdx) = 0

                    If Rnd() * 100 < 35 Then
                        Dim lTempX As Int32 = lSX
                        Dim lTempZ As Int32 = lSZ

                        If bXChg = True Then
                            lTempZ += lDirZ
                            If lTempZ < 0 Then
                                lTempZ += Height
                            ElseIf lTempZ > Height - 1 Then
                                lTempZ -= Height
                            End If
                        Else
                            lTempX += lDirX
                            If lTempX < 0 Then
                                lTempX += Width
                            ElseIf lTempX > Width - 1 Then
                                lTempX -= Width
                            End If
                        End If
                        lIdx = lTempZ * Width + lTempX
                        HeightMap(lIdx) = 0
                    End If

                    lSize -= 1
                End While

            End If
        End While

        'One last soften...
        For X As Int32 = 0 To Width - 1
            For Y As Int32 = 0 To Height - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight Then
                    HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
                End If
            Next Y
        Next X
    End Sub

    Private Sub DoDesert()
        Dim lIdx As Int32

        'smooth mountain regions
        Dim lCnt As Int32 = CInt(Rnd() * 5) + 3
        While lCnt > 0
            Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
            Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight Then
                lCnt -= 1

                Dim lSizeW As Int32 = CInt(Rnd() * 40) + 15
                Dim lHalfSizeW As Int32 = lSizeW \ 2
                Dim lSizeH As Int32 = CInt(Rnd() * 40) + 15
                Dim lHalfSizeH As Int32 = lSizeH \ 2

                For Y As Int32 = -(CInt(lHalfSizeH * Math.Max(Rnd, 0.3))) To (CInt(lHalfSizeH * Math.Max(Rnd, 0.3)))
                    For X As Int32 = -(CInt(lHalfSizeW * Math.Max(Rnd, 0.3))) To (CInt(lHalfSizeW * Math.Max(Rnd, 0.3)))

                        Dim lMapX As Int32 = lSX + X
                        Dim lMapY As Int32 = lSZ + Y

                        If lMapX < 0 Then lMapX += Width
                        If lMapX > Width - 1 Then lMapX -= Width
                        If lMapY < 0 Then lMapY += Height
                        If lMapY > Height - 1 Then lMapY -= Height

                        lIdx = lMapY * Width + lMapX

                        If HeightMap(lIdx) > WaterHeight Then
                            Dim lTemp As Int32 = HeightMap(lIdx)
                            lTemp += CInt(Rnd() * 20) + 100
                            If lTemp > 255 Then lTemp = 255
                            If lTemp < 0 Then lTemp = 0

                            HeightMap(lIdx) = CByte(lTemp)
                        End If
                    Next X
                Next Y

                'Now, add nobs to the top and bottoms
                For X As Int32 = -lHalfSizeW To lHalfSizeW
                    For Y As Int32 = CInt(Rnd() * 10) To 0 Step -1
                        Dim lMapX As Int32 = lSX + X
                        Dim lMapY As Int32 = lSZ - Y - lHalfSizeH

                        If lMapX < 0 Then lMapX += Width
                        If lMapX > Width - 1 Then lMapX -= Width
                        If lMapY < 0 Then lMapY += Height
                        If lMapY > Height - 1 Then lMapY -= Height

                        lIdx = lMapY * Width + lMapX

                        If HeightMap(lIdx) > WaterHeight Then
                            Dim lTemp As Int32 = HeightMap(lIdx)
                            lTemp += CInt(Rnd() * 20) + 100
                            If lTemp > 255 Then lTemp = 255
                            If lTemp < 0 Then lTemp = 0

                            HeightMap(lIdx) = CByte(lTemp)
                        End If
                    Next Y

                    For Y As Int32 = CInt(Rnd() * 10) To 0 Step -1
                        Dim lMapX As Int32 = lSX + X
                        Dim lMapY As Int32 = lSZ + Y + lHalfSizeH

                        If lMapX < 0 Then lMapX += Width
                        If lMapX > Width - 1 Then lMapX -= Width
                        If lMapY < 0 Then lMapY += Height
                        If lMapY > Height - 1 Then lMapY -= Height

                        lIdx = lMapY * Width + lMapX

                        If HeightMap(lIdx) > WaterHeight Then
                            Dim lTemp As Int32 = HeightMap(lIdx)
                            lTemp += CInt(Rnd() * 20) + 100
                            If lTemp > 255 Then lTemp = 255
                            If lTemp < 0 Then lTemp = 0

                            HeightMap(lIdx) = CByte(lTemp)
                        End If
                    Next Y
                Next X
            End If
        End While

    End Sub

    Private Sub DoLunar()
        Dim lPass As Int32
        Dim lPassMax As Int32
        Dim lStartX As Int32
        Dim lStartY As Int32
        Dim lSubX As Int32
        Dim lSubY As Int32
        Dim lSubXTmp As Int32
        Dim lSubYTmp As Int32
        Dim lSubYMax As Int32

        Dim lRadius As Int32
        Dim lImpactPower As Int32
        Dim lCraterBasinHeight As Int32
        Dim lDist As Int32

        'alrighty, let's do some crazy stuff...
        lPassMax = CInt(Rnd() * 10) + 20      'lpass is the number
        For lPass = 0 To lPassMax
            'ok, get ground-zero
            lStartX = CInt(Int(Rnd() * (Width - 1)))
            lStartY = CInt(Int(Rnd() * (Height - 1)))
            'and our radius
            lRadius = CInt(Int(Rnd() * 32) + 1)
            lImpactPower = CInt(Int(Rnd() * 10) + 1) 'impactpower
            lCraterBasinHeight = HeightMap(lStartY * Width + lStartX) - 60       'crater basin height
            If lCraterBasinHeight < 0 Then lCraterBasinHeight = 0

            lSubYMax = (lStartY + (lRadius + lImpactPower))
            If lSubYMax > Height - 1 Then lSubYMax = Height - 1

            For lSubY = (lStartY - (lRadius + lImpactPower)) To lSubYMax
                For lSubX = (lStartX - (lRadius + lImpactPower)) To (lStartX + (lRadius + lImpactPower))
                    lSubXTmp = lSubX
                    lSubYTmp = lSubY

                    If lSubYTmp < 0 Then lSubYTmp = 0
                    If lSubXTmp > Width - 1 Then lSubXTmp -= Width
                    If lSubXTmp < 0 Then lSubXTmp += Width

                    lDist = CInt(Math.Floor(Distance(lStartX, lStartY, lSubX, lSubY)))
                    Dim lTemp As Int32 = HeightMap(lSubYTmp * Width + lSubXTmp)

                    If lDist < lRadius Then
                        lTemp = CInt(lCraterBasinHeight + Int(Rnd() * 5) + 1)
                    ElseIf lDist = lRadius Then
                        lTemp = CInt(lCraterBasinHeight + 128)
                    ElseIf lDist - lRadius < lImpactPower Then
                        lTemp = CInt((lCraterBasinHeight + (12 * lImpactPower)) - (((lDist - lRadius) / lImpactPower) * lCraterBasinHeight))
                    End If
                    If lTemp < 0 Then lTemp = 0
                    If lTemp > 255 Then lTemp = 255
                    HeightMap(lSubYTmp * Width + lSubXTmp) = CByte(lTemp)
                Next lSubX
            Next lSubY
        Next lPass
    End Sub

    Private Function Distance(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
        Dim dX As Single
        Dim dY As Single

        dX = lX2 - lX1
        dY = lY2 - lY1
        dX = dX * dX
        dY = dY * dY
        Distance = CSng(Math.Sqrt(dX + dY))
    End Function

    Private Function GetWaterPerc(ByVal lType As Int32) As Int32
        Select Case lType
            Case PlanetType.eBarren
                Return 0
            Case PlanetType.eDesert
                Return CInt(Int(Rnd() * 15) + 1)
            Case PlanetType.eGeoPlastic
                Return CInt(Int(Rnd() * 40) + 15)
            Case PlanetType.eWaterWorld
                Return CInt(80 + (Int(Rnd() * 20) + 1))
            Case PlanetType.eAdaptable
                Return CInt(5 + CInt(Rnd() * 40))
            Case PlanetType.eTundra
                Return 20 + CInt(Rnd() * 40)
            Case PlanetType.eTerran
                Return 30 + CInt(Rnd() * 20)
            Case Else 'PlanetType.eTerran, PlanetType.eAcidic
                Return CInt(30 + (Int(Rnd() * 50) + 1))
        End Select
    End Function

    Private Function SmoothTerrainVal(ByVal X As Int32, ByVal Y As Int32, ByVal lVal As Int32) As Byte
        Dim fCorners As Single
        Dim fSides As Single
        Dim fCenter As Single
        Dim fTotal As Single

        Dim lBackX As Int32
        Dim lBackY As Int32
        Dim lForeX As Int32
        Dim lForeY As Int32

        If X = 0 Then
            lBackX = Width - 1
        Else : lBackX = X - 1
        End If
        If X = Width - 1 Then
            lForeX = 0
        Else : lForeX = X + 1
        End If
        If Y = 0 Then
            lBackY = 0
        Else : lBackY = Y - 1
        End If
        If Y = Height - 1 Then
            lForeY = Height - 2
        Else : lForeY = Y + 1
        End If

        fCorners = 0
        fCorners = fCorners + HeightMap((lBackY * Width) + lBackX) 'muTiles(lBackX, lBackY).lHeight
        fCorners = fCorners + HeightMap((lBackY * Width) + lForeX) 'muTiles(lForeX, lBackY).lHeight
        fCorners = fCorners + HeightMap((lForeY * Width) + lBackX) 'muTiles(lBackX, lForeY).lHeight
        fCorners = fCorners + HeightMap((lForeY * Width) + lForeX) 'muTiles(lForeX, lForeY).lHeight
        fCorners = fCorners / 16

        fSides = 0
        fSides = fSides + HeightMap((Y * Width) + lBackX) 'muTiles(lBackX, Y).lHeight
        fSides = fSides + HeightMap((Y * Width) + lForeX) 'muTiles(lForeX, Y).lHeight
        fSides = fSides + HeightMap((lBackY * Width) + X) 'muTiles(X, lBackY).lHeight
        fSides = fSides + HeightMap((lForeY * Width) + X) 'muTiles(X, lForeY).lHeight
        fSides = fSides / 8

        fCenter = lVal / 4.0F

        fTotal = fCorners + fSides + fCenter
        If fTotal < 0 Then fTotal = 0
        If fTotal > 255 Then fTotal = 255

        Return CByte(fTotal)
    End Function

    Public Sub PopulateData()
        'this sub takes the place of everything that I removed from Client, we need to do everything to get it ready...
        GenerateTerrain(mlSeed)
        ComputeNormals()
    End Sub

    'NOTE: DO NOT CALL THIS NORMALLY, THIS ROUTINE IS FOR SPEED MOD CALCULATIONS!!!
    'NOTE: Called from Planet.TerrainGradeBuildable
    'NOTE: Called from Planet.SetupStartLocGrid
    Public Function GetTerrainNormalEx(ByVal fVertX As Single, ByVal fVertY As Single) As Vector3
        Dim fTX As Single   'translated X 
        Dim fTZ As Single   'translated Z 
        Dim lCol As Int32
        Dim lRow As Int32

        Dim dZ As Single
        Dim dX As Single

        Dim vN1 As vector3
        Dim vN2 As Vector3
        Dim vN3 As Vector3
        Dim vN4 As Vector3

        Dim lIdx As Int32

        fTX = fVertX
        fTZ = fVertY

        lCol = CInt(Math.Floor(fTX))
        lRow = CInt(Math.Floor(fTZ))
        dX = fTX - lCol
        dZ = fTZ - lRow

        lIdx = (lRow * Width) + lCol
		If lIdx < 0 OrElse lIdx > HMNormals.Length - 1 Then vN1 = Vector3.Empty Else vN1 = HMNormals(lIdx)
        lIdx = (lRow * Width) + lCol + 1
		If lIdx < 0 OrElse lIdx > HMNormals.Length - 1 Then vN2 = Vector3.Empty Else vN2 = HMNormals(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol
		If lIdx < 0 OrElse lIdx > HMNormals.Length - 1 Then vN3 = Vector3.Empty Else vN3 = HMNormals(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol + 1
		If lIdx < 0 OrElse lIdx > HMNormals.Length - 1 Then vN4 = Vector3.Empty Else vN4 = HMNormals(lIdx)

        vN1.Multiply(1 - dX)
        vN2.Multiply(dX)
        Dim vA As Vector3 = Vector3.Add(vN1, vN2)
        vN3.Multiply(1 - dX)
        vN4.Multiply(dX)
        Dim vB As Vector3 = Vector3.Add(vN3, vN4)
        vA.Multiply(1 - dZ)
        vB.Multiply(dZ)
        Dim vF As Vector3 = Vector3.Add(vA, vB)
        vF.Normalize()

        Return vF
    End Function

    Private Sub ComputeNormals()
        Dim Z As Int32
        Dim X As Int32

        Dim vecX As Vector3 = New Vector3()
        Dim vecZ As Vector3 = New Vector3()
        Dim vecN As Vector3
 
        Dim lHeight As Int32
        Dim oVerts() As Vector3

        ReDim HMNormals(VertsTotal - 1)
        ReDim oVerts(VertsTotal - 1)

        For Z = 0 To Height - 1
            For X = 0 To Width - 1
                lHeight = HeightMap((Z * Height) + X)

                oVerts((Z * Height) + X) = New Vector3((X - (mlHalfWidth)) * CellSpacing, lHeight * ml_Y_Mult, (Z - (mlHalfHeight)) * CellSpacing)
            Next X
        Next Z

        For Z = 1 To QuadsZ - 1
            For X = 1 To QuadsX - 1
                vecX = Vector3.Subtract(oVerts(Z * Width + X + 1), oVerts(Z * Width + X - 1))
                vecZ = Vector3.Subtract(oVerts((Z + 1) * Width + X), oVerts((Z - 1) * Width + X))
                vecN = Vector3.CrossProduct(vecZ, vecX)
                vecN.Normalize()

                HMNormals(Z * Width + X) = vecN
            Next X
        Next Z

        vecX = Nothing
        vecZ = Nothing
        vecN = Nothing

        Erase oVerts

	End Sub

	Public Sub PopulateInstanceData()
		Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
		If sPath.EndsWith("\") = False Then sPath &= "\"
		Dim oFS As IO.FileStream = New IO.FileStream(sPath & "terr.txt", IO.FileMode.Open)
		Dim oRead As IO.StreamReader = New IO.StreamReader(oFS)

		WaterHeight = CByte(Val(oRead.ReadLine))
		ReDim HeightMap(Width * Height)
		For Y As Int32 = 0 To TerrainClass.Height - 1
			Dim sLine As String = oRead.ReadLine
			Dim sValues() As String = Split(sLine, ",")
			For X As Int32 = 0 To TerrainClass.Width - 1
				Dim lIdx As Int32 = (Y * TerrainClass.Width) + X
				HeightMap(lIdx) = CByte(Val(sValues(X)))
			Next X
		Next Y
		oRead.Close()
		oRead.Dispose()
		oFS.Close()
		oFS.Dispose()
		oRead = Nothing
		oFS = Nothing

		mbHMReady = True
		ComputeNormals()
	End Sub
End Class

