Option Strict On

Module GlobalVars
    Public Enum SpecialOp As Byte
        eNoSpecialOp = 0
        eBeginMiningOp
        eBeginConstruction

        eBeginDismantle
        eBeginRepair

        eJumpTarget
    End Enum

    Public goMovers() As Mover
    Public glMoverIdx(-1) As Int32      'objectIDs, be sure to check the objtypeids
    Public glMoverUB As Int32 = -1

    Public goMsgSys As MsgSystem
    Public glBoxOperatorID As Int32
    Public gsOperatorIP As String
    Public glOperatorPort As Int32

    'Public Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Public gfrmDisplayForm As Form1

    'Galaxy contains all
    Public goGalaxy As Galaxy

    Public goStarTypes() As StarType
    Public glStarTypeUB As Int32 = -1

    'Used for quick lookup of planets
    Public Structure ValPair
        Public lID As Int32
        Public lParentID As Int32
        Public oPlanetRef As Planet         'REFERENCE ONLY!!!
    End Structure
    Public goPlanetVP() As ValPair
    Public glPlanetVPUB As Int32 = -1

    Public Class ModelData
        Public lModelID As Int32
        Public lRectSize As Int32
        Public PlotPaths As Boolean
        Public bNaval As Boolean
    End Class

    Public goModels() As ModelData
    Public glModelUB As Int32 = -1

    Public Function ExecuteFleetRegroupRequest(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal lCurX As Int32, ByVal lCurZ As Int32, ByVal iCurAngle As Int16, ByVal lCurID As Int32, ByVal iCurTypeID As Int16, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16, ByVal lDestID As Int32, ByVal iDestTypeID As Int16, ByVal iModelID As Int16) As Byte()
        Dim colNewDests As Collection
        Dim yResp() As Byte
        'Dim X As Int32
        Dim lIdx As Int32
        Dim lFirstIdx As Int32 = -1

        colNewDests = GetDestList(lCurX, lCurZ, iCurAngle, lCurID, iCurTypeID, lDestX, lDestZ, iDestAngle, lDestID, iDestTypeID, Mover.GetPathLocs(iModelID))

        If colNewDests Is Nothing Then
            'return invalid dest
            ReDim yResp(7)
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
            Return yResp
        Else

            'Add one more dest
            Dim uDest As DestLoc
            With uDest
                .DestAngle = 0
                .DestX = lDestX
                .DestZ = lDestZ
                .DestAngle = iDestAngle
                .iDestTypeID = iDestTypeID
                .iSpecialOpTypeID = 0
                .lDestID = lDestID
                .lSpecialOpID = 0
                .yChangeEnvironment = ChangeEnvironmentType.eSystemToSystem
                .ySpecialOp = SpecialOp.eNoSpecialOp
            End With
            colNewDests.Add(uDest)

            lIdx = GetOrAddMover(lObjID, iObjTypeID, lCurX, lCurZ, iCurAngle, iModelID, lCurID, iCurTypeID)
            If lIdx < 0 Then Return Nothing

            goMovers(lIdx).colDests = colNewDests

            'Now, we kinda cheat here...
            yResp = goMovers(lIdx).ProcessStopMessage(lCurX, lCurZ, iCurAngle)

            If yResp.Length > 0 Then
                Return yResp
            Else
                'Ok, send invalid move
                goMovers(lIdx).bMoving = False
                ReDim yResp(7)
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                Return yResp
            End If
        End If

    End Function

    Public Function ExecuteMoveRequest(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal lCurX As Int32, ByVal lCurZ As Int32, ByVal iCurAngle As Int16, ByVal lCurID As Int32, ByVal iCurTypeID As Int16, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16, ByVal lDestID As Int32, ByVal iDestTypeID As Int16, ByVal iModelID As Int16) As Byte()
        Dim colNewDests As Collection
        Dim yResp() As Byte

        Dim lMoverIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lCurX, lCurZ, iCurAngle, iModelID, lCurID, iCurTypeID)
        If lMoverIdx < 0 Then Return Nothing
        Dim oMover As Mover = goMovers(lMoverIdx)
        If oMover Is Nothing Then Return Nothing

        If oMover.oFormation Is Nothing = False Then oMover.oFormation.DetachItem(oMover.ObjectID, oMover.ObjTypeID)
        oMover.oFormation = Nothing

        colNewDests = GetDestList(lCurX, lCurZ, iCurAngle, lCurID, iCurTypeID, lDestX, lDestZ, iDestAngle, lDestID, iDestTypeID, Mover.GetPathLocs(iModelID))

        If colNewDests Is Nothing Then
            'return invalid dest
            ReDim yResp(7)
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
            Return yResp
        Else
			'Ok, clear our current list
			oMover.colDests = Nothing
			oMover.colDests = colNewDests

            'Now, we kinda cheat here...
			yResp = oMover.ProcessStopMessage(lCurX, lCurZ, iCurAngle)

            If yResp.Length > 0 Then
                Return yResp
            Else
                'Ok, send invalid move
				oMover.bMoving = False
                ReDim yResp(7)
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                Return yResp
            End If
        End If
	End Function
	Public Function ExecuteMoveRequestNoLookup(ByRef oMover As Mover, ByVal lCurX As Int32, ByVal lCurZ As Int32, ByVal iCurAngle As Int16, ByVal lCurID As Int32, ByVal iCurTypeID As Int16, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16, ByVal lDestID As Int32, ByVal iDestTypeID As Int16, ByVal iModelID As Int16) As Byte()
		Dim colNewDests As Collection
		Dim yResp() As Byte

		If oMover Is Nothing Then Return Nothing

		colNewDests = GetDestList(lCurX, lCurZ, iCurAngle, lCurID, iCurTypeID, lDestX, lDestZ, iDestAngle, lDestID, iDestTypeID, Mover.GetPathLocs(iModelID))

		If colNewDests Is Nothing Then
			'return invalid dest
			ReDim yResp(7)
			System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
			oMover.GetGUIDAsString.CopyTo(yResp, 2)
			Return yResp
		Else
			'Ok, clear our current list
			oMover.colDests = Nothing
			oMover.colDests = colNewDests

			'Now, we kinda cheat here...
			yResp = oMover.ProcessStopMessage(lCurX, lCurZ, iCurAngle)

			If yResp.Length > 0 Then
				Return yResp
			Else
				'Ok, send invalid move
				oMover.bMoving = False
				ReDim yResp(7)
				System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
				oMover.GetGUIDAsString.CopyTo(yResp, 2)
				Return yResp
			End If
		End If
	End Function

    Public Function GetOrAddMover(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal lCurX As Int32, ByVal lCurZ As Int32, ByVal iCurAngle As Int16, ByVal iModelID As Int16, ByVal lCurID As Int32, ByVal iCurTypeID As Int16) As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIdx As Int32 = -1

        Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glMoverIdx(X) = lObjID Then
                If goMovers(X) Is Nothing Then
                    glMoverIdx(X) = -1
                    Continue For
                End If
                If goMovers(X).ObjTypeID = iObjTypeID Then
                    goMovers(X).lCurLocX = lCurX
                    goMovers(X).lCurLocZ = lCurZ
                    goMovers(X).lEnvirID = lCurID
                    goMovers(X).iEnvirTypeID = iCurTypeID
                    If lFirstIdx <> -1 Then glMoverIdx(lFirstIdx) = -1
                    Return X 'goMovers(X)
                End If
            ElseIf glMoverIdx(X) = -1 AndAlso lFirstIdx = -1 Then
                glMoverIdx(X) = -2
                lFirstIdx = X
            End If
        Next X

        If lIdx = -1 Then
            If lFirstIdx = -1 Then
                SyncLock glMoverIdx
                    lIdx = glMoverUB + 1
                    ReDim Preserve goMovers(lIdx)
                    ReDim Preserve glMoverIdx(lIdx)
                    glMoverIdx(lIdx) = -2
                    glMoverUB += 1
                End SyncLock
            Else
                lIdx = lFirstIdx
            End If

            Dim oMover As New Mover()
            With oMover
                .lGlobalIdx = lIdx
                .lCurAngle = iCurAngle
                .lCurLocX = lCurX
                .lCurLocZ = lCurZ
                .lCurLocY = 0
                .ObjectID = lObjID
                .ObjTypeID = iObjTypeID
                .iModelID = iModelID
                .lEnvirID = lCurID
                .iEnvirTypeID = iCurTypeID
                .rcArea = Mover.GetRectWidthHeight(.iModelID)
                .rcArea.X = .lCurLocX
                .rcArea.Y = .lCurLocZ
            End With
            'goMovers(lIdx) = New Mover()
            goMovers(lIdx) = oMover
            glMoverIdx(lIdx) = lObjID
        End If
        Return lIdx 'goMovers(lIdx)
    End Function

    Public Function ExecuteAddWaypointRequest(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal lCurX As Int32, ByVal lCurZ As Int32, ByVal iCurAngle As Int16, ByVal lCurID As Int32, ByVal iCurTypeID As Int16, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16, ByVal lDestID As Int32, ByVal iDestTypeID As Int16, ByVal iModelID As Int16) As Byte()
        Dim uDest As DestLoc
        Dim colNewDests As Collection
        Dim lIdx As Int32 = -1
        Dim yResp() As Byte

        'Dim X As Int32
        'Dim lFirstIdx As Int32 = -1
        'For X = 0 To glMoverUB
        '    If glMoverIdx(X) = lObjID Then
        '        If goMovers(X).ObjTypeID = iObjTypeID Then
        '            lIdx = X
        '            Exit For
        '        End If
        '    End If
        'Next X

        'If lIdx = -1 Then
        '    If lFirstIdx = -1 Then
        '        glMoverUB += 1
        '        ReDim Preserve goMovers(glMoverUB)
        '        ReDim Preserve glMoverIdx(glMoverUB)
        '        lIdx = glMoverUB
        '    Else : lIdx = lFirstIdx
        '    End If
        '    goMovers(lIdx) = New Mover()
        '    glMoverIdx(lIdx) = lObjID
        '    With goMovers(lIdx)
        '        .lGlobalIdx = lIdx
        '        .lCurAngle = iCurAngle
        '        .lCurLocX = lCurX
        '        .lCurLocZ = lCurZ
        '        .lCurLocY = 0
        '        .lEnvirID = lCurID
        '        .iEnvirTypeID = iCurTypeID
        '        .ObjectID = lObjID
        '        .ObjTypeID = iObjTypeID
        '        .iModelID = iModelID
        '        .rcArea = Mover.GetRectWidthHeight(.iModelID)
        '        .rcArea.X = .lCurLocX
        '        .rcArea.Y = .lCurLocZ
        '    End With
        'End If
        lIdx = GetOrAddMover(lObjID, iObjTypeID, lCurX, lCurZ, iCurAngle, iModelID, lCurID, iCurTypeID)
        If lIdx = -1 Then Return Nothing

        If goMovers(lIdx).colDests Is Nothing Then goMovers(lIdx).colDests = New Collection()

        'we bring in curx, curz, curangle in the event that the current dest list is empty
        If goMovers(lIdx).colDests.Count = 0 Then
            colNewDests = GetDestList(lCurX, lCurZ, iCurAngle, lCurID, iCurTypeID, lDestX, lDestZ, iDestAngle, lDestID, iDestTypeID, Mover.GetPathLocs(iModelID))
            If colNewDests Is Nothing Then
                'return invalid dest
                ReDim yResp(7)
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                Return yResp
            Else
                goMovers(lIdx).colDests = colNewDests
            End If
        Else
            ' go the last item
            uDest = CType(goMovers(lIdx).colDests.Item(goMovers(lIdx).colDests.Count), DestLoc)
            ' and get the dest list from that
            colNewDests = GetDestList(uDest.DestX, uDest.DestZ, uDest.DestAngle, uDest.lDestID, uDest.iDestTypeID, lDestX, lDestZ, iDestAngle, lDestID, iDestTypeID, Mover.GetPathLocs(iModelID))
            If colNewDests Is Nothing Then
                'return invalid dest
                ReDim yResp(7)
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                Return yResp
            Else
                'Now, move our colNewDests over to colDests
                For Each uDest In colNewDests
                    goMovers(lIdx).colDests.Add(uDest)
                Next
            End If

        End If
        'Now, we kinda cheat here...
        If goMovers(lIdx).bMoving = False Then
            yResp = goMovers(lIdx).ProcessStopMessage(lCurX, lCurZ, iCurAngle)

            If yResp.Length > 0 Then
                Return yResp
            Else
                'Ok, send invalid move
                ReDim yResp(7)
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                Return yResp
            End If
        End If
        Return Nothing
    End Function

    Public Function GetDestList(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iLocAngle As Int16, ByVal lLocID As Int32, ByVal iLocTypeID As Int16, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16, ByVal lDestID As Int32, ByVal iDestTypeID As Int16, ByVal yPathTerrain As Byte) As Collection
        Dim colDests As Collection
        Dim uDest As DestLoc

        Dim oLoc As Object = Nothing
        Dim oDest As Object = Nothing
        Dim X As Int32
        Dim lSysID As Int32
        Dim Y As Int32

        colDests = New Collection()

        'Ok, determine if there is anything we need to do from a change environment standpoint
        If lDestID <> lLocID OrElse iDestTypeID <> iLocTypeID Then
            'get our current object
            If iLocTypeID = ObjectType.ePlanet Then
                lSysID = -1
                For X = 0 To glPlanetVPUB
                    If goPlanetVP(X).lID = lLocID Then
                        lSysID = goPlanetVP(X).lParentID
                        Exit For
                    End If
                Next X

                'Now, go thru and find our system to get our object
                For X = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X).ObjectID = lSysID Then
                        For Y = 0 To goGalaxy.moSystems(X).PlanetUB
                            If goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lLocID Then
                                oLoc = goGalaxy.moSystems(X).moPlanets(Y)
                                Exit For
                            End If
                        Next Y
                        Exit For
                    End If
                Next X
            Else
                For X = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X).ObjectID = lLocID Then
                        oLoc = goGalaxy.moSystems(X)
                        Exit For
                    End If
                Next X
            End If

            'get our destination object
            If iDestTypeID = ObjectType.ePlanet Then
                lSysID = -1
                For X = 0 To glPlanetVPUB
                    If goPlanetVP(X).lID = lDestID Then
                        lSysID = goPlanetVP(X).lParentID
                        Exit For
                    End If
                Next X

                'Now, go thru and find our system to get our object
                For X = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X).ObjectID = lSysID Then
                        For Y = 0 To goGalaxy.moSystems(X).PlanetUB
                            If goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lDestID Then
                                oDest = goGalaxy.moSystems(X).moPlanets(Y)
                                Exit For
                            End If
                        Next Y
                        Exit For
                    End If
                Next X
            Else
                For X = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X).ObjectID = lDestID Then
                        oDest = goGalaxy.moSystems(X)
                        Exit For
                    End If
                Next X
            End If

            'Return false or something...
            If oDest Is Nothing OrElse oLoc Is Nothing Then Return Nothing

            'Ok, we have our objects, now, figure our route...
            If iLocTypeID = ObjectType.ePlanet Then
                If iDestTypeID = ObjectType.ePlanet Then
                    If CType(oLoc, Planet).ParentSystem.ObjectID = CType(oDest, Planet).ParentSystem.ObjectID Then
                        'Loc to Loc.Parent to Dest
                        PlanetToSystemMove(colDests, CType(oLoc, Planet), lLocX, lLocZ)
                        'get our last loc
                        uDest = CType(colDests(colDests.Count), DestLoc)
                        SystemToPlanetMove(colDests, CType(oDest, Planet), uDest.DestX, uDest.DestZ, lDestX, lDestZ)
                    Else
                        'MSC - 03/07/07 - Removed this because a FLEET is required to do a system to system move now
                        Return Nothing

                        'Loc to Loc.Parent to Dest.Parent to Dest
                        PlanetToSystemMove(colDests, CType(oLoc, Planet), lLocX, lLocZ)
                        SystemToSystemMove(colDests, CType(oLoc, Planet).ParentSystem, CType(oDest, Planet).ParentSystem)
                        'get our last loc
                        uDest = CType(colDests(colDests.Count), DestLoc)
                        SystemToPlanetMove(colDests, CType(oDest, Planet), uDest.DestX, uDest.DestZ, lDestX, lDestZ)
                    End If
                Else    'dest is system
                    If CType(oLoc, Planet).ParentSystem.ObjectID = CType(oDest, Base_GUID).ObjectID Then
                        'Loc to Dest
                        PlanetToSystemMove(colDests, CType(oLoc, Planet), lLocX, lLocZ)
                    Else
                        'MSC - 03/07/07 - Removed this because a FLEET is required to do a system to system move now
                        Return Nothing

                        'Loc to Loc.Parent to Dest
                        PlanetToSystemMove(colDests, CType(oLoc, Planet), lLocX, lLocZ)
                        SystemToSystemMove(colDests, CType(oLoc, Planet).ParentSystem, CType(oDest, SolarSystem))
                    End If
                End If
            ElseIf iDestTypeID = ObjectType.ePlanet Then    'loc is system
                If CType(oDest, Planet).ParentSystem.ObjectID = CType(oLoc, SolarSystem).ObjectID Then
                    'Loc to Dest
                    SystemToPlanetMove(colDests, CType(oDest, Planet), lLocX, lLocZ, lDestX, lDestZ)
                Else
                    'MSC - 03/07/07 - Removed this because a FLEET is required to do a system to system move now
                    Return Nothing

                    'Loc to Dest.Parent to Dest
                    SystemToSystemMove(colDests, CType(oLoc, SolarSystem), CType(oDest, Planet).ParentSystem)
                    'Get our last loc
                    uDest = CType(colDests(colDests.Count), DestLoc)
                    SystemToPlanetMove(colDests, CType(oDest, Planet), uDest.DestX, uDest.DestZ, lDestX, lDestZ)
                End If
            Else
                'MSC - 03/07/07 - Removed this because a FLEET is required to do a system to system move now
                Return Nothing

                'Loc to Dest
                SystemToSystemMove(colDests, CType(oLoc, SolarSystem), CType(oDest, SolarSystem))
            End If
        End If

        'Now, finally, do our move to dest in the final envir... this is where we do the pathfinding...
        If yPathTerrain <> 0 Then
            'Ok... pull up the planet object (if we are indeed on one...
            If iDestTypeID = ObjectType.ePlanet Then
                'Ok... now, get our planet
                Dim oPlanet As Planet = Nothing

                If lDestID >= 500000000 Then
                    oPlanet = Planet.GetInstancePlanet()
                Else
                    For X = 0 To glPlanetVPUB
                        If goPlanetVP(X).lID = lDestID Then
                            oPlanet = goPlanetVP(X).oPlanetRef
                            Exit For
                        End If
                    Next X
                End If

                If oPlanet Is Nothing = False Then

                    Dim lCellSpacing As Int32 = oPlanet.GetCellSpacing
                    Dim lSX As Int32 = (lLocX + oPlanet.lHalfExtent) \ lCellSpacing
                    Dim lSZ As Int32 = (lLocZ + oPlanet.lHalfExtent) \ lCellSpacing
                    Dim lEX As Int32 = (lDestX + oPlanet.lHalfExtent) \ lCellSpacing
                    Dim lEZ As Int32 = (lDestZ + oPlanet.lHalfExtent) \ lCellSpacing

                    If lSX = lEX AndAlso lSZ = lEZ Then
                        'Just add the dest directly
                        uDest.DestAngle = iDestAngle
                        uDest.DestX = lDestX
                        uDest.DestZ = lDestZ
                        uDest.iDestTypeID = iDestTypeID
                        uDest.lDestID = lDestID
                        uDest.yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
                        uDest.ySpecialOp = CByte(SpecialOp.eNoSpecialOp)
                        colDests.Add(uDest)
                    Else
                        Dim ptList() As System.Drawing.Point = Nothing
                        If yPathTerrain = 1 Then
                            ptList = oPlanet.oPather.GetPath(lSX, lSZ, iLocAngle, lEX, lEZ)
                            If ptList Is Nothing = False Then
                                'Next, we need to optimize that path for our uses...
                                ptList = oPlanet.oPather.OptimizePath(ptList, New Point(lSX, lSZ), False)
                            End If
                        ElseIf yPathTerrain = 2 Then
                            ptList = oPlanet.oPather.GetPathNaval(lSX, lSZ, iLocAngle, lEX, lEZ)
                            If ptList Is Nothing = False Then
                                'Next, we need to optimize that path for our uses...
                                ptList = oPlanet.oPather.OptimizePath(ptList, New Point(lSX, lSZ), True)
                            End If
                        End If

                        If ptList Is Nothing = False Then
                            'Set up the dest's defaults
                            With uDest
                                .iDestTypeID = iDestTypeID
                                .lDestID = lDestID
                                .ySpecialOp = CByte(SpecialOp.eNoSpecialOp)
                                .yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
                            End With

                            'Now... we need to transpose that loc list into our final loc list...
                            Dim lLastX As Int32 = lLocX
                            Dim lLastZ As Int32 = lLocZ
                            Dim ptTemp As System.Drawing.Point
                            Dim lHalfWidth As Int32 = TerrainClass.Width \ 2
                            Dim lTotalWidth As Int32 = TerrainClass.Width * oPlanet.GetCellSpacing()
                            For X = 1 To ptList.GetUpperBound(0) - 1
                                ptTemp.X = (ptList(X).X - lHalfWidth) * lCellSpacing
                                ptTemp.Y = (ptList(X).Y - lHalfWidth) * lCellSpacing
                                With uDest
                                    .DestAngle = CShort(LineAngleDegrees(lLastX, lLastZ, ptList(X).X, ptList(X).Y) * 10)
                                    .DestX = ptTemp.X
                                    .DestZ = ptTemp.Y
                                End With
                                lLastX = ptTemp.X
                                lLastZ = ptTemp.Y
                                colDests.Add(uDest)
                            Next X
                            ''If we are the only location in the list...
                            'If ptList.GetUpperBound(0) < 1 Then
                            '    'Take map wrap into consideration
                            '    If lDestX < lLocX Then
                            '        'Ok, normally, we would go left...
                            '        Dim lTmpDX As Int32 = lDestX + lTotalWidth
                            '        If Math.Abs(lTmpDX - lLocX) < Math.Abs(lLocX - lDestX) Then
                            '            lDestX = lTmpDX
                            '        End If
                            '    Else
                            '        'Ok, normally, we would go right...
                            '        Dim lTmpDX As Int32 = lDestX - lTotalWidth
                            '        If Math.Abs(lLocX - lTmpDX) < Math.Abs(lLocX - lDestX) Then
                            '            lDestX = lTmpDX
                            '        End If
                            '    End If
                            'Else
                            '    'Add all but the first and last dest...

                            '    'Ok, one last verification of the map wrap...
                            '    Dim yMapWrapType As Byte = 0
                            '    If ptList(0).X < 40 Then
                            '        yMapWrapType = 1
                            '    ElseIf ptList(0).X > 200 Then
                            '        yMapWrapType = 2
                            '    End If

                            '    For X = 1 To ptList.GetUpperBound(0) - 1

                            '        If ptList(X).X < 40 Then        'now in map wrap 1
                            '            If yMapWrapType = 2 Then    'was i in map wrap 2??
                            '                'yes, so we have a map wrap scenario. ensure that my point is WAY positive
                            '                ptList(X).X += 240
                            '            End If
                            '            yMapWrapType = 1
                            '        ElseIf ptList(X).X > 200 Then   'now in map wrap 2
                            '            If yMapWrapType = 1 Then    'was i in map wrap 1???
                            '                'yes, so we have a map wrap scenario. ensure that my point is way negative
                            '                ptList(X).X -= 240
                            '            End If
                            '            yMapWrapType = 2
                            '        Else
                            '            yMapWrapType = 0
                            '        End If

                            '        ptTemp.X = (ptList(X).X - lHalfWidth) * lCellSpacing
                            '        ptTemp.Y = (ptList(X).Y - lHalfWidth) * lCellSpacing
                            '        With uDest
                            '            .DestAngle = CShort(LineAngleDegrees(lLastX, lLastZ, ptTemp.X, ptTemp.Y) * 10)
                            '            .DestX = ptTemp.X
                            '            .DestZ = ptTemp.Y
                            '        End With
                            '        colDests.Add(uDest)
                            '        lLastX = ptTemp.X
                            '        lLastZ = ptTemp.Y
                            '    Next X

                            '    If lDestX > lTotalWidth \ 2 Then lDestX -= lTotalWidth
                            '    If lDestX < lTotalWidth \ -2 Then lDestX += lTotalWidth

                            '    Dim lQrtrWidth As Int32 = CInt(lTotalWidth * 0.25F)
                            '    If lDestX > (lQrtrWidth * 3) Then
                            '        If yMapWrapType = 1 Then
                            '            lDestX -= lTotalWidth
                            '        End If
                            '    ElseIf lDestX < lQrtrWidth Then
                            '        If yMapWrapType = 2 Then
                            '            lDestX += lTotalWidth
                            '        End If
                            '    End If
                            '    ''On the last dest, take map wrap into consideration
                            '    'If lDestX < lLastX Then
                            '    '    'Ok, normally, we would go left...
                            '    '    Dim lTmpDX As Int32 = lDestX + lTotalWidth
                            '    '    If Math.Abs(lTmpDX - lLastX) < Math.Abs(lLastX - lDestX) Then
                            '    '        lDestX = lTmpDX
                            '    '    End If
                            '    'Else
                            '    '    'Ok, normally, we would go right...
                            '    '    Dim lTmpDX As Int32 = lDestX - lTotalWidth
                            '    '    If Math.Abs(lLastX - lTmpDX) < Math.Abs(lLastX - lDestX) Then
                            '    '        lDestX = lTmpDX
                            '    '    End If
                            '    'End If
                            'End If

                            'Now, we add the final dest specifically
                            uDest.DestAngle = iDestAngle
                            uDest.DestX = lDestX
                            uDest.DestZ = lDestZ
                            uDest.iDestTypeID = iDestTypeID
                            uDest.lDestID = lDestID
                            uDest.yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
                            uDest.ySpecialOp = CByte(SpecialOp.eNoSpecialOp)
                            colDests.Add(uDest)

                        Else : Return Nothing
                        End If
                    End If


                Else : Return Nothing     'no planet to path on...
                End If
            Else : Return Nothing       'Suppose to path on terrain, but there is no terrain to path on...
            End If
        Else
            ''Figure out if we need to check for map wrapping or not
            'If iDestTypeID = ObjectType.ePlanet AndAlso (iLocTypeID = ObjectType.ePlanet AndAlso lLocID = lDestID) Then
            '    'Ok, this is for map wrapping East-West..
            '    Dim oPlanet As Planet = Nothing
            '    If lDestID >= 500000000 Then
            '        oPlanet = Planet.GetInstancePlanet
            '    Else
            '        For X = 0 To glPlanetVPUB
            '            If goPlanetVP(X).lID = lDestID Then
            '                oPlanet = goPlanetVP(X).oPlanetRef
            '                Exit For
            '            End If
            '        Next X
            '    End If
            '    If oPlanet Is Nothing = False Then

            '        Dim lTotalWidth As Int32 = TerrainClass.Width * oPlanet.GetCellSpacing()

            '        If lDestX < lLocX Then
            '            'Ok, normally, we would go left...
            '            Dim lTmpDX As Int32 = lDestX + lTotalWidth
            '            If Math.Abs(lTmpDX - lLocX) < Math.Abs(lLocX - lDestX) Then
            '                lDestX = lTmpDX
            '            End If
            '        Else
            '            'Ok, normally, we would go right...
            '            Dim lTmpDX As Int32 = lDestX - lTotalWidth
            '            If Math.Abs(lLocX - lTmpDX) < Math.Abs(lLocX - lDestX) Then
            '                lDestX = lTmpDX
            '            End If
            '        End If

            '        'If Math.Abs(lLocX - lDestX) > Math.Abs(lLocX - Math.Abs(lDestX - lTotalWidth)) Then
            '        '    lDestX -= lTotalWidth
            '        'End If
            '    End If
            '    oPlanet = Nothing
            'End If
            uDest.DestAngle = iDestAngle
            uDest.DestX = lDestX
            uDest.DestZ = lDestZ
            uDest.iDestTypeID = iDestTypeID
            uDest.lDestID = lDestID
            uDest.yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
            uDest.ySpecialOp = CByte(SpecialOp.eNoSpecialOp)
            colDests.Add(uDest)
        End If

        Return colDests

    End Function

    Private Sub PlanetToSystemMove(ByRef colDests As Collection, ByVal oPlanet As Planet, ByVal lLocX As Int32, ByVal lLocZ As Int32)
        'ok, we add our special dest...
        Dim uDest As DestLoc
        Dim fRadius As Single = oPlanet.PlanetRadius

        'ok, first dest
        With uDest
            .ySpecialOp = SpecialOp.eNoSpecialOp
            .yChangeEnvironment = ChangeEnvironmentType.ePlanetToSystem
            .DestX = lLocX + 1000
            .DestZ = lLocZ
            .DestAngle = CShort(LineAngleDegrees(lLocX, lLocZ, lLocX + 50, lLocZ) * 10)
            .iDestTypeID = oPlanet.ObjTypeID
            .lDestID = oPlanet.ObjectID
        End With
        colDests.Add(uDest)

		Dim fFinalDestX As Single
		Dim fFinalDestZ As Single

        'Will be sent to Primary as the initial location for the entity at the new environment
        With uDest
            .yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
            .ySpecialOp = SpecialOp.eNoSpecialOp
			.DestAngle = 0

			Dim fTmpX As Single = (oPlanet.PlanetRadius / 2.0F)
			Dim fTmpZ As Single = 0
			RotatePoint(0, 0, fTmpX, fTmpZ, Rnd() * 360.0F)

			fFinalDestX = fTmpX * 2
			fFinalDestZ = fTmpZ * 2

			.DestX = CInt(oPlanet.LocX + fTmpX)
			.DestZ = CInt(oPlanet.LocZ + fTmpZ)
			'.DestX = oPlanet.LocX
			'.DestZ = oPlanet.LocZ
            .iDestTypeID = oPlanet.ParentSystem.ObjTypeID
            .lDestID = oPlanet.ParentSystem.ObjectID
        End With
        colDests.Add(uDest)

        'then, a new dest at the 0 degrees... 
		uDest.DestX = CInt(oPlanet.LocX + fFinalDestX)
		uDest.DestZ = CInt(oPlanet.LocZ + fFinalDestZ)
        uDest.DestAngle = 0
        uDest.iDestTypeID = ObjectType.eSolarSystem
        uDest.lDestID = oPlanet.ParentSystem.ObjectID
        uDest.yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
        uDest.ySpecialOp = SpecialOp.eNoSpecialOp
        colDests.Add(uDest)
    End Sub

    Private Sub SystemToPlanetMove(ByRef colDests As Collection, ByVal oPlanet As Planet, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal lDestX As Int32, ByVal lDestZ As Int32)
        'ok, set up our approach...
        Dim fAngle As Single = LineAngleDegrees(lLocX, lLocZ, oPlanet.LocX, oPlanet.LocY)
        Dim lX As Int32 = oPlanet.LocX + oPlanet.PlanetRadius
        Dim lZ As Int32 = oPlanet.LocZ
        Dim uDest As DestLoc

        'Now, rotate our point around the planet to determine the location to move to
		'fAngle += 180.0F
        If fAngle > 360.0F Then fAngle -= 360.0F
        RotatePoint(oPlanet.LocX, oPlanet.LocZ, lX, lZ, fAngle)

        'Now, add that as a dest
        With uDest
            .DestX = lX
            .DestZ = lZ
            .DestAngle = CShort(fAngle * 10)
            .iDestTypeID = oPlanet.ParentSystem.ObjTypeID
            .lDestID = oPlanet.ParentSystem.ObjectID
            .yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
            .ySpecialOp = SpecialOp.eNoSpecialOp
        End With
        colDests.Add(uDest)

        'Now, set our loc to the center of hte planet
        With uDest
            .DestX = oPlanet.LocX
            .DestZ = oPlanet.LocZ
            .DestAngle = CShort(fAngle * 10)
            .yChangeEnvironment = ChangeEnvironmentType.eSystemToPlanet
            .ySpecialOp = SpecialOp.eNoSpecialOp
            .iDestTypeID = oPlanet.ObjTypeID
            .lDestID = oPlanet.ObjectID
        End With
        colDests.Add(uDest)

        'Now, next dest is the dest to initially put the entity on the planet at max Y
        With uDest
            .DestX = lDestX - 50
            .DestZ = lDestZ
            .DestAngle = CShort(LineAngleDegrees(lDestX, lDestZ, lDestX - 50, lDestZ) * 10)
            .yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
            .ySpecialOp = SpecialOp.eNoSpecialOp
            .iDestTypeID = oPlanet.ObjTypeID
            .lDestID = oPlanet.ObjectID
        End With
        colDests.Add(uDest)

        'our final loc will be set in the calling function...
    End Sub

    Private Sub SystemToSystemMove(ByRef colDests As Collection, ByVal oSystem1 As SolarSystem, ByVal oSystem2 As SolarSystem)
        Dim fLeaveAngle As Single = LineAngleDegrees(oSystem1.LocX, oSystem1.LocZ, oSystem2.LocX, oSystem2.LocZ)
        Dim fEntryAngle As Single = -fLeaveAngle

        Dim lX As Int32 = 5000000      '10,000,000 is max size
        Dim lZ As Int32 = 0

        Dim uDest As DestLoc

        'Create our exit location...
        'Now, rotate our point
        RotatePoint(0, 0, lX, lZ, fLeaveAngle)
        With uDest
            .DestX = lX
            .DestZ = lZ
            .DestAngle = CShort(fLeaveAngle * 10)
            .yChangeEnvironment = ChangeEnvironmentType.eSystemToSystem
            .ySpecialOp = SpecialOp.eNoSpecialOp
            .iDestTypeID = oSystem1.ObjTypeID
            .lDestID = oSystem1.ObjectID
        End With
        colDests.Add(uDest)

        'Now, this is the location to create the entity at the destination
        lX = 5000000
        lZ = 0
        RotatePoint(0, 0, lX, lZ, fEntryAngle)
        With uDest
            .DestAngle = CShort(fEntryAngle * 10)
            .DestX = lX
            .DestZ = lZ
            .yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
            .ySpecialOp = SpecialOp.eNoSpecialOp
            .iDestTypeID = oSystem2.ObjTypeID
            .lDestID = oSystem2.ObjectID
        End With
        colDests.Add(uDest)
    End Sub

    Public Sub LoadModelVP()
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"
        Dim oINI As InitFile = New InitFile(sFile & "PF_Model.dat")
        Dim X As Int32
        Dim lModelCnt As Int32 = CInt(Val(oINI.GetString("MODEL_LIST", "ModelCnt", "0")))
        Dim lVal As Int32

        'I know, a little goofy, but trust me...
        glModelUB = lModelCnt
        ReDim goModels(glModelUB + 1)

        For X = 1 To lModelCnt
            lVal = CInt(Val(oINI.GetString("Model_" & X, "Spacing", gl_FINAL_GRID_SQUARE_SIZE.ToString)))
            If lVal <> -1 Then
                goModels(X) = New ModelData
                With goModels(X)
                    .lModelID = X
                    .lRectSize = lVal
                    .PlotPaths = CBool(Val(oINI.GetString("Model_" & X, "PlotPaths", "0")) <> 0)
                    .bNaval = CBool(Val(oINI.GetString("Model_" & X, "IsNaval", "0")) <> 0)
                End With
            End If
        Next X


    End Sub

#Region " Trig Functions "
    'For Geometric Calculations
    Public Const gdPi As Single = 3.14159274F
    Public Const gdHalfPie As Single = gdPi / 2.0F
    Public Const gdPieAndAHalf As Single = gdPi * 1.5F
    Public Const gdTwoPie As Single = gdPi * 2.0F
    Public Const gdDegreePerRad As Single = 180.0F / gdPi
    Public Const gdRadPerDegree As Single = Math.PI / 180.0F

    Public Function RadianToDegree(ByVal fRads As Single) As Single
        Return fRads * gdDegreePerRad
    End Function

    Public Function DegreeToRadian(ByVal fDegree As Single) As Single
        Return fDegree * gdRadPerDegree
    End Function

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Int32, ByRef lEndY As Int32, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * (gdPi / 180.0F)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CInt((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CInt((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef fEndX As Single, ByRef fEndY As Single, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * (gdPi / 180.0F)
        dDX = fEndX - lAxisX
        dDY = fEndY - lAxisY
        fEndX = lAxisX + CSng((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        fEndY = lAxisY + -CSng((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Function LineAngleDegrees(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
        Dim dDeltaX As Single
        Dim dDeltaY As Single
        Dim dAngle As Single

        dDeltaX = lX2 - lX1
        dDeltaY = lY2 - lY1

        If dDeltaX = 0 Then     'vertical
            If dDeltaY < 0 Then
                dAngle = gdHalfPie
            Else
                dAngle = gdPieAndAHalf
            End If
        ElseIf dDeltaY = 0 Then     'horizontal
            If dDeltaX < 0 Then
                dAngle = gdPi
            Else
                dAngle = 0
            End If
        Else    'angled
            dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))
            'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
            If dDeltaX > -1 And dDeltaY > -1 Then       'VB Lower Right
                dAngle = gdTwoPie - dAngle
            ElseIf dDeltaX < 0 And dDeltaY > -1 Then    'VB Lower Left
                dAngle = gdPi + dAngle
            ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left
                dAngle = gdPi - dAngle
            End If
        End If

        Return 360.0F - (dAngle * gdDegreePerRad)

    End Function

    Public Function LineAngleDegrees_Pts(ByVal ptStart As Point, ByVal ptEnd As Point) As Int32
        Dim dDeltaX As Single
        Dim dDeltaY As Single
        Dim dAngle As Single

        dDeltaX = ptEnd.X - ptStart.X
        dDeltaY = ptEnd.Y - ptStart.Y

        If dDeltaX = 0 Then     'vertical
            If dDeltaY < 0 Then
                dAngle = gdHalfPie
            Else
                dAngle = gdPieAndAHalf
            End If
        ElseIf dDeltaY = 0 Then     'horizontal
            If dDeltaX < 0 Then
                dAngle = gdPi
            Else
                dAngle = 0
            End If
        Else    'angled
            dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))
            'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
            If dDeltaX > -1 And dDeltaY > -1 Then       'VB Lower Right
                dAngle = gdTwoPie - dAngle
            ElseIf dDeltaX < 0 And dDeltaY > -1 Then    'VB Lower Left
                dAngle = gdPi + dAngle
            ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left
                dAngle = gdPi - dAngle
            End If
        End If

        Return CInt(360.0F - (dAngle * gdDegreePerRad))
    End Function
#End Region
#Region "  String Management  "
    Public Function BytesToString(ByVal yBytes() As Byte) As String
        Dim lLen As Int32
        Dim X As Int32

        lLen = yBytes.Length
        For X = 0 To yBytes.Length - 1
            If yBytes(X) = 0 Then
                lLen = X
                Exit For
            End If
        Next X
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yBytes), 1, lLen)
    End Function

    Public Function StringToBytes(ByVal sVal As String) As Byte()
        Return System.Text.ASCIIEncoding.ASCII.GetBytes(sVal)
    End Function

    Public Function MakeDBStr(ByVal sVal As String) As String
        Return Replace$(sVal, "'", "''")
    End Function

    Public Function GetStringFromBytes(ByRef yData() As Byte, ByVal lPos As Int32, ByVal lDataLen As Int32) As String
        Dim yTemp(lDataLen - 1) As Byte
        Array.Copy(yData, lPos, yTemp, 0, lDataLen)
        Dim lLen As Int32 = yTemp.Length
        For Y As Int32 = 0 To yTemp.Length - 1
            If yTemp(Y) = 0 Then
                lLen = Y
                Exit For
            End If
        Next Y
        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)
    End Function
#End Region

#Region "  Formation Management  "
	Public goFormation() As Formation
	Public glFormationUB As Int32 = -1
	Public glFormationIdx() As Int32

	Public goFormationDef() As FormationDefinition
	Public glFormationDefIdx() As Int32
	Public glFormationDefUB As Int32 = -1

	Public Sub AddFormationInstance(ByRef oFormation As Formation)
		Dim lIdx As Int32 = -1

		If glFormationIdx Is Nothing Then ReDim glFormationIdx(-1)
		For X As Int32 = 0 To glFormationUB
			If glFormationIdx(X) = -1 Then
				lIdx = X
				Exit For
			End If
		Next X

		SyncLock glFormationIdx
			Try
				If lIdx = -1 Then
					lIdx = glFormationUB + 1
					If lIdx > glFormationIdx.GetUpperBound(0) Then
						ReDim Preserve glFormationIdx(glFormationUB + 100)
						ReDim Preserve goFormation(glFormationUB + 100)
					End If
					glFormationIdx(lIdx) = -2
					glFormationUB += 1
				End If

				oFormation.ServerIndex = lIdx
				goFormation(lIdx) = oFormation
				glFormationIdx(lIdx) = oFormation.oFormationDef.lFormationID
			Catch
				'do nothing, they will need to send it moving again
			End Try
		End SyncLock
	End Sub
#End Region
End Module
