Option Strict On

'The Mover class contains the data necessary to path and move an object... the Pathfinding server does not actually
'  do the moving of the object... it only does the pathfinding which could be a fairly large task to accomplish on
'  the domain server with other tasks occurring...
Public Class Mover
    Inherits Base_GUID

    Public lGlobalIdx As Int32

    Public lCurLocX As Int32
    Public lCurLocY As Int32    'loc y not as important as x and z
    Public lCurLocZ As Int32
    Public lCurAngle As Int16

    Public lCurDestX As Int32
    Public lCurDestZ As Int32

    Public yLastTransitionType As Byte

    Public colDests As Collection = New Collection()

    Public rcArea As Rectangle
    'Public iModelID As Int16
    Private miModelID As Int16
    Public Property iModelID() As Int16
        Get
            Return miModelID
        End Get
        Set(ByVal value As Int16)
            'MSC: The Pathfinding Server ONLY cares about the first byte of the modelid - the model itself
            miModelID = (value And 255S)
        End Set
    End Property
    Public bRectSet As Boolean = False

    'The last known environment ID and environment typeid
    Public lEnvirID As Int32
    Public iEnvirTypeID As Int16

    Public bMoving As Boolean = False

    Public oFormation As Formation = Nothing

    Public bAlertAtFinalDest As Boolean = False

    Public bInARoute As Boolean = False

    Public yInAForcedMoveSpeed As Byte = 0


    Public Function ProcessStopMessage(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal lLocAngle As Int16) As Byte()
        Dim uDest As DestLoc
        Dim yTemp() As Byte = Nothing

        lCurLocX = lLocX
        lCurLocZ = lLocZ
        lCurAngle = lLocAngle

        yLastTransitionType = 0

        'now, determine if this is our final dest
        If colDests Is Nothing = False AndAlso colDests.Count > 0 Then
            'get the next dest in the list
            uDest = CType(colDests(1), DestLoc)
            'remove that dest from the list
            colDests.Remove(1)

            Me.lEnvirID = uDest.lDestID
            Me.iEnvirTypeID = uDest.iDestTypeID

            bMoving = True
            If uDest.yChangeEnvironment = ChangeEnvironmentType.eDocking OrElse uDest.yChangeEnvironment = ChangeEnvironmentType.eSystemToSystem Then
                ReDim yTemp(24)

                'gfrmDisplayForm.LogSpecificEvent("Docking Unit " & Me.ObjectID & " with " & lEnvirID & ", " & iEnvirTypeID)

                System.BitConverter.GetBytes(GlobalMessageCode.eEntityChangeEnvironment).CopyTo(yTemp, 0)
                GetGUIDAsString.CopyTo(yTemp, 2)

                With uDest
                    lCurDestX = .DestX
                    lCurDestZ = .DestZ
                    System.BitConverter.GetBytes(.DestX).CopyTo(yTemp, 8)
                    System.BitConverter.GetBytes(.DestZ).CopyTo(yTemp, 12)
                    System.BitConverter.GetBytes(.DestAngle).CopyTo(yTemp, 16)
                    System.BitConverter.GetBytes(.lDestID).CopyTo(yTemp, 18)
                    System.BitConverter.GetBytes(.iDestTypeID).CopyTo(yTemp, 22)
                    yTemp(24) = .yChangeEnvironment
                End With
                goMsgSys.SendToPrimary(yTemp)

                'Now, erase my array and set moving to false
                ReDim yTemp(-1)
                bMoving = False
            ElseIf uDest.yChangeEnvironment <> ChangeEnvironmentType.eNoChangeEnvironment Then
                'Ok, gotta change the environment... we send the move message as normal but with a special instruction
                ReDim yTemp(18)
                System.BitConverter.GetBytes(GlobalMessageCode.eEntityChangeEnvironment).CopyTo(yTemp, 0)
                GetGUIDAsString.CopyTo(yTemp, 2)
                System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 8)
                System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 12)
                System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yTemp, 16)
                yTemp(18) = uDest.yChangeEnvironment

                lCurDestX = uDest.DestX
                lCurDestZ = uDest.DestZ

                yLastTransitionType = uDest.yChangeEnvironment
            ElseIf uDest.ySpecialOp = SpecialOp.eNoSpecialOp Then
                'Ok, normal
                If yInAForcedMoveSpeed <> 0 Then ReDim yTemp(18) Else ReDim yTemp(17)
                'ReDim yTemp(17)		'0 to 17 = 18 bytes
                If colDests.Count > 0 Then
                    System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yTemp, 0)
                Else : System.BitConverter.GetBytes(GlobalMessageCode.eFinalMoveCommand).CopyTo(yTemp, 0)
                End If
                GetGUIDAsString.CopyTo(yTemp, 2)
                System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 8)
                System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 12)
                System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yTemp, 16)

                If yInAForcedMoveSpeed <> 0 Then yTemp(18) = yInAForcedMoveSpeed

                lCurDestX = uDest.DestX
                lCurDestZ = uDest.DestZ
            ElseIf uDest.ySpecialOp = SpecialOp.eBeginMiningOp Then
                'Ok, special op = mining op

                If lLocX = uDest.DestX AndAlso lLocZ = uDest.DestZ Then

                    ReDim yTemp(13)
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetMiningLoc).CopyTo(yTemp, 0)
                    GetGUIDAsString.CopyTo(yTemp, 2)
                    System.BitConverter.GetBytes(uDest.lSpecialOpID).CopyTo(yTemp, 8)
                    System.BitConverter.GetBytes(uDest.iSpecialOpTypeID).CopyTo(yTemp, 12)

                    'This is special, we send to the primary and erase our array to be sent to the Domain
                    goMsgSys.SendToPrimary(yTemp)

                    'Now, erase my array and set moving to false
                    ReDim yTemp(-1)
                    bMoving = False
                Else
                    'Ok, normal
                    ReDim yTemp(17)     '0 to 17 = 18 bytes
                    System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yTemp, 0)
                    GetGUIDAsString.CopyTo(yTemp, 2)
                    System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 8)
                    System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 12)
                    System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yTemp, 16)
                    colDests.Add(uDest)
                End If

            ElseIf uDest.ySpecialOp = SpecialOp.eBeginConstruction Then
                'Ok, special op = begin construction
                ReDim yTemp(23)
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yTemp, 0)
                GetGUIDAsString.CopyTo(yTemp, 2)
                System.BitConverter.GetBytes(uDest.lSpecialOpID).CopyTo(yTemp, 8)
                System.BitConverter.GetBytes(uDest.iSpecialOpTypeID).CopyTo(yTemp, 12)
                System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 14)
                System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 18)
                System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yTemp, 22)

                'Send to primary, then erase it
                goMsgSys.SendToPrimary(yTemp)
                'now erase it so the domain doesn't get it
                ReDim yTemp(-1)
                bMoving = False
            ElseIf uDest.ySpecialOp = SpecialOp.eJumpTarget Then
                'Ok, special op = Jump Target
                ReDim yTemp(29)
                System.BitConverter.GetBytes(GlobalMessageCode.eJumpTarget).CopyTo(yTemp, 0)
                GetGUIDAsString.CopyTo(yTemp, 2)
                System.BitConverter.GetBytes(uDest.lSpecialOpID).CopyTo(yTemp, 8)
                System.BitConverter.GetBytes(uDest.iSpecialOpTypeID).CopyTo(yTemp, 12)
                System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 14)
                System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 18)
                System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yTemp, 22)
                System.BitConverter.GetBytes(uDest.lDestID).CopyTo(yTemp, 24)
                System.BitConverter.GetBytes(uDest.iDestTypeID).CopyTo(yTemp, 28)

                'Send to primary, then erase it
                goMsgSys.SendToPrimary(yTemp)
                'now erase it so the domain doesn't get it
                ReDim yTemp(-1)
                bMoving = False
            ElseIf uDest.ySpecialOp = SpecialOp.eBeginDismantle OrElse uDest.ySpecialOp = SpecialOp.eBeginRepair Then
                ReDim yTemp(21)
                If uDest.ySpecialOp = SpecialOp.eBeginDismantle Then
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetDismantleTarget).CopyTo(yTemp, 0)
                Else : System.BitConverter.GetBytes(GlobalMessageCode.eSetRepairTarget).CopyTo(yTemp, 0)
                End If
                GetGUIDAsString.CopyTo(yTemp, 2)
                System.BitConverter.GetBytes(uDest.lSpecialOpID).CopyTo(yTemp, 8)
                System.BitConverter.GetBytes(uDest.iSpecialOpTypeID).CopyTo(yTemp, 12)
                System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 14)
                System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 18)

                'Send to primary, then erase it
                goMsgSys.SendToPrimary(yTemp)
                'now erase it so the domain doesn't get it
                ReDim yTemp(-1)
                bMoving = False
            End If
        Else
            ReDim yTemp(7)
            System.BitConverter.GetBytes(GlobalMessageCode.eFinalizeStopEvent).CopyTo(yTemp, 0)
            Me.GetGUIDAsString.CopyTo(yTemp, 2)
            bMoving = False
        End If

        If bMoving = False AndAlso bAlertAtFinalDest = True Then
            bAlertAtFinalDest = False
            Dim yNewMsg(21) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eAlertDestinationReached).CopyTo(yNewMsg, lPos) : lPos += 2
            Me.GetGUIDAsString.CopyTo(yNewMsg, lPos) : lPos += 6
            System.BitConverter.GetBytes(lCurLocX).CopyTo(yNewMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lCurLocZ).CopyTo(yNewMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lEnvirID).CopyTo(yNewMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yNewMsg, lPos) : lPos += 2
            goMsgSys.SendToPrimary(yNewMsg)
        End If

        Return yTemp
    End Function

    Public Shared Function FinalizeStopEvent(ByRef oMover As Mover) As Boolean

        Dim lExtent As Int32

        If oMover.bInARoute = True Then
            'gfrmDisplayForm.AddEventLine("FinalizeStopEvent on a mover that is in a route. MoverGUID: " & oMover.ObjectID & ", " & oMover.ObjTypeID)
            Return True
        End If

        If oMover.iEnvirTypeID = ObjectType.eSolarSystem Then
            lExtent = 10000000
        ElseIf oMover.iEnvirTypeID = ObjectType.ePlanet Then
            If oMover.lEnvirID >= 500000000 Then
                Dim oP As Planet = Planet.GetInstancePlanet
                lExtent = (oP.GetCellSpacing * TerrainClass.Width) \ 2
            Else
                For X As Int32 = 0 To glPlanetVPUB
                    If goPlanetVP(X).lID = oMover.lEnvirID Then
                        'Ok, found it...
                        lExtent = (goPlanetVP(X).oPlanetRef.GetCellSpacing * TerrainClass.Width) \ 2
                        Exit For
                    End If
                Next X
            End If
        End If



        With oMover
            If .bRectSet = False Then
                .rcArea = Mover.GetRectWidthHeight(.iModelID)
                .bRectSet = True
            End If

            .rcArea.X = .lCurLocX - (.rcArea.Width \ 2)
            .rcArea.Y = .lCurLocZ - (.rcArea.Height \ 2)
        End With
        'Now, let's figure out where to place this guy
        Dim muUnitRect As Rectangle = oMover.rcArea
        Dim muFinalDestRect As Rectangle = muUnitRect

        Dim lCurrentDistX As Int32 = 0
        Dim fCurrentAngle As Single = 0.0F
        Dim fAngleAdjust As Single = 0.0F
        Dim lXAdjust As Int32 = muUnitRect.Width '\ 2
        Dim bResult As Boolean = True
        Dim bDone As Boolean = False
        Dim lCnt As Int32 = 0


        Dim oPlanet As Planet = Nothing
        If oMover.iEnvirTypeID = ObjectType.ePlanet Then
            If oMover.lEnvirID > 500000000 Then
                oPlanet = Planet.GetInstancePlanet
            Else
                For X As Int32 = 0 To glPlanetVPUB
                    If goPlanetVP(X).lID = oMover.lEnvirID Then
                        oPlanet = goPlanetVP(X).oPlanetRef
                        Exit For
                    End If
                Next X
            End If
        End If

        Dim yLandUnit As Byte = GetPathLocs(oMover.iModelID)
        If yLandUnit <> 0 Then
            If oMover.iEnvirTypeID = ObjectType.ePlanet AndAlso oPlanet Is Nothing = False Then
                'ok, on a planet, with a planet-based mover... is my current loc water?
                If yLandUnit = 1 AndAlso oPlanet.GetHeightAtPoint(oMover.lCurLocX, oMover.lCurLocZ) <= oPlanet.WaterRealHeight Then
                    'Yup.. ok move me out
                    Dim lCellSpacing As Int32 = oPlanet.GetCellSpacing
                    Dim lSX As Int32 = (oMover.lCurLocX + oPlanet.lHalfExtent) \ lCellSpacing
                    Dim lSZ As Int32 = (oMover.lCurLocZ + oPlanet.lHalfExtent) \ lCellSpacing

                    Dim ptResult As Point = oPlanet.oPather.GetClosestTileTypeToLocation(lSX, lSZ, 0)
                    If ptResult.IsEmpty = False Then
                        muFinalDestRect.X = (ptResult.X - 120) * lCellSpacing : muFinalDestRect.Y = (ptResult.Y - 120) * lCellSpacing
                        bDone = True
                        bResult = False
                    End If
                ElseIf yLandUnit = 2 AndAlso oPlanet.GetHeightAtPoint(oMover.lCurLocX, oMover.lCurLocZ) > oPlanet.WaterRealHeight Then
                    'Yup.. ok move me out
                    Dim lCellSpacing As Int32 = oPlanet.GetCellSpacing
                    Dim lSX As Int32 = (oMover.lCurLocX + oPlanet.lHalfExtent) \ lCellSpacing
                    Dim lSZ As Int32 = (oMover.lCurLocZ + oPlanet.lHalfExtent) \ lCellSpacing

                    Dim ptResult As Point = oPlanet.oPather.GetClosestTileTypeToLocation(lSX, lSZ, 1)
                    If ptResult.IsEmpty = False Then
                        muFinalDestRect.X = (ptResult.X - 120) * lCellSpacing : muFinalDestRect.Y = (ptResult.Y - 120) * lCellSpacing
                        bDone = True
                        bResult = False
                    End If
                End If
            End If
        End If

        While bDone = False AndAlso lCnt < 2000 AndAlso oMover.iEnvirTypeID = ObjectType.ePlanet AndAlso yLandUnit <> 0
            bDone = True
            lCnt += 1
            Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If glMoverIdx(X) > 0 Then 'AndAlso goMovers(X) Is Nothing = False AndAlso goMovers(X).lEnvirID = oMover.lEnvirID AndAlso goMovers(X).iEnvirTypeID = oMover.iEnvirTypeID Then
                    Dim oTmpMover As Mover = goMovers(X)
                    If oTmpMover Is Nothing Then Continue For
                    If oTmpMover.lEnvirID <> oMover.lEnvirID OrElse oTmpMover.iEnvirTypeID <> oMover.iEnvirTypeID Then Continue For

                    'now, check for whether we are looking at myself
                    If glMoverIdx(X) <> oMover.ObjectID OrElse oTmpMover.ObjTypeID <> oMover.ObjTypeID Then
                        'ok, not me, so let's compare...
                        If oTmpMover.bRectSet = False Then
                            oTmpMover.rcArea = Mover.GetRectWidthHeight(oTmpMover.iModelID)
                            oTmpMover.rcArea.X = oTmpMover.lCurLocX - (oTmpMover.rcArea.Width \ 2)
                            oTmpMover.rcArea.Y = oTmpMover.lCurLocY - (oTmpMover.rcArea.Height \ 2)
                            oTmpMover.bRectSet = True
                        End If

                        Dim rcTemp As Rectangle = oTmpMover.rcArea
                        If oTmpMover.bMoving = True Then
                            rcTemp.X = oTmpMover.lCurDestX
                            rcTemp.Y = oTmpMover.lCurDestZ
                        End If

                        If rcTemp.IntersectsWith(muFinalDestRect) = True OrElse muFinalDestRect.X < -lExtent OrElse muFinalDestRect.X > lExtent OrElse muFinalDestRect.Y < -lExtent OrElse muFinalDestRect.Y > lExtent Then
                            bDone = False
                            bResult = False
                            If lCurrentDistX = 0 Then
                                lCurrentDistX = lXAdjust
                                fAngleAdjust = LineAngleDegrees(0, 0, lCurrentDistX, (Math.Max(muUnitRect.Width, muUnitRect.Height)))
                            Else
                                Dim lTmpX As Int32 = lCurrentDistX
                                Dim lTmpZ As Int32 = 0
                                fCurrentAngle += fAngleAdjust
                                If fCurrentAngle >= 360 Then
                                    lCurrentDistX += lXAdjust
                                    lTmpX = lCurrentDistX : lTmpZ = 0
                                    fAngleAdjust = LineAngleDegrees(0, 0, lCurrentDistX, (Math.Max(muUnitRect.Width, muUnitRect.Height)))
                                    fCurrentAngle = 0
                                Else
                                    RotatePoint(0, 0, lTmpX, lTmpZ, fCurrentAngle)
                                End If

                                Dim lFinalLocX As Int32 = muUnitRect.X + lTmpX
                                Dim lFinalLocZ As Int32 = muUnitRect.Y + lTmpZ

                                If oMover.iEnvirTypeID = ObjectType.ePlanet AndAlso oPlanet Is Nothing = False Then
                                    If yLandUnit = 1 Then
                                        'normal land unit
                                        If oPlanet.GetHeightAtPoint(lFinalLocX, lFinalLocZ) > oPlanet.WaterRealHeight Then
                                            muFinalDestRect.X = lFinalLocX : muFinalDestRect.Y = lFinalLocZ
                                        End If
                                    ElseIf yLandUnit = 2 Then
                                        'naval unit
                                        If oPlanet.GetHeightAtPoint(lFinalLocX, lFinalLocZ) < oPlanet.WaterRealHeight Then
                                            muFinalDestRect.X = lFinalLocX : muFinalDestRect.Y = lFinalLocZ
                                        End If
                                    Else
                                        'air unit
                                        muFinalDestRect.X = lFinalLocX : muFinalDestRect.Y = lFinalLocZ
                                    End If
                                Else
                                    muFinalDestRect.X = lFinalLocX : muFinalDestRect.Y = lFinalLocZ
                                End If

                            End If
                            Exit For
                        End If
                    End If
                End If

            Next X

        End While

        'Now, we have a new dest...
        If bResult = False Then
            Dim uDest As DestLoc

            uDest.DestX = muFinalDestRect.X + (muFinalDestRect.Width \ 2)
            uDest.DestZ = muFinalDestRect.Y + (muFinalDestRect.Height \ 2)

            'MSC - added to remove never ending movement...
            'TODO: Make this.... cleaner
            If Math.Abs(uDest.DestX) > lExtent OrElse Math.Abs(uDest.DestZ) > lExtent Then Return True

            uDest.DestAngle = oMover.lCurAngle
            uDest.iDestTypeID = oMover.iEnvirTypeID
            uDest.lDestID = oMover.lEnvirID
            uDest.yChangeEnvironment = 0
            uDest.ySpecialOp = 0

            oMover.colDests.Add(uDest)
        ElseIf oMover.oFormation Is Nothing = False Then
            oMover.oFormation.ItemInPlace(oMover.ObjectID, oMover.ObjTypeID)
            oMover.oFormation.CheckForArrival()
        End If

        Return bResult

    End Function

    'Public Shared Function FinalizeStopEvent(ByRef oMover As Mover) As Boolean
    '    Dim X As Int32

    '    Dim lDiffX As Int32
    '    Dim lDiffZ As Int32

    '    Dim lExtent As Int32

    '    If oMover.iEnvirTypeID = ObjectType.eSolarSystem Then
    '        lExtent = 10000000
    '    ElseIf oMover.iEnvirTypeID = ObjectType.ePlanet Then
    '        For X = 0 To glPlanetVPUB
    '            If goPlanetVP(X).lID = oMover.lEnvirID Then
    '                'Ok, found it...
    '                lExtent = (goPlanetVP(X).oPlanetRef.GetCellSpacing * TerrainClass.Width) \ 2
    '                Exit For
    '            End If
    '        Next X
    '    End If

    '    With oMover
    '        If .bRectSet = False Then
    '            .rcArea = Mover.GetRectWidthHeight(.iModelID)
    '            .bRectSet = True
    '        End If

    '        .rcArea.X = .lCurLocX - (.rcArea.Width \ 2)
    '        .rcArea.Y = .lCurLocZ - (.rcArea.Height \ 2)

    '        For X = 0 To glMoverUB
    '            If glMoverIdx(X) <> -1 AndAlso goMovers(X).lEnvirID = .lEnvirID AndAlso goMovers(X).iEnvirTypeID = .iEnvirTypeID Then
    '                'now, check for whether we are looking at myself
    '                If glMoverIdx(X) <> .ObjectID OrElse goMovers(X).ObjTypeID <> .ObjTypeID Then
    '                    'ok, not me, so let's compare...
    '                    If goMovers(X).bRectSet = False Then
    '                        goMovers(X).rcArea = Mover.GetRectWidthHeight(goMovers(X).iModelID)
    '                        goMovers(X).rcArea.X = .lCurLocX - (.rcArea.Width \ 2)
    '                        goMovers(X).rcArea.Y = .lCurLocY - (.rcArea.Height \ 2)
    '                        goMovers(X).bRectSet = True
    '                    End If

    '                    If goMovers(X).rcArea.IntersectsWith(.rcArea) = True Then
    '                        'Ok, it does... so, let's figure it out... where we gonna go?
    '                        lDiffX = .rcArea.X - goMovers(X).rcArea.X
    '                        lDiffZ = .rcArea.Y - goMovers(X).rcArea.Y

    '                        If lDiffX > goMovers(X).rcArea.Width / 2 Then
    '                            'Move right
    '                            lDiffX = goMovers(X).rcArea.Width - lDiffX
    '                        Else
    '                            'Move left
    '                            lDiffX = -(lDiffX + .rcArea.Width)
    '                        End If

    '                        If lDiffZ > goMovers(X).rcArea.Height / 2 Then
    '                            'Move down
    '                            lDiffZ = goMovers(X).rcArea.Height - lDiffZ
    '                        Else
    '                            'Move up
    '                            lDiffZ = -(lDiffZ + .rcArea.Height)
    '                        End If

    '                        'Now, we have a new dest...
    '                        Dim uDest As DestLoc

    '                        uDest.DestX = .lCurLocX + lDiffX
    '                        uDest.DestZ = .lCurLocZ + lDiffZ

    '                        'MSC - added to remove never ending movement...
    '                        'TODO: Make this.... cleaner
    '                        If Math.Abs(uDest.DestX) > lExtent OrElse Math.Abs(uDest.DestZ) > lExtent Then Return True

    '                        uDest.DestAngle = .lCurAngle
    '                        uDest.iDestTypeID = .iEnvirTypeID
    '                        uDest.lDestID = .lEnvirID
    '                        uDest.yChangeEnvironment = 0
    '                        uDest.ySpecialOp = 0

    '                        .colDests.Add(uDest)

    '                        Return False
    '                    End If
    '                End If
    '            End If
    '        Next X
    '        Return True
    '    End With
    'End Function

    Public Shared Function GetRectWidthHeight(ByVal iModelID As Int16) As Rectangle
        Dim rcRet As Rectangle
        Dim lSize As Int32 = gl_FINAL_GRID_SQUARE_SIZE
        'MSC: PF only cares about the model portion of the modelid
        iModelID = (iModelID And 255S)

        If iModelID < goModels.Length AndAlso iModelID > 0 Then
            lSize = goModels(iModelID).lRectSize
        Else
            For X As Int32 = 0 To glModelUB
                If goModels(X) Is Nothing = False AndAlso goModels(X).lModelID = iModelID Then
                    lSize = goModels(X).lRectSize
                    Exit For
                End If
            Next X
        End If

        rcRet.Width = lSize
        rcRet.Height = lSize

        Return rcRet
    End Function

    Public Shared Function GetPathLocs(ByVal iModelID As Int16) As Byte
        'MSC: PF only cares about the model portion of the modelid
        iModelID = (iModelID And 255S)

        If iModelID < goModels.Length AndAlso iModelID > -1 Then
            If goModels(iModelID).bNaval = True Then Return 2
            If goModels(iModelID).PlotPaths = True Then Return 1
            Return 0
            'Return goModels(iModelID).PlotPaths
        Else
            For X As Int32 = 0 To glModelUB
                If goModels(X) Is Nothing = False AndAlso goModels(X).lModelID = iModelID Then
                    If goModels(X).bNaval = True Then Return 2
                    If goModels(X).PlotPaths = True Then Return 1
                    Return 0
                    'Return goModels(X).PlotPaths
                End If
            Next X
        End If
        Return 1 'True
    End Function

    Public Function ProcessDecelerateImminentMsg(ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestA As Int16) As Byte()
        'Ok, decelerate imminent msgs are ONLY meant to handle moving to the next waypoint, everything else should wait for the stop msg
        'furthermore, if it is the LAST dest in the list, then we do nothing... just wait for final stop msg

        'If oFormation Is Nothing = False Then oFormation.ItemInPlace(Me.ObjectID, Me.ObjTypeID)

        If colDests Is Nothing = False AndAlso colDests.Count > 0 Then
            Dim uDest As DestLoc = CType(colDests(1), DestLoc)

            'Now... if this next loc is a normal move loc...
            If yLastTransitionType = 0 AndAlso uDest.ySpecialOp = SpecialOp.eNoSpecialOp Then

                If uDest.yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment Then
                    'remove that dest from the list
                    colDests.Remove(1)

                    'Ok, normal dest...
                    Dim yResp(17) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yResp, 0)
                    GetGUIDAsString.CopyTo(yResp, 2)
                    System.BitConverter.GetBytes(uDest.DestX).CopyTo(yResp, 8)
                    System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yResp, 12)
                    System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yResp, 16)

                    Return yResp
                Else 'If uDest.yChangeEnvironment <> ChangeEnvironmentType.eDocking AndAlso uDest.yChangeEnvironment <> ChangeEnvironmentType.eUndocking Then
                    'remove that dest from the list
                    Return Me.ProcessStopMessage(lDestX, lDestZ, iDestA)
                End If

            End If
        ElseIf oFormation Is Nothing = False Then
            oFormation.ItemInPlace(ObjectID, ObjTypeID)
            oFormation.CheckForArrival()
        End If

        Return Nothing
    End Function
    '   Public Function ProcessDecelerateImminentMsg(ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestA As Int16) As Byte()
    '       'Ok, decelerate imminent msgs are ONLY meant to handle moving to the next waypoint, everything else should wait for the stop msg
    '	'furthermore, if it is the LAST dest in the list, then we do nothing... just wait for final stop msg

    '	'If oFormation Is Nothing = False Then oFormation.ItemInPlace(Me.ObjectID, Me.ObjTypeID)

    '       If colDests Is Nothing = False AndAlso colDests.Count > 0 Then
    '           Dim uDest As DestLoc = CType(colDests(1), DestLoc)

    '           'Now... if this next loc is a normal move loc...
    '		If yLastTransitionType = 0 AndAlso uDest.yChangeEnvironment <> ChangeEnvironmentType.eDocking AndAlso uDest.yChangeEnvironment <> ChangeEnvironmentType.eUndocking AndAlso uDest.ySpecialOp = SpecialOp.eNoSpecialOp Then
    '			'remove that dest from the list
    '			colDests.Remove(1)

    '			'Ok, normal dest...
    '			Dim yResp(17) As Byte
    '			System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yResp, 0)
    '			GetGUIDAsString.CopyTo(yResp, 2)
    '			System.BitConverter.GetBytes(uDest.DestX).CopyTo(yResp, 8)
    '			System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yResp, 12)
    '			System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yResp, 16)

    '			Return yResp
    '		End If
    '       End If

    '       Return Nothing
    'End Function

    Public Shared Function GetNearestAntiStackingPoint(ByVal lMoverID As Int32, ByVal iMoverTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32) As Point

        Dim lExtent As Int32

        Dim oMover As Mover = Nothing
        Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glMoverIdx(X) = lMoverID AndAlso goMovers(X) Is Nothing = False AndAlso goMovers(X).ObjTypeID = iMoverTypeID Then
                oMover = goMovers(X)
                Exit For
            End If
        Next X
        If oMover Is Nothing Then Return New Point(lLocX, lLocZ)

        If oMover.iEnvirTypeID = ObjectType.eSolarSystem Then
            lExtent = 10000000
        ElseIf oMover.iEnvirTypeID = ObjectType.ePlanet Then
            If oMover.lEnvirID >= 500000000 Then
                Dim oP As Planet = Planet.GetInstancePlanet
                lExtent = (oP.GetCellSpacing * TerrainClass.Width) \ 2
            Else
                For X As Int32 = 0 To glPlanetVPUB
                    If goPlanetVP(X).lID = oMover.lEnvirID Then
                        'Ok, found it...
                        lExtent = (goPlanetVP(X).oPlanetRef.GetCellSpacing * TerrainClass.Width) \ 2
                        Exit For
                    End If
                Next X
            End If
        End If

        With oMover
            If .bRectSet = False Then
                .rcArea = Mover.GetRectWidthHeight(.iModelID)
                .bRectSet = True
            End If

            .rcArea.X = lLocX - (.rcArea.Width \ 2)
            .rcArea.Y = lLocZ - (.rcArea.Height \ 2)
        End With

        'Now, let's figure out where to place this guy
        Dim muUnitRect As Rectangle = oMover.rcArea
        Dim muFinalDestRect As Rectangle = muUnitRect

        Dim lCurrentDistX As Int32 = 0
        Dim fCurrentAngle As Single = 0.0F
        Dim fAngleAdjust As Single = 0.0F
        Dim lXAdjust As Int32 = muUnitRect.Width '\ 2
        Dim bResult As Boolean = True
        Dim bDone As Boolean = False
        Dim lCnt As Int32 = 0

        While bDone = False AndAlso lCnt < 2000
            bDone = True
            lCnt += 1
            For X As Int32 = 0 To lCurUB
                If glMoverIdx(X) > 0 Then 'AndAlso goMovers(X) Is Nothing = False AndAlso goMovers(X).lEnvirID = oMover.lEnvirID AndAlso goMovers(X).iEnvirTypeID = oMover.iEnvirTypeID Then
                    Dim oTmpMover As Mover = goMovers(X)
                    If oTmpMover Is Nothing Then Continue For
                    If oTmpMover.lEnvirID <> oMover.lEnvirID OrElse oTmpMover.iEnvirTypeID <> oMover.iEnvirTypeID Then Continue For

                    'now, check for whether we are looking at myself
                    If glMoverIdx(X) <> oMover.ObjectID OrElse oTmpMover.ObjTypeID <> oMover.ObjTypeID Then
                        'ok, not me, so let's compare...
                        If oTmpMover.bRectSet = False Then
                            oTmpMover.rcArea = Mover.GetRectWidthHeight(oTmpMover.iModelID)
                            oTmpMover.rcArea.X = oTmpMover.lCurLocX - (oTmpMover.rcArea.Width \ 2)
                            oTmpMover.rcArea.Y = oTmpMover.lCurLocY - (oTmpMover.rcArea.Height \ 2)
                            oTmpMover.bRectSet = True
                        End If

                        Dim rcTemp As Rectangle = oTmpMover.rcArea
                        If oTmpMover.bMoving = True Then
                            rcTemp.X = oTmpMover.lCurDestX
                            rcTemp.Y = oTmpMover.lCurDestZ
                        End If

                        If rcTemp.IntersectsWith(muFinalDestRect) = True OrElse muFinalDestRect.X < -lExtent OrElse muFinalDestRect.X > lExtent OrElse muFinalDestRect.Y < -lExtent OrElse muFinalDestRect.Y > lExtent Then
                            bDone = False
                            bResult = False
                            If lCurrentDistX = 0 Then
                                lCurrentDistX = lXAdjust
                                fAngleAdjust = LineAngleDegrees(0, 0, lCurrentDistX, (Math.Max(muUnitRect.Width, muUnitRect.Height)))
                            Else
                                Dim lTmpX As Int32 = lCurrentDistX
                                Dim lTmpZ As Int32 = 0
                                fCurrentAngle += fAngleAdjust
                                If fCurrentAngle >= 360 Then
                                    lCurrentDistX += lXAdjust
                                    lTmpX = lCurrentDistX : lTmpZ = 0
                                    fAngleAdjust = LineAngleDegrees(0, 0, lCurrentDistX, (Math.Max(muUnitRect.Width, muUnitRect.Height)))
                                    fCurrentAngle = 0
                                Else
                                    RotatePoint(0, 0, lTmpX, lTmpZ, fCurrentAngle)
                                End If

                                muFinalDestRect.X = muUnitRect.X + lTmpX : muFinalDestRect.Y = muUnitRect.Y + lTmpZ
                            End If
                            Exit For
                        End If
                    End If
                End If

            Next X

        End While

        'Now, we have a new dest...
        If bResult = False Then
            Return New Point(muFinalDestRect.X + (muFinalDestRect.Width \ 2), muFinalDestRect.Y + (muFinalDestRect.Height \ 2))
        Else : Return New Point(lLocX, lLocZ)
        End If

    End Function

End Class

Public Structure DestLoc
    Public DestX As Int32
    Public DestZ As Int32
    Public DestAngle As Int16

    'Environment details
    Public lDestID As Int32
    Public iDestTypeID As Int16

    Public yChangeEnvironment As Byte

    Public ySpecialOp As Byte       'should always default to 0
    Public lSpecialOpID As Int32    'an ObjectID related to the Special Op
    Public iSpecialOpTypeID As Int16 'a typeid related to the special op
End Structure

Public Structure LocationalGrouping
    Public VertX As Int32
    Public VertZ As Int32

    Public colDestList As Collection
End Structure