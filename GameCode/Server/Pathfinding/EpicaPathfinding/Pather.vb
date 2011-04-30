Option Strict On

Public Class Pather
    Private Const ml_INITIAL_LIST_SIZE As Int32 = 100

    Private Const ml_BLOCK_ARRAY_BLOCK_SIZE As Int32 = 50

    Private Enum NodeListType As Byte
        NotOnAList = 0
        OnOpenList = 1
        OnClosedList = 2
    End Enum

    Private Class NodeData
        Public yList As Byte

        Public FVar As Single
        Public GVar As Single
        Public HVar As Single

        Public LocX As Byte
        Public LocY As Byte

        'location of the node before this node in the linked list
        Public PrevX As Byte
        Public PrevY As Byte

        Public HMHeight As Byte

        Public yDirection As Byte       'direction from parent

        Public TileType As Byte         '0 normal, 1 Water
    End Class

    Private moNodes(,) As NodeData
    Private myNodeUsed(,) As Byte
    Private mlCellSpacing As Int32

    Private mlOpenListX() As Int32
    Private mlOpenListY() As Int32
    Private mlOpenListUB As Int32 = -1

    Private moHm() As Byte
    Private miPtGrp() As Int16
    Private mlWaterHeight As Int32

    Public lNewCnt As Int32
    Public lTotalNew As Int32

    Public bOpenListMerge As Boolean

    Public Sub New(ByRef oTerrain As TerrainClass)
        ReDim moNodes(TerrainClass.Width - 1, TerrainClass.Height - 1)
        ReDim myNodeUsed(TerrainClass.Width - 1, TerrainClass.Height - 1)

        For Y As Int32 = 0 To TerrainClass.Height - 1
            For X As Int32 = 0 To TerrainClass.Width - 1
                myNodeUsed(X, Y) = 0
            Next X
        Next Y

        moHm = oTerrain.HeightMap

        mlCellSpacing = oTerrain.CellSpacing
        mlWaterHeight = oTerrain.WaterHeight

        SetLandmassGroups()
    End Sub

    Private Sub SetLandmassGroups()
        ReDim miPtGrp(moHm.GetUpperBound(0))

        Dim iCurrGrp As Int16 = 1
        For Y As Int32 = 0 To TerrainClass.Height - 1
            For X As Int32 = 0 To TerrainClass.Width - 1
                Dim lIdx As Int32 = Y * TerrainClass.Height + X

                'Ok... is this point above water?
                If moHm(lIdx) > mlWaterHeight Then
                    'Yes, Ok, is already assigned a group?
                    If miPtGrp(lIdx) = 0 Then
                        miPtGrp(lIdx) = iCurrGrp
                        iCurrGrp += 1S
                    End If
                    'Now, set the 8 squares around it if they are above water
                    For lSubY As Int32 = -1 To 1
                        For lSubX As Int32 = -1 To 1
                            Dim lTmpX As Int32 = X + lSubX
                            Dim lTmpY As Int32 = Y + lSubY

                            'Handle map wrap
                            If lTmpX < 0 Then
                                lTmpX = 0 ' TerrainClass.Width - 1
                            ElseIf lTmpX > TerrainClass.Width - 1 Then
                                lTmpX = TerrainClass.Width - 1 '0
                            End If
                            If lTmpY > -1 AndAlso lTmpY < TerrainClass.Height Then
                                Dim lTmpIdx As Int32 = lTmpY * TerrainClass.Height + lTmpX
                                'Ok, is this square above water?
                                If moHm(lTmpIdx) > mlWaterHeight Then
                                    'does the square = 0?
                                    If miPtGrp(lTmpIdx) = 0 Then
                                        'all is well, set this square to current grp
                                        miPtGrp(lTmpIdx) = miPtGrp(lIdx)
                                    ElseIf miPtGrp(lTmpIdx) <> miPtGrp(lIdx) Then
                                        'Ok, not good, found another group, who is the lower group number?
                                        Dim iReplaceGrp As Int16 = Math.Max(miPtGrp(lTmpIdx), miPtGrp(lIdx))
                                        Dim iWithGrp As Int16 = Math.Min(miPtGrp(lTmpIdx), miPtGrp(lIdx))
                                        ReplacePtGroup(iReplaceGrp, iWithGrp)
                                    End If
                                End If
                            End If
                        Next lSubX
                    Next lSubY
                Else
                    If miPtGrp(lIdx) = 0 Then
                        miPtGrp(lIdx) = iCurrGrp
                        iCurrGrp += 1S
                    End If
                    'Now, set the 8 squares around it if they are below water
                    For lSubY As Int32 = -1 To 1
                        For lSubX As Int32 = -1 To 1
                            Dim lTmpX As Int32 = X + lSubX
                            Dim lTmpY As Int32 = Y + lSubY

                            'Handle map wrap
                            If lTmpX < 0 Then
                                lTmpX = 0 'TerrainClass.Width - 1
                            ElseIf lTmpX > TerrainClass.Width - 1 Then
                                lTmpX = TerrainClass.Width - 1 '0
                            End If
                            If lTmpY > -1 AndAlso lTmpY < TerrainClass.Height Then
                                Dim lTmpIdx As Int32 = lTmpY * TerrainClass.Height + lTmpX
                                'Ok, is this square below water?
                                If moHm(lTmpIdx) <= mlWaterHeight Then
                                    'does the square = 0?
                                    If miPtGrp(lTmpIdx) = 0 Then
                                        'all is well, set this square to current grp
                                        miPtGrp(lTmpIdx) = miPtGrp(lIdx)
                                    ElseIf miPtGrp(lTmpIdx) <> miPtGrp(lIdx) Then
                                        'Ok, not good, found another group, who is the lower group number?
                                        Dim iReplaceGrp As Int16 = Math.Max(miPtGrp(lTmpIdx), miPtGrp(lIdx))
                                        Dim iWithGrp As Int16 = Math.Min(miPtGrp(lTmpIdx), miPtGrp(lIdx))
                                        ReplacePtGroup(iReplaceGrp, iWithGrp)
                                    End If
                                End If
                            End If
                        Next lSubX
                    Next lSubY
                End If
            Next X
        Next Y
    End Sub
    'This routine simply replaces a group with another group throughout the entire array
    Private Sub ReplacePtGroup(ByVal iReplace As Int16, ByVal iWithGrp As Int16)
        For X As Int32 = 0 To TerrainClass.Width - 1
            For Y As Int32 = 0 To TerrainClass.Height - 1
                Dim lIdx As Int32 = Y * TerrainClass.Height + X
                If miPtGrp(lIdx) = iReplace Then miPtGrp(lIdx) = iWithGrp
            Next Y
        Next X
    End Sub

    Private Sub InitNode(ByVal lX As Int32, ByVal lY As Int32)
        lNewCnt += 1
        lTotalNew += 1

        moNodes(lX, lY) = New NodeData
        With moNodes(lX, lY)
            Dim lIdx As Int32 = lY * TerrainClass.Height + lX

            .FVar = 0
            .HVar = 0
            .GVar = 0
            .LocX = CByte(lX)
            .LocY = CByte(lY)
            .PrevX = 255
            .PrevY = 255
            .HMHeight = moHm(lIdx)
            .yList = NodeListType.NotOnAList

            If .HMHeight < mlWaterHeight Then
                .TileType = 1
            Else : .TileType = 0
            End If

            For lTmpY As Int32 = lY - 1 To lY + 1
                For lTmpX As Int32 = lX - 1 To lX + 1
                    If lTmpX <> lX OrElse lTmpY <> lY Then
                        If lTmpX > -1 AndAlso lTmpX < TerrainClass.Width Then
                            If lTmpY > -1 AndAlso lTmpY < TerrainClass.Height Then
                                Dim lTmpIdx As Int32 = lTmpY * TerrainClass.Height + lTmpX
                                If .HMHeight < mlWaterHeight Then
                                    If moHm(lTmpIdx) > mlWaterHeight Then
                                        .TileType = 2
                                        Exit For
                                    End If
                                Else
                                    If moHm(lTmpIdx) < mlWaterHeight Then
                                        .TileType = 2
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next lTmpX
                If .TileType = 2 Then Exit For
            Next lTmpY

        End With
        myNodeUsed(lX, lY) = 255
    End Sub

    'Public Function GetUsedPoints(ByRef ptRet() As Point, ByRef lOpenListUB As Int32) As Point()

    '    Dim ptResult(mlOpenListUB) As Point
    '    Dim lPtResUB As Int32 = mlOpenListUB

    '    lOpenListUB = mlOpenListUB

    '    For X As Int32 = 0 To mlOpenListUB
    '        ptResult(X).X = mlOpenListX(X)
    '        ptResult(X).Y = mlOpenListY(X)
    '    Next X

    '    For Y As Int32 = 0 To TerrainClass.Height - 1
    '        For X As Int32 = 0 To TerrainClass.Width - 1
    '            If myNodeUsed(X, Y) <> 0 Then
    '                If moNodes(X, Y).yList = NodeListType.OnClosedList Then
    '                    Dim bFound As Boolean = False
    '                    For lIdx As Int32 = 0 To ptRet.GetUpperBound(0)
    '                        If ptRet(lIdx).X = X AndAlso ptRet(lIdx).Y = Y Then
    '                            bFound = True
    '                        End If
    '                    Next
    '                    If bFound = False Then
    '                        lPtResUB += 1
    '                        ReDim Preserve ptResult(lPtResUB)
    '                        ptResult(lPtResUB).X = X
    '                        ptResult(lPtResUB).Y = Y
    '                    End If
    '                End If
    '            End If
    '        Next X
    '    Next Y

    '    Return ptResult
    'End Function

    Private ml_MAX_LOOPS As Int32 = 15000
    Public Function GetPathNaval(ByVal lStartX As Int32, ByVal lStartY As Int32, ByVal iStartAngle As Int16, ByVal lEndX As Int32, ByVal lEndY As Int32) As System.Drawing.Point()
        lNewCnt = 0

        'Ok, first, we would check if the locs are within bounds and adjust on the map wrap accordingly. But for now, we'll just return nothing
        If lStartX < 0 Then lStartX = 0 'lStartX += TerrainClass.Width
        If lStartX > TerrainClass.Width - 1 Then lStartX = TerrainClass.Width - 1 ' lStartX -= TerrainClass.Width
        If lEndX < 0 Then lEndX = 0 ' lEndX += TerrainClass.Width
        If lEndX > TerrainClass.Width - 1 Then lEndX = TerrainClass.Width - 1 ' lEndX -= TerrainClass.Width
        If lStartX < 0 OrElse lStartY < 0 OrElse lEndX < 0 OrElse lEndY < 0 Then Return Nothing
        'If lStartX > TerrainClass.Width - 1 OrElse lStartY > TerrainClass.Height - 1 Then Return Nothing
        'If lEndX > TerrainClass.Width - 1 OrElse lEndY > TerrainClass.Height - 1 Then Return Nothing

        'Ok, check our pt groups
        Dim lInitialPointIdx As Int32 = lStartY * TerrainClass.Width + lStartX
        Dim lFinalPointIdx As Int32 = lEndY * TerrainClass.Width + lEndX
        If miPtGrp(lInitialPointIdx) <> miPtGrp(lFinalPointIdx) Then Return Nothing

        Dim lCurrentDirection As Int32

        Dim bInitialDirection As Boolean = True

        If iStartAngle > 3600 Then iStartAngle -= 3600S
        If iStartAngle < 0 Then iStartAngle += 3600S
        lCurrentDirection = iStartAngle \ 450I

        'Now, set all tiles to No list, and the f and g values to int.max
        For Y As Int32 = 0 To TerrainClass.Height - 1
            For X As Int32 = 0 To TerrainClass.Width - 1
                If myNodeUsed(X, Y) <> 0 Then
                    moNodes(X, Y).yList = NodeListType.NotOnAList
                    moNodes(X, Y).FVar = Int32.MaxValue
                    moNodes(X, Y).GVar = Int32.MaxValue
                End If
            Next X
        Next Y

        'Reset our open list (ensure it is clear)
        ResetOpenList()

        'Init our starting point
        If myNodeUsed(lStartX, lStartY) = 0 Then InitNode(lStartX, lStartY)
        With moNodes(lStartX, lStartY)
            'set our variables
            .GVar = 0
            .HVar = Math.Abs(lStartX - lEndX) + Math.Abs(lStartY - lEndY)
            .FVar = .HVar

            'set our open state
            .yList = NodeListType.OnOpenList

            .yDirection = CByte(lCurrentDirection)
        End With

        'Determine if our user is an idiot...
        If myNodeUsed(lEndX, lEndY) = 0 Then InitNode(lEndX, lEndY)
        'Is the start loc water-based (0 or 2)
        If moNodes(lStartX, lStartY).TileType = 1 OrElse moNodes(lStartX, lStartY).TileType = 2 Then
            'It is, so is the end loc ground-based? (0)
            If moNodes(lEndX, lEndY).TileType = 0 Then
                'it is, so, move the end loc towards the start loc until end loc is on water
                Dim ptTemp As Point = GetClosestTileType(lStartX, lStartY, lEndX, lEndY, 1)
                If ptTemp.X = Int32.MinValue OrElse ptTemp.Y = Int32.MinValue Then Return Nothing
                lEndX = ptTemp.X
                lEndY = ptTemp.Y
            End If
        Else
            'No, start loc is ground-based (0), so is the end loc ground based? (0 only because 2 is edge)
            If moNodes(lEndX, lEndY).TileType = 0 Then
                'It is, so move the end loc towards the start loc until the end loc is on water
                Dim ptTemp As Point = GetClosestTileType(lStartX, lStartY, lEndX, lEndY, 1)
                If ptTemp.X = Int32.MinValue OrElse ptTemp.Y = Int32.MinValue Then Return Nothing
                lEndX = ptTemp.X
                lEndY = ptTemp.Y
            End If
        End If

        'put our node into the open list
        PushToOpenList(lStartX, lStartY)

        'Now, search until we have reached the end or our open list is empty
        Dim lLoops As Int32 = 0
        While mlOpenListUB > -1

            lLoops += 1
            If lLoops > ml_MAX_LOOPS Then Return Nothing

            'Get our ACTIVE node
            Dim lOpenListIdx As Int32 = GetLowestFVal()
            If lOpenListIdx = -1 Then
                'Clear our open list
                ResetOpenList()
                Return Nothing
            End If

            Dim oActive As NodeData = moNodes(mlOpenListX(lOpenListIdx), mlOpenListY(lOpenListIdx))
            PopFromOpenList(lOpenListIdx)

            'Remove the node from the OPEN list
            oActive.yList = NodeListType.NotOnAList

            lCurrentDirection = oActive.yDirection

            If oActive.LocX = lEndX AndAlso oActive.LocY = lEndY Then
                'Ok, we have found it, generate and return our path
                Dim lCnt As Int32 = 0

                Dim oNode As NodeData = oActive
                While oNode.PrevX <> 255 AndAlso oNode.PrevY <> 255 AndAlso (oNode.LocX <> lStartX OrElse oNode.LocY <> lStartY)
                    If lCnt > 2000 Then
                        ResetOpenList()
                        Return Nothing
                    End If

                    lCnt += 1
                    oNode = moNodes(oNode.PrevX, oNode.PrevY)
                End While

                lCnt -= 1
                Dim ptReturn(lCnt) As System.Drawing.Point

                oNode = oActive
                While oNode.PrevX <> 255 AndAlso oNode.PrevY <> 255 AndAlso (oNode.LocX <> lStartX OrElse oNode.LocY <> lStartY)
                    ptReturn(lCnt).X = oNode.LocX
                    ptReturn(lCnt).Y = oNode.LocY
                    oNode = moNodes(oNode.PrevX, oNode.PrevY)
                    lCnt -= 1
                End While

                'Clear our open list
                ResetOpenList()

                Return ptReturn
            Else
                'Determine what direction the end result is located...
                Dim lDirToEnd As Int32
                Dim lDirToEndX As Int32 = oActive.LocX - lEndX
                Dim lDirToEndY As Int32 = oActive.LocY - lEndY

                If lDirToEndX = 0 Then
                    If lDirToEndY > 0 Then
                        lDirToEnd = 2
                    Else : lDirToEnd = 6
                    End If
                ElseIf lDirToEndX > 0 Then
                    If lDirToEndY = 0 Then
                        lDirToEnd = 4
                    ElseIf lDirToEndY > 0 Then
                        lDirToEnd = 3
                    Else : lDirToEnd = 5
                    End If
                Else
                    If lDirToEndY = 0 Then
                        lDirToEnd = 0
                    ElseIf lDirToEndY > 0 Then
                        lDirToEnd = 1
                    Else : lDirToEnd = 7
                    End If
                End If

                'For each of the 8 neighbors of oActive...
                For Y As Int32 = -1 To 1
                    Dim lNodeY As Int32 = oActive.LocY + Y

                    If lNodeY < 0 Then Continue For
                    If lNodeY > TerrainClass.Height - 1 Then Continue For

                    For X As Int32 = -1 To 1
                        If X = 0 AndAlso Y = 0 Then Continue For

                        Dim lNodeX As Int32 = oActive.LocX + X

                        'If lNodeX < 0 Then lNodeX += TerrainClass.Width
                        If lNodeX < 0 Then lNodeX = 0
                        'If lNodeX > TerrainClass.Width - 1 Then lNodeX -= TerrainClass.Width
                        If lNodeX > TerrainClass.Width - 1 Then lNodeX = TerrainClass.Width - 1

                        '   bInList = False
                        Dim bInList As Boolean = False

                        'Calc new scores for neighbor
                        If myNodeUsed(lNodeX, lNodeY) = 0 Then InitNode(lNodeX, lNodeY)
                        Dim oNeighbor As NodeData = moNodes(lNodeX, lNodeY)
                        Dim lNeighborDirection As Int32
                        Dim fNewGVar As Single = oActive.GVar '0.0F
                        Dim fNewFVar As Single = 0.0F
                        Dim fNewHVar As Single = 0.0F

                        'First, should we even consider this node?
                        If oNeighbor.TileType = 1 OrElse _
                          (oNeighbor.TileType = 2 AndAlso _
                             ((Math.Abs(oNeighbor.LocX - lEndX) < 3 AndAlso Math.Abs(oNeighbor.LocY - lEndY) < 3) _
                                OrElse (Math.Abs(oNeighbor.LocX - lStartX) < 3 AndAlso Math.Abs(oNeighbor.LocY - lStartY) < 3))) Then

                            If X = 0 Then
                                If Y > 0 Then
                                    lNeighborDirection = 6
                                Else : lNeighborDirection = 2
                                End If
                            ElseIf X > 0 Then
                                If Y = 0 Then
                                    lNeighborDirection = 0
                                ElseIf Y > 0 Then
                                    lNeighborDirection = 7
                                Else : lNeighborDirection = 1
                                End If
                            Else
                                If Y = 0 Then
                                    lNeighborDirection = 4
                                ElseIf Y > 0 Then
                                    lNeighborDirection = 5
                                Else : lNeighborDirection = 3
                                End If
                            End If

                            '       If neighbor is diagonal, GCost = 1.414, else GCost = 1
                            'If Math.Abs(X) = 1 AndAlso Math.Abs(Y) = 1 Then
                            fNewGVar += 1.414F
                            'Else : fNewGVar += 2.0F
                            'End If



                            'If angle to neighbor is different, modify by TurnCost 
                            Dim lDiff As Int32 = 0
                            If lCurrentDirection <> lNeighborDirection AndAlso bInitialDirection = False Then
                                lDiff = Math.Abs(lNeighborDirection - lCurrentDirection)
                                If lDiff > 4 Then
                                    Dim lTmpDir As Int32 = lCurrentDirection - 4
                                    Dim lTmpDir2 As Int32 = lNeighborDirection - 4
                                    If lTmpDir < 0 Then lTmpDir += 8
                                    If lTmpDir2 < 0 Then lTmpDir2 += 8
                                    lDiff = Math.Abs(lTmpDir - lTmpDir2)
                                End If

                                ''Now...
                                If lNeighborDirection = lDirToEnd Then
                                    'we give a bonus for this
                                    lDiff -= 2
                                    If lDiff < 0 Then lDiff = 0
                                End If

                                fNewGVar += (lDiff * 3.0F)
                            End If

                            '       Modify GCost by HeightCost
                            Dim fTmpVar As Single = ((CInt(oNeighbor.HMHeight) - CInt(oActive.HMHeight)) + 4) / 2.0F
                            If fTmpVar > 0 Then fNewGVar += fTmpVar

                            'If oNeighbor.LocX < lEndX Then
                            '    fTmpVar = (oNeighbor.LocX + TerrainClass.Width) - lEndX
                            'Else : fTmpVar = (lEndX + TerrainClass.Width) - oNeighbor.LocX
                            'End If
                            fTmpVar = oNeighbor.LocX - lEndX
                            Dim fTmpx As Single = Math.Min(Math.Abs(lEndX - oNeighbor.LocX), fTmpVar)
                            fTmpx *= fTmpx
                            Dim fTmpY As Single = oNeighbor.LocY - lEndY
                            fTmpY *= fTmpY
                            fNewHVar = CSng(Math.Sqrt(fTmpx + fTmpY)) * 5.0F

                            '       NewFVar = GCost + HCost
                            fNewFVar = fNewGVar + fNewHVar

                            '   If Neighbor is on Open or Closed list...
                            If oNeighbor.yList = NodeListType.OnOpenList Then

                                If oNeighbor.GVar > fNewGVar Then
                                    '       If NewFVar < Neighbor.FVar
                                    'If fNewFVar < oNeighbor.FVar Then
                                    '           Neighbor.g = GCost
                                    oNeighbor.GVar = fNewGVar
                                    '           Neighbor.f = NewFVar
                                    oNeighbor.FVar = fNewFVar
                                    '           Neighbor.Parent = oActive
                                    oNeighbor.HVar = fNewHVar

                                    If bOpenListMerge = False Then
                                        oNeighbor.PrevX = oActive.LocX
                                        oNeighbor.PrevY = oActive.LocY
                                    End If

                                    oNeighbor.yDirection = CByte(lNeighborDirection)

                                End If


                                '       End If

                                '       bInList = True
                                bInList = True
                            ElseIf oNeighbor.yList = NodeListType.OnClosedList Then
                                bInList = True
                            End If
                            '   End IF

                            '   If bInList = False Then
                            If bInList = False Then
                                With oNeighbor
                                    '       Neighbor.f = NewFVar
                                    .FVar = fNewFVar
                                    '       Neighbor.g = GCost
                                    .GVar = fNewGVar
                                    '       Neighbor.Parent = oActive
                                    .HVar = fNewHVar
                                    .PrevX = oActive.LocX
                                    .PrevY = oActive.LocY
                                    '       Neighbor.yList = NodeListType.OnOpenList
                                    .yList = NodeListType.OnOpenList

                                    .yDirection = CByte(lNeighborDirection)
                                End With
                                PushToOpenList(oNeighbor.LocX, oNeighbor.LocY)
                            End If

                        End If
                    Next X
                Next Y
                'Next

                oActive.yList = NodeListType.OnClosedList

                bInitialDirection = False
            End If
        End While

        Return Nothing
    End Function

    ''' <summary>
    ''' This function plots a path along the terrain (the 240x240) grid taking height, water, etc... into consideration. Returns a Point Array.
    ''' </summary>
    ''' <param name="lStartX"> The Starting Grid Loc X (0 to 240) </param>
    ''' <param name="lStartY"> The Starting Grid Loc Y (0 to 240)</param>
    ''' <param name="iStartAngle"> The entity's current angle (0 to 3600) </param>
    ''' <param name="lEndX"> The End Grid Loc X (0 to 240) </param>
    ''' <param name="lEndY"> The End Grid Loc Y (0 to 240) </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPath(ByVal lStartX As Int32, ByVal lStartY As Int32, ByVal iStartAngle As Int16, ByVal lEndX As Int32, ByVal lEndY As Int32) As System.Drawing.Point()
        lNewCnt = 0

        'Ok, first, we would check if the locs are within bounds and adjust on the map wrap accordingly. But for now, we'll just return nothing
        'If lStartX < 0 Then lStartX += TerrainClass.Width
        If lStartX < 0 Then lStartX = 0
        If lStartX > TerrainClass.Width - 1 Then lStartX = TerrainClass.Width - 1 'lStartX -= TerrainClass.Width
        If lEndX < 0 Then lEndX = 0 ' lEndX += TerrainClass.Width
        If lEndX > TerrainClass.Width - 1 Then lEndX = TerrainClass.Width - 1 'lEndX -= TerrainClass.Width
        If lStartX < 0 OrElse lStartY < 0 OrElse lEndX < 0 OrElse lEndY < 0 Then Return Nothing
        'If lStartX > TerrainClass.Width - 1 OrElse lStartY > TerrainClass.Height - 1 Then Return Nothing
        'If lEndX > TerrainClass.Width - 1 OrElse lEndY > TerrainClass.Height - 1 Then Return Nothing
        If lEndY > TerrainClass.Height - 1 OrElse lEndY < 0 Then Return Nothing
        If lStartY > TerrainClass.Height - 1 OrElse lStartY < 0 Then Return Nothing

        'Ok, check our pt groups
        Dim lInitialPointIdx As Int32 = lStartY * TerrainClass.Width + lStartX
        Dim lFinalPointIdx As Int32 = lEndY * TerrainClass.Width + lEndX
        If miPtGrp(lInitialPointIdx) <> miPtGrp(lFinalPointIdx) Then Return Nothing

        Dim lCurrentDirection As Int32

        Dim bInitialDirection As Boolean = True

        If iStartAngle > 3600 Then iStartAngle -= 3600S
        If iStartAngle < 0 Then iStartAngle += 3600S
        lCurrentDirection = iStartAngle \ 450I

        'Now, set all tiles to No list, and the f and g values to int.max
        For Y As Int32 = 0 To TerrainClass.Height - 1
            For X As Int32 = 0 To TerrainClass.Width - 1
                If myNodeUsed(X, Y) <> 0 Then
                    moNodes(X, Y).yList = NodeListType.NotOnAList
                    moNodes(X, Y).FVar = Int32.MaxValue
                    moNodes(X, Y).GVar = Int32.MaxValue
                End If
            Next X
        Next Y

        'Reset our open list (ensure it is clear)
        ResetOpenList()

        'Init our starting point
        If myNodeUsed(lStartX, lStartY) = 0 Then InitNode(lStartX, lStartY)
        With moNodes(lStartX, lStartY)
            'set our variables
            .GVar = 0
            .HVar = Math.Abs(lStartX - lEndX) + Math.Abs(lStartY - lEndY)
            .FVar = .HVar

            'set our open state
            .yList = NodeListType.OnOpenList

            .yDirection = CByte(lCurrentDirection)
        End With

        'Determine if our user is an idiot...
        If myNodeUsed(lEndX, lEndY) = 0 Then InitNode(lEndX, lEndY)
        'Is the start loc ground-based (0 or 2)
        If moNodes(lStartX, lStartY).TileType = 0 OrElse moNodes(lStartX, lStartY).TileType = 2 Then
            'It is, so is the end loc water-based? (1)
            If moNodes(lEndX, lEndY).TileType = 1 Then
                'it is, so, move the end loc towards the start loc until end loc is on ground
                Dim ptTemp As Point = GetClosestTileType(lStartX, lStartY, lEndX, lEndY, 0)
                If ptTemp.X = Int32.MinValue OrElse ptTemp.Y = Int32.MinValue Then Return Nothing
                lEndX = ptTemp.X
                lEndY = ptTemp.Y
            End If
        Else
            'No, start loc is water-based (1), so is the end loc ground based? (0 only because 2 is edge)
            If moNodes(lEndX, lEndY).TileType = 0 Then
                'It is, so move the end loc towards the start loc until the end loc is on water
                Dim ptTemp As Point = GetClosestTileType(lStartX, lStartY, lEndX, lEndY, 1)
                If ptTemp.X = Int32.MinValue OrElse ptTemp.Y = Int32.MinValue Then Return Nothing
                lEndX = ptTemp.X
                lEndY = ptTemp.Y
            End If
        End If

        'put our node into the open list
        PushToOpenList(lStartX, lStartY)

        'Now, search until we have reached the end or our open list is empty
        While mlOpenListUB > -1
            'Get our ACTIVE node
            Dim lOpenListIdx As Int32 = GetLowestFVal()
            If lOpenListIdx = -1 Then
                'Clear our open list
                ResetOpenList()
                Return Nothing
            End If

            Dim oActive As NodeData = moNodes(mlOpenListX(lOpenListIdx), mlOpenListY(lOpenListIdx))
            PopFromOpenList(lOpenListIdx)

            'Remove the node from the OPEN list
            oActive.yList = NodeListType.NotOnAList

            lCurrentDirection = oActive.yDirection

            If oActive.LocX = lEndX AndAlso oActive.LocY = lEndY Then
                'Ok, we have found it, generate and return our path
                Dim lCnt As Int32 = 0

                Dim oNode As NodeData = oActive
                While oNode.PrevX <> 255 AndAlso oNode.PrevY <> 255 AndAlso (oNode.LocX <> lStartX OrElse oNode.LocY <> lStartY)
                    If lCnt > 2000 Then
                        ResetOpenList()
                        Return Nothing
                    End If

                    lCnt += 1
                    oNode = moNodes(oNode.PrevX, oNode.PrevY)
                End While

                lCnt -= 1
                Dim ptReturn(lCnt) As System.Drawing.Point

                oNode = oActive
                While oNode.PrevX <> 255 AndAlso oNode.PrevY <> 255 AndAlso (oNode.LocX <> lStartX OrElse oNode.LocY <> lStartY)
                    ptReturn(lCnt).X = oNode.LocX
                    ptReturn(lCnt).Y = oNode.LocY
                    oNode = moNodes(oNode.PrevX, oNode.PrevY)
                    lCnt -= 1
                End While

                'Clear our open list
                ResetOpenList()

                Return ptReturn
            Else
                'Determine what direction the end result is located...
                Dim lDirToEnd As Int32
                Dim lDirToEndX As Int32 = oActive.LocX - lEndX
                Dim lDirToEndY As Int32 = oActive.LocY - lEndY

                If lDirToEndX = 0 Then
                    If lDirToEndY > 0 Then
                        lDirToEnd = 2
                    Else : lDirToEnd = 6
                    End If
                ElseIf lDirToEndX > 0 Then
                    If lDirToEndY = 0 Then
                        lDirToEnd = 4
                    ElseIf lDirToEndY > 0 Then
                        lDirToEnd = 3
                    Else : lDirToEnd = 5
                    End If
                Else
                    If lDirToEndY = 0 Then
                        lDirToEnd = 0
                    ElseIf lDirToEndY > 0 Then
                        lDirToEnd = 1
                    Else : lDirToEnd = 7
                    End If
                End If

                'For each of the 8 neighbors of oActive...
                For Y As Int32 = -1 To 1
                    Dim lNodeY As Int32 = oActive.LocY + Y

                    If lNodeY < 0 Then Continue For
                    If lNodeY > TerrainClass.Height - 1 Then Continue For

                    For X As Int32 = -1 To 1
                        If X = 0 AndAlso Y = 0 Then Continue For

                        Dim lNodeX As Int32 = oActive.LocX + X

                        'If lNodeX < 0 Then continue For
                        If lNodeX < 0 Then lNodeX += TerrainClass.Width
                        'If lNodeX > TerrainClass.Width - 1 Then Continue For
                        If lNodeX > TerrainClass.Width - 1 Then lNodeX -= TerrainClass.Width

                        '   bInList = False
                        Dim bInList As Boolean = False

                        'Calc new scores for neighbor
                        If myNodeUsed(lNodeX, lNodeY) = 0 Then InitNode(lNodeX, lNodeY)
                        Dim oNeighbor As NodeData = moNodes(lNodeX, lNodeY)
                        Dim lNeighborDirection As Int32
                        Dim fNewGVar As Single = oActive.GVar '0.0F
                        Dim fNewFVar As Single = 0.0F
                        Dim fNewHVar As Single = 0.0F

                        'First, should we even consider this node?
                        If oNeighbor.TileType = 0 OrElse _
                          (oNeighbor.TileType = 2 AndAlso _
                             ((Math.Abs(oNeighbor.LocX - lEndX) < 3 AndAlso Math.Abs(oNeighbor.LocY - lEndY) < 3) _
                                OrElse (Math.Abs(oNeighbor.LocX - lStartX) < 3 AndAlso Math.Abs(oNeighbor.LocY - lStartY) < 3))) Then

                            If X = 0 Then
                                If Y > 0 Then
                                    lNeighborDirection = 6
                                Else : lNeighborDirection = 2
                                End If
                            ElseIf X > 0 Then
                                If Y = 0 Then
                                    lNeighborDirection = 0
                                ElseIf Y > 0 Then
                                    lNeighborDirection = 7
                                Else : lNeighborDirection = 1
                                End If
                            Else
                                If Y = 0 Then
                                    lNeighborDirection = 4
                                ElseIf Y > 0 Then
                                    lNeighborDirection = 5
                                Else : lNeighborDirection = 3
                                End If
                            End If

                            '       If neighbor is diagonal, GCost = 1.414, else GCost = 1
                            'If Math.Abs(X) = 1 AndAlso Math.Abs(Y) = 1 Then
                            fNewGVar += 1.414F
                            'Else : fNewGVar += 2.0F
                            'End If



                            'If angle to neighbor is different, modify by TurnCost 
                            Dim lDiff As Int32 = 0
                            If lCurrentDirection <> lNeighborDirection AndAlso bInitialDirection = False Then
                                lDiff = Math.Abs(lNeighborDirection - lCurrentDirection)
                                If lDiff > 4 Then
                                    Dim lTmpDir As Int32 = lCurrentDirection - 4
                                    Dim lTmpDir2 As Int32 = lNeighborDirection - 4
                                    If lTmpDir < 0 Then lTmpDir += 8
                                    If lTmpDir2 < 0 Then lTmpDir2 += 8
                                    lDiff = Math.Abs(lTmpDir - lTmpDir2)
                                End If

                                ''Now...
                                If lNeighborDirection = lDirToEnd Then
                                    'we give a bonus for this
                                    lDiff -= 2
                                    If lDiff < 0 Then lDiff = 0
                                End If

                                fNewGVar += (lDiff * 3.0F)
                            End If

                            '       Modify GCost by HeightCost
                            Dim fTmpVar As Single = ((CInt(oNeighbor.HMHeight) - CInt(oActive.HMHeight)) + 4) / 2.0F
                            If fTmpVar > 0 Then fNewGVar += fTmpVar

                            '       HCost = Math.Abs(lNeighborX - lEndX) + Math.Abs(lNeighborY - lEndY)
                            'fNewHVar = (Math.Abs(oNeighbor.LocX - lEndX) + Math.Abs(oNeighbor.LocY - lEndY)) * 5.0F
                            'Dim fTmpx As Single = oNeighbor.LocX - lEndX
                            ''For map wrapping, we need the absolute distance...
                            'fTmpVar = oNeighbor.LocX - Math.Abs(lEndX - TerrainClass.Width)
                            'If Math.Abs(fTmpx) > Math.Abs(fTmpVar) Then fTmpx = fTmpVar

                            'If oNeighbor.LocX < lEndX Then
                            '    fTmpVar = (oNeighbor.LocX + TerrainClass.Width) - lEndX
                            'Else : fTmpVar = (lEndX + TerrainClass.Width) - oNeighbor.LocX
                            'End If
                            fTmpVar = oNeighbor.LocX - lEndX
                            Dim fTmpx As Single = Math.Min(Math.Abs(lEndX - oNeighbor.LocX), fTmpVar)
                            fTmpx *= fTmpx
                            Dim fTmpY As Single = oNeighbor.LocY - lEndY
                            fTmpY *= fTmpY
                            fNewHVar = CSng(Math.Sqrt(fTmpx + fTmpY)) * 5.0F

                            '       NewFVar = GCost + HCost
                            fNewFVar = fNewGVar + fNewHVar

                            '   If Neighbor is on Open or Closed list...
                            If oNeighbor.yList = NodeListType.OnOpenList Then

                                If oNeighbor.GVar > fNewGVar Then
                                    '       If NewFVar < Neighbor.FVar
                                    'If fNewFVar < oNeighbor.FVar Then
                                    '           Neighbor.g = GCost
                                    oNeighbor.GVar = fNewGVar
                                    '           Neighbor.f = NewFVar
                                    oNeighbor.FVar = fNewFVar
                                    '           Neighbor.Parent = oActive
                                    oNeighbor.HVar = fNewHVar

                                    If bOpenListMerge = False Then
                                        oNeighbor.PrevX = oActive.LocX
                                        oNeighbor.PrevY = oActive.LocY
                                    End If

                                    oNeighbor.yDirection = CByte(lNeighborDirection)

                                End If


                                '       End If

                                '       bInList = True
                                bInList = True
                            ElseIf oNeighbor.yList = NodeListType.OnClosedList Then
                                bInList = True
                            End If
                            '   End IF

                            '   If bInList = False Then
                            If bInList = False Then
                                With oNeighbor
                                    '       Neighbor.f = NewFVar
                                    .FVar = fNewFVar
                                    '       Neighbor.g = GCost
                                    .GVar = fNewGVar
                                    '       Neighbor.Parent = oActive
                                    .HVar = fNewHVar
                                    .PrevX = oActive.LocX
                                    .PrevY = oActive.LocY
                                    '       Neighbor.yList = NodeListType.OnOpenList
                                    .yList = NodeListType.OnOpenList

                                    .yDirection = CByte(lNeighborDirection)
                                End With
                                PushToOpenList(oNeighbor.LocX, oNeighbor.LocY)
                            End If

                        End If
                    Next X
                Next Y
                'Next

                oActive.yList = NodeListType.OnClosedList

                bInitialDirection = False
            End If
        End While

        Return Nothing
    End Function


    Private Function GetLowestFVal() As Int32
        Dim fLowest As Single = Int32.MaxValue
        Dim fLowests_H As Single
        Dim lIdx As Int32 = -1

        For X As Int32 = 0 To mlOpenListUB
            Dim fVar As Single = moNodes(mlOpenListX(X), mlOpenListY(X)).FVar

            If fVar < fLowest Then
                fLowest = fVar
                fLowests_H = moNodes(mlOpenListX(X), mlOpenListY(X)).HVar
                lIdx = X
            ElseIf fVar = fLowest Then
                If moNodes(mlOpenListX(X), mlOpenListY(X)).HVar < fLowests_H Then
                    fLowest = fVar
                    fLowests_H = moNodes(mlOpenListX(X), mlOpenListY(X)).HVar
                    lIdx = X
                End If
            End If
        Next X

        Return lIdx
    End Function

    Private Sub ResetOpenList()
        ReDim mlOpenListX(ml_INITIAL_LIST_SIZE)
        ReDim mlOpenListY(ml_INITIAL_LIST_SIZE)
        mlOpenListUB = -1
    End Sub

    Private Sub PushToOpenList(ByVal lX As Int32, ByVal lY As Int32)
        mlOpenListUB += 1
        If mlOpenListUB > mlOpenListX.GetUpperBound(0) Then
            ReDim Preserve mlOpenListX(mlOpenListX.Length + ml_INITIAL_LIST_SIZE)
            ReDim Preserve mlOpenListY(mlOpenListY.Length + ml_INITIAL_LIST_SIZE)
        End If
        mlOpenListX(mlOpenListUB) = lX
        mlOpenListY(mlOpenListUB) = lY
    End Sub

    Private Sub PopFromOpenList(ByVal lIndex As Int32)
        'Ok, instead of shifting everything... we simply move our last item to the empty spot and change mlOpenListUB to -1
        If mlOpenListUB <> 0 AndAlso lIndex <> mlOpenListUB Then
            mlOpenListX(lIndex) = mlOpenListX(mlOpenListUB)
            mlOpenListY(lIndex) = mlOpenListY(mlOpenListUB)
        End If
        mlOpenListUB -= 1
    End Sub

    ''' <summary>
    ''' Takes the result of GetPath() and optimizes it. Then returns an array of points representing the optimized path.
    ''' </summary>
    ''' <param name="ptPath"> The result from GetPath </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function OptimizePath(ByRef ptPath() As System.Drawing.Point, ByVal ptStart As System.Drawing.Point, ByVal bNaval As Boolean) As System.Drawing.Point()
        If ptPath Is Nothing Then Return Nothing

        If ptPath.GetUpperBound(0) = 0 Then Return ptPath

        Dim ptReturn(ptPath.GetUpperBound(0)) As System.Drawing.Point
        'Dim yMapWrap(ptPath.GetUpperBound(0)) As Byte

        Dim lStartIdx As Int32 = 0
        Dim lEndIdx As Int32 = 1
        Dim lFinalIdx As Int32 = ptPath.GetUpperBound(0)
        Dim lDestIdx As Int32 = 1

        Dim lInitAngle As Int32
        Dim yPrevEdge As Byte = 0        '0 - not on edge, 1 - left edge, 2- right edge

        'MSC - 11/06/07 - added to fix the situation where start is on a map wrap situation
        'Dim bStartMapWrap As Boolean = False
        'Dim lStartMapWrapX As Int32 = -1
        'Dim ptFirstLoc As Point

        ''check for possible map scenario on the initial step
        'If ptStart.X = 0 Then
        '    If ptPath(0).X = TerrainClass.Width - 1 Then
        '        'initial map wrap scenario
        '        bStartMapWrap = True
        '        lStartMapWrapX = -1
        '        ptFirstLoc = ptPath(0)
        '    End If
        'ElseIf ptStart.X = TerrainClass.Width - 1 Then
        '    If ptPath(0).X = 0 Then
        '        'initial map wrap scenario
        '        bStartMapWrap = True
        '        lStartMapWrapX = TerrainClass.Width
        '        ptFirstLoc = ptPath(0)
        '    End If
        'End If
        'end of 11/06/07 add

        ptReturn(0) = ptPath(0)

        While lStartIdx <> lFinalIdx
            lInitAngle = LineAngleDegrees_Pts(ptPath(lStartIdx), ptPath(lEndIdx))

            While lEndIdx <> lFinalIdx AndAlso LineAngleDegrees_Pts(ptPath(lStartIdx), ptPath(lEndIdx)) = lInitAngle
                'yPrevEdge = 0
                'If ptPath(lEndIdx).X = 0 Then
                '    yPrevEdge = 1
                'ElseIf ptPath(lEndIdx).X = TerrainClass.Width - 1 Then
                '    yPrevEdge = 2
                'End If

                'Now, increment our end idx
                lEndIdx += 1

                'If yPrevEdge <> 0 AndAlso lEndIdx <> lFinalIdx Then
                '    If ptPath(lEndIdx).X = 0 AndAlso yPrevEdge = 2 Then
                '        'we crossed, so check our value now...
                '        Dim ptTmp As Point = ptPath(lEndIdx)
                '        ptTmp.X = TerrainClass.Width
                '        Dim lVal As Int32 = LineAngleDegrees_Pts(ptPath(lStartIdx), ptTmp)
                '        If lVal = lInitAngle Then
                '            lStartIdx = lEndIdx
                '            lEndIdx += 1
                '            'yMapWrap(lDestIdx) = 255
                '        End If
                '    ElseIf ptPath(lEndIdx).X = TerrainClass.Width - 1 AndAlso yPrevEdge = 1 Then
                '        'we crossed
                '        Dim ptTmp As Point = ptPath(lEndIdx)
                '        ptTmp.X = -1
                '        Dim lVal As Int32 = LineAngleDegrees_Pts(ptPath(lStartIdx), ptTmp)
                '        If lVal = lInitAngle Then
                '            lStartIdx = lEndIdx
                '            lEndIdx += 1
                '            'yMapWrap(lDestIdx) = 255
                '        End If
                '    End If
                'End If
            End While
            If lEndIdx = lFinalIdx Then
                ptReturn(lDestIdx) = ptPath(lEndIdx)
                lStartIdx = lEndIdx
            Else
                ptReturn(lDestIdx) = ptPath(lEndIdx - 1)
                lDestIdx += 1
                lStartIdx = lEndIdx - 1
            End If
        End While

        ReDim Preserve ptReturn(lDestIdx)
        'ReDim Preserve yMapWrap(lDestIdx)

        Dim fCalcCosts(lDestIdx) As Single

        'Ok, calculate our costs for the points
        For X As Int32 = 0 To lDestIdx - 1
            fCalcCosts(X) = GetCostsBetweenBasePoints(ptReturn(X), ptReturn(X + 1), False, bNaval) ' yMapWrap(lDestIdx) <> 0, bNaval)
        Next X

        Dim ptSuperOpt(lDestIdx) As System.Drawing.Point
        Dim lSuperOptUB As Int32 = -1

        ''====================== added for initial map wrap situation fix (11/06/07)
        'If bStartMapWrap = True Then
        '    If ptReturn(0) <> ptFirstLoc Then
        '        ReDim Preserve ptSuperOpt(lDestIdx + 1)
        '        lSuperOptUB += 1
        '        ptSuperOpt(lSuperOptUB) = ptFirstLoc
        '        ptSuperOpt(lSuperOptUB).X = lStartMapWrapX
        '    Else
        '        ptReturn(0).X = lStartMapWrapX
        '    End If
        'End If
        ''======================

        lSuperOptUB += 1
        ptSuperOpt(lSuperOptUB) = ptReturn(0)

        'Ok, go thru and compare our calculated result to the actual results...
        For X As Int32 = 0 To lDestIdx - 2
            Dim fCalculatedRouteCost As Single = 0.0F
            Dim lNextIndex As Int32 = X + 1

            Dim fEffRating As Single = 1.0F
            'If yMapWrap(X + 1) = 0 Then
            For Y As Int32 = X + 2 To Math.Min(lDestIdx, (X + 8))
                'Get the route's calculated cost to get here...
                For Z As Int32 = X To Y - 1
                    fCalculatedRouteCost += fCalcCosts(Z)
                Next Z

                Dim fTemp As Single = GetCostsBetweenPoints(ptReturn(X), ptReturn(Y), bNaval)

                If fTemp < fCalculatedRouteCost Then
                    Dim fNewEffRating As Single = fTemp / fCalculatedRouteCost
                    If fNewEffRating < fEffRating Then
                        'its faster to go directly there...
                        lNextIndex = Y
                        fEffRating = fNewEffRating
                    End If
                End If
            Next Y
            'Else
            '    If ptSuperOpt(lSuperOptUB).X < ptReturn(lNextIndex).X Then
            '        'wrapping from east to west, reduce ptReturn(lNextIndex).X by map width
            '        ptReturn(lNextIndex).X -= TerrainClass.Width
            '    Else
            '        If ptReturn(lNextIndex).X = 0 Then ptReturn(lNextIndex).X = 241 Else ptReturn(lNextIndex).X += TerrainClass.Width
            '    End If
            'End If


            lSuperOptUB += 1
            ptSuperOpt(lSuperOptUB) = ptReturn(lNextIndex)

            X = lNextIndex - 1
        Next X

        If ptSuperOpt(lSuperOptUB) <> ptReturn(lDestIdx) Then
            lSuperOptUB += 1
            ptSuperOpt(lSuperOptUB) = ptReturn(lDestIdx)
            'If yMapWrap(lDestIdx) <> 0 Then
            '    If ptSuperOpt(lSuperOptUB - 1).X < ptSuperOpt(lSuperOptUB).X Then
            '        'wrapping from east to west, reduce ptReturn(lNextIndex).X by map width
            '        ptSuperOpt(lSuperOptUB).X -= (TerrainClass.Width - 1)
            '    Else
            '        ptSuperOpt(lSuperOptUB).X += (TerrainClass.Width - 1)
            '    End If
            'End If
        End If

        ReDim Preserve ptSuperOpt(lSuperOptUB)

        Return ptSuperOpt
    End Function

    Private Function GetCostsBetweenBasePoints(ByVal pt1 As Point, ByVal pt2 As Point, ByVal bMapWrap As Boolean, ByVal bNaval As Boolean) As Single
        Dim fVecX As Single = pt2.X - pt1.X
        Dim fVecY As Single = pt2.Y - pt1.Y
        Dim fValue As Single

        Dim fLocX As Single = pt1.X
        Dim fLocY As Single = pt1.Y

        If bMapWrap = True Then
            If pt2.X > pt1.X Then
                fVecX = (pt2.X - TerrainClass.Width) - pt1.X
            Else : fVecX = pt2.X - (pt1.X - TerrainClass.Width)
            End If
        End If

        fValue = Math.Abs(fVecX) + Math.Abs(fVecY)
        fVecX = fVecX / fValue
        fVecY = fVecY / fValue

        If Math.Abs(fVecX) <> 1 AndAlso Math.Abs(fVecY) <> 1 Then
            If Math.Abs(fVecX) > Math.Abs(fVecY) Then
                fVecY /= Math.Abs(fVecX)
                If fVecX < 0 Then fVecX = -1 Else fVecX = 1
            Else
                fVecX /= Math.Abs(fVecY)
                If fVecY < 0 Then fVecY = -1 Else fVecY = 1
            End If
        End If

        'Now, walk our points
        Dim lX As Int32 = CInt(Math.Floor(fLocX))
        Dim lY As Int32 = CInt(Math.Floor(fLocY))
        Dim lIdx As Int32
        Dim lLastHt As Int32

        Dim lPrevX As Int32
        Dim lPrevY As Int32

        lIdx = lY * TerrainClass.Width + lX
        lLastHt = moHm(lIdx)

        fValue = 0.0F

        While lX <> pt2.X OrElse lY <> pt2.Y
            'If bMapWrap = False Then
            If lX < pt2.X AndAlso pt1.X > pt2.X Then Exit While
            If lX > pt2.X AndAlso pt1.X < pt2.X Then Exit While
            If lY < pt2.Y AndAlso pt1.Y > pt2.Y Then Exit While
            If lY > pt2.Y AndAlso pt1.Y < pt2.Y Then Exit While
            'Else
            '    If lX = 0 AndAlso fVecX < 0 Then
            '        lX += TerrainClass.Width - 1
            '        fVecX *= -1
            '        bMapWrap = False
            '    ElseIf lX = TerrainClass.Width - 1 AndAlso fVecX > 0 Then
            '        lX -= (TerrainClass.Width - 1)
            '        fVecX *= -1
            '        bMapWrap = False
            '    End If
            'End If

            fLocX += fVecX
            fLocY += fVecY

            lX = CInt(Math.Floor(fLocX))
            lY = CInt(Math.Floor(fLocY))

            If lPrevX <> lX OrElse lPrevY <> lY Then

                lIdx = lY * TerrainClass.Width + lX
                Dim lNewHt As Int32 = moHm(lIdx)

                If bNaval = True Then
                    If lNewHt > mlWaterHeight Then Continue While
                ElseIf lNewHt < mlWaterHeight Then
                    Continue While
                End If
                Dim lDiff As Int32 = lNewHt - lLastHt
                If lDiff > 0 Then fValue += (lDiff * lDiff) * 2 '(lDiff * 15)

                lLastHt = lNewHt
            End If
        End While

        Return fValue
    End Function

    Private Function GetCostsBetweenPoints(ByVal pt1 As Point, ByVal pt2 As Point, ByVal bNaval As Boolean) As Single
        Dim fVecX As Single = pt2.X - pt1.X
        Dim fVecY As Single = pt2.Y - pt1.Y
        Dim fValue As Single

        Dim fLocX As Single = pt1.X
        Dim fLocY As Single = pt1.Y

        fValue = Math.Abs(fVecX) + Math.Abs(fVecY)
        fVecX = fVecX / fValue
        fVecY = fVecY / fValue

        If Math.Abs(fVecX) <> 1 AndAlso Math.Abs(fVecY) <> 1 Then
            If Math.Abs(fVecX) > Math.Abs(fVecY) Then
                fVecY /= Math.Abs(fVecX)
                If fVecX < 0 Then fVecX = -1 Else fVecX = 1
            Else
                fVecX /= Math.Abs(fVecY)
                If fVecY < 0 Then fVecY = -1 Else fVecY = 1
            End If
        End If

        'Now, walk our points
        Dim lX As Int32 = CInt(Math.Floor(fLocX))
        Dim lY As Int32 = CInt(Math.Floor(fLocY))
        Dim lIdx As Int32
        Dim lLastHt As Int32

        Dim lPrevX As Int32
        Dim lPrevY As Int32

        lIdx = lY * TerrainClass.Width + lX
        lLastHt = moHm(lIdx)

        fValue = 0.0F

        While lX <> pt2.X OrElse lY <> pt2.Y

            If lX < pt2.X AndAlso pt1.X > pt2.X Then Exit While
            If lX > pt2.X AndAlso pt1.X < pt2.X Then Exit While
            If lY < pt2.Y AndAlso pt1.Y > pt2.Y Then Exit While
            If lY > pt2.Y AndAlso pt1.Y < pt2.Y Then Exit While

            fLocX += fVecX
            fLocY += fVecY

            lX = CInt(Math.Floor(fLocX))
            lY = CInt(Math.Floor(fLocY))

            If lPrevX <> lX OrElse lPrevY <> lY Then

                lIdx = lY * TerrainClass.Width + lX
                Dim lNewHt As Int32 = moHm(lIdx)

                If bNaval = False Then
                    If lNewHt < mlWaterHeight Then
                        Return Single.MaxValue
                    Else
                        For lTmpY As Int32 = lY - 1 To lY + 1
                            For lTmpX As Int32 = lX - 1 To lX + 1
                                If lTmpX <> lX OrElse lTmpY <> lY Then
                                    If lTmpX > -1 AndAlso lTmpX < TerrainClass.Width Then
                                        If lTmpY > -1 AndAlso lTmpY < TerrainClass.Height Then
                                            Dim lTmpIdx As Int32 = lTmpY * TerrainClass.Height + lTmpX
                                            If lNewHt < mlWaterHeight Then
                                                If moHm(lTmpIdx) > mlWaterHeight Then
                                                    Return Single.MaxValue
                                                End If
                                            Else
                                                If moHm(lTmpIdx) < mlWaterHeight Then
                                                    Return Single.MaxValue
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next lTmpX
                        Next lTmpY
                    End If
                Else
                    If lNewHt > mlWaterHeight Then
                        Return Single.MaxValue
                    Else
                        For lTmpY As Int32 = lY - 1 To lY + 1
                            For lTmpX As Int32 = lX - 1 To lX + 1
                                If lTmpX <> lX OrElse lTmpY <> lY Then
                                    If lTmpX > -1 AndAlso lTmpX < TerrainClass.Width Then
                                        If lTmpY > -1 AndAlso lTmpY < TerrainClass.Height Then
                                            Dim lTmpIdx As Int32 = lTmpY * TerrainClass.Height + lTmpX
                                            If lNewHt > mlWaterHeight Then
                                                If moHm(lTmpIdx) < mlWaterHeight Then
                                                    Return Single.MaxValue
                                                End If
                                            Else
                                                If moHm(lTmpIdx) > mlWaterHeight Then
                                                    Return Single.MaxValue
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next lTmpX
                        Next lTmpY
                    End If
                End If

                Dim lDiff As Int32 = lNewHt - lLastHt
                If lDiff > 0 Then fValue += (lDiff * lDiff) * 2 '(lDiff * 15)

                lLastHt = lNewHt
            End If
        End While

        Return fValue
    End Function

    Private Function GetClosestTileType(ByVal lStartX As Int32, ByVal lStartY As Int32, ByVal lEndX As Int32, ByVal lEndY As Int32, ByVal yTileType As Byte) As Point
        'This function will look around lEndX, lEndY moving towards StartX, StartY for a tile that is of yTileType and return that tile
        'We will be moving from END to START
        Dim fVecX As Single = lStartX - lEndX
        Dim fVecY As Single = lStartY - lEndY
        Dim fValue As Single

        Dim fLocX As Single = lEndX
        Dim fLocY As Single = lEndY

        fValue = Math.Abs(fVecX) + Math.Abs(fVecY)
        fVecX = fVecX / fValue
        fVecY = fVecY / fValue

        If Math.Abs(fVecX) <> 1 AndAlso Math.Abs(fVecY) <> 1 Then
            If Math.Abs(fVecX) > Math.Abs(fVecY) Then
                fVecY /= Math.Abs(fVecX)
                If fVecX < 0 Then fVecX = -1 Else fVecX = 1
            Else
                fVecX /= Math.Abs(fVecY)
                If fVecY < 0 Then fVecY = -1 Else fVecY = 1
            End If
        End If

        'Now, walk our points
        Dim lX As Int32 = CInt(Math.Floor(fLocX))
        Dim lY As Int32 = CInt(Math.Floor(fLocY))

        Dim lPrevX As Int32
        Dim lPrevY As Int32

        fValue = 0.0F

        While lX <> lStartX OrElse lY <> lStartY

            If lX < lStartX AndAlso lEndX > lStartX Then Exit While
            If lX > lStartX AndAlso lEndX < lStartX Then Exit While
            If lY < lStartY AndAlso lEndY > lStartY Then Exit While
            If lY > lStartY AndAlso lEndY < lStartY Then Exit While

            fLocX += fVecX
            fLocY += fVecY

            lX = CInt(Math.Floor(fLocX))
            lY = CInt(Math.Floor(fLocY))

            If lPrevX <> lX OrElse lPrevY <> lY Then

                If myNodeUsed(lX, lY) = 0 Then InitNode(lX, lY)

                If moNodes(lX, lY).TileType = yTileType Then
                    Return New Point(lX, lY)
                End If

                lPrevX = lX
                lPrevY = lY
            End If
        End While

        Return New Point(Int32.MinValue, Int32.MinValue)
    End Function

	Public Function GetClosestTileTypeToLocation(ByVal lLocX As Int32, ByVal lLocY As Int32, ByVal yTileType As Byte) As Point
        If lLocX < 0 Then lLocX = 0 'lLocX += TerrainClass.Width
        If lLocX > TerrainClass.Width - 1 Then lLocX = TerrainClass.Width - 1 'lLocX -= TerrainClass.Width
		If lLocX < 0 OrElse lLocY < 0 Then Return Point.Empty
		If lLocX > TerrainClass.Width - 1 OrElse lLocY > TerrainClass.Height - 1 Then Return Point.Empty

		'Ok, first, is the point we currently stand not good?
		If myNodeUsed(lLocX, lLocY) = 0 Then InitNode(lLocX, lLocY)
		If moNodes(lLocX, lLocY).TileType = yTileType Then
			Return New Point(lLocX, lLocY)
		End If

		'Ok, current point is NOT good... so look for a replacement
		Dim lDistVal As Int32 = 1
		Dim bDone As Boolean = False
		While bDone = False
			'ok, check our values
			For X As Int32 = -lDistVal To lDistVal
				For Y As Int32 = -lDistVal To lDistVal
					If X = 0 And Y = 0 Then Continue For

					Dim lTmpX As Int32 = lLocX + X
					Dim lTmpY As Int32 = lLocY + Y

                    If lTmpX < 0 Then lTmpX = 0 ' lTmpX += TerrainClass.Width
                    If lTmpX > TerrainClass.Width - 1 Then lTmpX = TerrainClass.Width - 1 'lTmpX -= TerrainClass.Width
					If lTmpY < 0 OrElse lTmpY > TerrainClass.Height - 1 Then Continue For

					If myNodeUsed(lTmpX, lTmpY) = 0 Then InitNode(lTmpX, lTmpY)
					If moNodes(lTmpX, lTmpY).TileType = yTileType Then
						Return New Point(lTmpX, lTmpY)
					End If

				Next Y
			Next X

			lDistVal += 1
			If lDistVal > 30 Then bDone = True
		End While

		Return Point.Empty
	End Function
End Class