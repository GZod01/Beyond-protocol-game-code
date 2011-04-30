Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class SolarSystem
    Inherits Base_GUID


    Public Enum elSystemType As Int32
        SpawnSystem = 0
        HubSystem = 1
        RespawnSystem = 2
        HubHubSystem = 3
        UnlockedSystem = 4
        TutorialSystem = 255
    End Enum

    Public SystemName As String
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32

    Public StarType1Idx As Int32 = -1   'index in the StarType array
    Public StarType2Idx As Int32 = -1   'index in the StarType array
    Public StarType3Idx As Int32 = -1   'index in the StarType array

    Public SystemType As Byte = 0

    Public FleetJumpPointX As Int32
    Public FleetJumpPointZ As Int32

    Private mlCurrentPlanetIdx As Int32 = -1

    Public moPlanets() As Planet
    Public PlanetUB As Int32 = -1

    Public moWormholes() As Wormhole
    Public WormholeUB As Int32 = -1

    Public lVoterIDs() As Int32
    Public yVoterCnts() As Byte
    Public lVoterUB As Int32 = -1
    Public yPlanetCnt As Byte = 0
    Private mbRequestedGalMapViewDetails As Boolean = False
    Public Sub RequestGalMapViewDetails()
        If mbRequestedGalMapViewDetails = True Then Return
        mbRequestedGalMapViewDetails = True

        Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yMsg, 0)
        Me.GetGUIDAsString.CopyTo(yMsg, 2)
        goUILib.SendMsgToPrimary(yMsg)
    End Sub
    Public Sub HandleRequestGalMapViewDetails(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        lPos += 6   'for guid
        yPlanetCnt = yData(lPos) : lPos += 1

        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
        Dim lIDs(lUB) As Int32
        Dim yCnts(lUB) As Byte

        For X As Int32 = 0 To lUB
            lIDs(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            yCnts(X) = yData(lPos) : lPos += 1
        Next X

        'Ok, sort the list by counts
        Dim lSorted(lUB) As Int32
        Dim lSortedIdx As Int32 = -1
        For X As Int32 = 0 To lUB
            Dim lIdx As Int32 = -1
            For Y As Int32 = 0 To lSortedIdx
                If yCnts(lSorted(Y)) < yCnts(X) Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedIdx += 1
            If lIdx = -1 Then
                lSorted(lSortedIdx) = X
            Else
                For Y As Int32 = lSortedIdx To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                Next Y
                lSorted(lIdx) = X
            End If
        Next X

        'Now, the sorted list needs to be coerced into the final list
        Dim lFinal(lUB) As Int32
        Dim yFinal(lUB) As Byte
        For X As Int32 = 0 To lUB
            lFinal(X) = lIDs(lSorted(X))
            yFinal(X) = yCnts(lSorted(X))
        Next

        lVoterUB = -1
        lVoterIDs = lFinal
        yVoterCnts = yFinal
        lVoterUB = lUB
    End Sub

	Public Shared clrStarCollectiveDiffuse As System.Drawing.Color
	Public Shared clrStarCollectiveAmbient As System.Drawing.Color
	Public Shared clrStarCollectiveSpecular As System.Drawing.Color
	Private Shared mlCollectiveSystemID As Int32 = -1

    Public Sub AddWormhole(ByRef oWormhole As Wormhole)
        For X As Int32 = 0 To WormholeUB
            If moWormholes(X) Is Nothing = False AndAlso oWormhole.ObjectID = moWormholes(X).ObjectID Then Return
        Next X
        WormholeUB += 1
        ReDim Preserve moWormholes(WormholeUB)
        moWormholes(WormholeUB) = oWormhole
    End Sub

#Region " For Rendering Space... "
    Private Shared mbSkyboxReady As Boolean = False
    Private Shared moSkyboxVerts() As CustomVertex.PositionColored
    Private Shared mvbSkybox As VertexBuffer
    Private Shared moSkyboxMat As Material
    Private Shared mlFarPlane As Int32

    Public Shared Sub ResetSkyBox()
        mbSkyboxReady = False
        Try
            If moSkyboxVerts Is Nothing = False Then Erase moSkyboxVerts
            If mvbSkybox Is Nothing = False Then mvbSkybox.Dispose()
        Catch
        End Try
        mvbSkybox = Nothing
        moSkyboxVerts = Nothing
    End Sub

    'Removed Color Sphere, going to go the textured sphere route
    'Private Shared moColorSphere As Mesh
    'Private Shared mbColorSphereReady As Boolean = False
    Private Shared moCosmoSphere As Mesh
	Private Shared moCosmoTex As Texture 'Private moCosmoTex As Texture
	Private Shared mlCosmoID As Int32 = -1
	Private Structure uColVal
		Public R As Int32
		Public G As Int32
		Public B As Int32
    End Structure
    Public Shared Sub ResetCosmoSphere()
        If moCosmoTex Is Nothing = False Then moCosmoTex.Dispose()
        moCosmoTex = Nothing
        If moCosmoSphere Is Nothing = False Then moCosmoSphere.Dispose()
        moCosmoSphere = Nothing
        If moCosmoDetail Is Nothing = False Then moCosmoDetail.Dispose()
        moCosmoDetail = Nothing
        If moCosmoShader Is Nothing = False Then moCosmoShader.DisposeMe()
        moCosmoShader = Nothing
    End Sub

    'Public Shared Sub SaveCosmoSphere()
    '    SurfaceLoader.Save("C:\out.bmp", ImageFileFormat.Bmp, moCosmoTex.GetSurfaceLevel(0))
    'End Sub

    'These values are important if the system resides in a nebula
    Public InANebula As Boolean = False
    Public NebulaCloudColor As System.Drawing.Color = System.Drawing.Color.Purple       'purple by default
    Public FogStart As Int32 = 50000
    Public FogEnd As Int32 = 75000

    Private moCelestialVecs() As CustomVertex.PositionColored
    Public SquareCenterX As Int32
    Public SquareCenterZ As Int32
    Public Shared lSysMapView2CameraX As Int32 = 0
    Public Shared lSysMapView2CameraZ As Int32 = 0

    Public Shared moCosmoDetail As Texture
    Private Shared mfGammaValue As Single
    Public Shared moCosmoShader As CosmoShader
    Private Shared Sub CreateCosmoSphere(ByVal lSeed As Int32, ByRef oDevice As Device)
        If moCosmoTex Is Nothing = False Then moCosmoTex.Dispose()
        moCosmoTex = Nothing

        mlCosmoID = lSeed

        Dim lWidth As Int32 = 512 '128
        Dim lHeight As Int32 = 512 '128

        Dim oRandom As New Random(lSeed)
        'Rnd(-1)
        'Randomize(lSeed)

        'Create the line of the galaxy... along the top and bottom (fuzzy)
        Dim colMap(lWidth - 1, lHeight - 1) As uColVal ' System.Drawing.Color
        Dim lCnt As Int32 = CInt(oRandom.NextDouble() * 10) + 1
        For lPass As Int32 = 0 To 1
            If lPass = 0 Then
                lWidth = 64 : lHeight = 64
            ElseIf lPass = 1 Then
                lWidth = 128 : lHeight = 128
            Else
                lWidth = 256 : lHeight = 256
            End If

            'now, create spatial anomalies...
            For lItem As Int32 = 1 To CInt(Math.Ceiling(lCnt / 2.0F))
                Dim lItemX As Int32 = CInt(oRandom.NextDouble() * lWidth * 0.75F) + (lWidth \ 5)
                Dim lItemY As Int32 = CInt(oRandom.NextDouble() * (lHeight \ 2)) + (lWidth \ 4)

                Dim lBaseR As Int32 = 0 'CInt(GRnd() * 255)
                Dim lBaseG As Int32 = 0 'CInt(oRandom.NextDouble() * 255)
                Dim lBaseB As Int32 = 0 'CInt(oRandom.NextDouble() * 255)
                Dim lDist As Int32 = CInt(oRandom.NextDouble() * (lWidth \ 5)) + 20

                lBaseR = CInt(oRandom.NextDouble() * 255)
                Dim lMaxVal As Int32 = Math.Min(255, 512 - lBaseR)
                lBaseB = CInt(oRandom.NextDouble() * lMaxVal)
                lMaxVal = Math.Max(lBaseR, lBaseB)
                lBaseG = CInt(oRandom.NextDouble() * lMaxVal)

                If lBaseR + lBaseG + lBaseB > 384 Then
                    lBaseR \= 2
                    lBaseG \= 2
                    lBaseB \= 2
                ElseIf lBaseR + lBaseG + lBaseB < 200 Then
                    Dim vecTemp As Vector3 = New Vector3(lBaseR, lBaseG, lBaseB)
                    vecTemp.Normalize()
                    lBaseR = Math.Min(CInt(vecTemp.X * 128), 255)
                    lBaseG = Math.Min(CInt(vecTemp.Y * 128), 255)
                    lBaseB = Math.Min(CInt(vecTemp.Z * 128), 255)
                End If

                Dim oStack As New Stack()
                Dim yClosed(lWidth - 1, lHeight - 1) As Byte

                yClosed(lItemX, lItemY) = 255
                'oStack.Push(New Point(lItemX, lItemY))

                'Ok, let's push all 8 items onto the list
                For X As Int32 = -1 To 1
                    For Y As Int32 = -1 To 1
                        Dim lTempX As Int32 = lItemX + X
                        If lTempX > lWidth - 1 Then lTempX -= lWidth
                        If lTempX < 0 Then lTempX += lWidth
                        Dim lTempY As Int32 = lItemY + Y
                        If lTempY > lHeight - 1 Then lTempY = lHeight - 1
                        If lTempY < 0 Then lTempY = 0

                        Dim pt As Point
                        pt.X = lTempX
                        pt.Y = lTempY

                        yClosed(pt.X, pt.Y) = 255

                        oStack.Push(pt)
                    Next Y
                Next X

                'Dim lQtrR As Int32 = lBaseR \ 2
                'Dim lQtrG As Int32 = lBaseG \ 2
                'Dim lQtrB As Int32 = lBaseB \ 2
                While oStack.Count > 0
                    Dim curPt As Point = CType(oStack.Pop(), Point)

                    Dim lR As Int32
                    Dim lG As Int32
                    Dim lB As Int32
                    With colMap(curPt.X, curPt.Y)

                        Dim fMod As Single = Math.Abs(curPt.X - lItemX) + Math.Abs(curPt.Y - lItemY)
                        fMod /= (lDist * 2)

                        If fMod > 1 Then fMod = 1.0F
                        fMod = 1.0F - fMod
                        If fMod < 0.001F Then fMod = 0.0F

                        lR = CInt(oRandom.NextDouble() * lBaseR)
                        lG = CInt(oRandom.NextDouble() * lBaseG)
                        lB = CInt(oRandom.NextDouble() * lBaseB)
                        'lR = CInt(oRandom.NextDouble() * lQtrR) + (lQtrR) ' * 3)
                        'lG = CInt(oRandom.NextDouble() * lQtrG) + (lQtrG) ' * 3)
                        'lB = CInt(oRandom.NextDouble() * lQtrB) + (lQtrB) ' * 3)

                        If .R < lR Then
                            lR = CInt(Math.Min(255, Math.Max(.R, lR * fMod)))
                        End If
                        If .G < lG Then
                            lG = CInt(Math.Min(255, Math.Max(.G, lG * fMod)))
                        End If
                        If .B < lB Then
                            lB = CInt(Math.Min(255, Math.Max(.B, lB * fMod)))
                        End If

                        .R = lR : .G = lG : .B = lB
                    End With
                    'colMap(curPt.X, curPt.Y) = System.Drawing.Color.FromArgb(255, lR, lG, lB)

                    'Ok, check for the eight surrounding this one
                    For X As Int32 = -1 To 1
                        For Y As Int32 = -1 To 1
                            Dim lTempX As Int32 = curPt.X + X
                            If lTempX > lWidth - 1 Then lTempX -= lWidth
                            If lTempX < 0 Then lTempX += lWidth

                            Dim lTempY As Int32 = curPt.Y + Y
                            If lTempY > lHeight - 1 Then lTempY = lHeight - 1
                            If lTempY < 0 Then lTempY = 0

                            If yClosed(lTempX, lTempY) <> 255 Then
                                Dim fChance As Single = Math.Abs(curPt.X + X - lItemX) + Math.Abs(curPt.Y + Y - lItemY)
                                fChance /= (lWidth + lHeight)       '0 to 1... with farther being better
                                fChance = (1.0F - fChance)          '0 to 1 with farther being worse
                                fChance *= 26.0F

                                With colMap(lTempX, lTempY)
                                    If oRandom.NextDouble() * 100 < fChance Then
                                        yClosed(lTempX, lTempY) = 255
                                        oStack.Push(New Point(lTempX, lTempY))
                                    Else
                                        Dim lNeighborCnt As Int32 = 0
                                        For lSubX As Int32 = -1 To 1
                                            For lSubY As Int32 = -1 To 1
                                                Dim lTempSubX As Int32 = lTempX + lSubX
                                                Dim lTempSubY As Int32 = lTempY + lSubY
                                                If lTempSubX > lWidth - 1 Then lTempSubX -= lWidth
                                                If lTempSubX < 0 Then lTempSubX += lWidth
                                                If lTempSubY > lHeight - 1 Then lTempSubY = lHeight - 1
                                                If lTempSubY < 0 Then lTempSubY = 0

                                                If yClosed(lTempSubX, lTempSubY) = 255 Then lNeighborCnt += 1
                                            Next lSubY
                                        Next lSubX
                                        If lNeighborCnt > 5 Then
                                            yClosed(lTempX, lTempY) = 255
                                            oStack.Push(New Point(lTempX, lTempY))
                                        End If
                                    End If

                                    If yClosed(lTempX, lTempY) = 0 Then
                                        yClosed(lTempX, lTempY) = 128


                                        Dim fMod As Single = Math.Abs(lTempX - lItemX) + Math.Abs(lTempY - lItemY)
                                        fMod /= (lDist * 2)

                                        If fMod > 1 Then fMod = 1.0F
                                        fMod = 1.0F - fMod
                                        If fMod < 0.001F Then fMod = 0.0F

                                        lR = CInt(oRandom.NextDouble() * (lBaseR \ 2))
                                        lG = CInt(oRandom.NextDouble() * (lBaseG \ 2))
                                        lB = CInt(oRandom.NextDouble() * (lBaseB \ 2))

                                        If .R < lR Then
                                            lR = CInt(Math.Min(255, Math.Max(.R, lR * fMod)))
                                        End If
                                        If .G < lG Then
                                            lG = CInt(Math.Min(255, Math.Max(.G, lG * fMod)))
                                        End If
                                        If .B < lB Then
                                            lB = CInt(Math.Min(255, Math.Max(.B, lB * fMod)))
                                        End If
                                    End If

                                End With
                            End If

                        Next Y
                    Next X
                End While
            Next lItem

            'Now, check if we are the first pass
            If lPass = 0 Then 'OrElse lPass = 1 Then
                If colMap.GetUpperBound(0) + 1 <> lWidth Then
                    For X As Int32 = lWidth - 1 To 0 Step -1
                        For Y As Int32 = lHeight - 1 To 0 Step -1
                            colMap(X * 2, Y * 2) = colMap(X, Y)
                        Next Y
                    Next X
                End If
            End If
        Next lPass

        'Now, do some smooth
        For lLoop As Int32 = 0 To CInt(oRandom.NextDouble() * 2)
            For Y As Int32 = 1 To lHeight - 2
                For X As Int32 = 1 To lWidth - 2

                    Dim lLeftX As Int32 = X - 1
                    If lLeftX < 0 Then lLeftX += lWidth
                    Dim lRightX As Int32 = X + 1
                    If lRightX > lWidth - 1 Then lRightX -= lWidth
                    Dim lTopY As Int32 = Y + 1
                    If lTopY > lHeight - 1 Then lTopY = lHeight - 1
                    Dim lBottomY As Int32 = Y - 1
                    If lBottomY < 0 Then lBottomY = 0

                    Dim fCenterR As Single = colMap(X, Y).R / 4.0F
                    Dim fCenterG As Single = colMap(X, Y).G / 4.0F
                    Dim fCenterB As Single = colMap(X, Y).B / 4.0F

                    Dim fSidesR As Single = (CInt(colMap(lRightX, Y).R) + CInt(colMap(lLeftX, Y).R) + CInt(colMap(X, lTopY).R) + CInt(colMap(X, lBottomY).R)) / 8.0F
                    Dim fSidesG As Single = (CInt(colMap(lRightX, Y).G) + CInt(colMap(lLeftX, Y).G) + CInt(colMap(X, lTopY).G) + CInt(colMap(X, lBottomY).G)) / 8.0F
                    Dim fSidesB As Single = (CInt(colMap(lRightX, Y).B) + CInt(colMap(lLeftX, Y).B) + CInt(colMap(X, lTopY).B) + CInt(colMap(X, lBottomY).B)) / 8.0F

                    Dim fCornerR As Single = (CInt(colMap(lRightX, lTopY).R) + CInt(colMap(lRightX, lBottomY).R) + CInt(colMap(lLeftX, lTopY).R) + CInt(colMap(lLeftX, lBottomY).R)) / 16.0F
                    Dim fCornerG As Single = (CInt(colMap(lRightX, lTopY).G) + CInt(colMap(lRightX, lBottomY).G) + CInt(colMap(lLeftX, lTopY).G) + CInt(colMap(lLeftX, lBottomY).G)) / 16.0F
                    Dim fCornerB As Single = (CInt(colMap(lRightX, lTopY).B) + CInt(colMap(lRightX, lBottomY).B) + CInt(colMap(lLeftX, lTopY).B) + CInt(colMap(lLeftX, lBottomY).B)) / 16.0F

                    Dim lFinalR As Int32 = CInt((fCenterR + fSidesR + fCornerR)) ' * 0.9F)
                    Dim lFinalG As Int32 = CInt((fCenterG + fSidesG + fCornerG)) ' * 0.9F)
                    Dim lFinalB As Int32 = CInt((fCenterB + fSidesB + fCornerB)) ' * 0.9F)

                    'colMap(X, Y) = System.Drawing.Color.FromArgb(255, lFinalR, lFinalG, lFinalB)
                    With colMap(X, Y)
                        .R = lFinalR : .G = lFinalG : .B = lFinalB
                    End With
                Next X
            Next Y
        Next lLoop

        Dim lCenterLineWidth As Int32 = CInt(oRandom.NextDouble() * 16) + (lHeight \ 5)
        Dim lHalfHeight As Int32 = lHeight \ 2
        For X As Int32 = 0 To lWidth - 1
            For Y As Int32 = 0 To lHeight - 1

                If Math.Abs(Y - lHalfHeight) < lCenterLineWidth Then

                    Dim lR As Int32 = colMap(X, Y).R
                    Dim lG As Int32 = colMap(X, Y).G
                    Dim lB As Int32 = colMap(X, Y).B

                    Dim fTemp As Single = 1.0F - CSng(Math.Abs(Y - lHalfHeight) / lCenterLineWidth)
                    lR = CInt(Math.Max(lR, ((16 * fTemp) + ((oRandom.NextDouble() * 16) - 8))))
                    lG = CInt(Math.Max(lG, (32 * fTemp) + ((oRandom.NextDouble() * 16) - 8)))
                    lB = CInt(Math.Max(lB, (64 * fTemp) + ((oRandom.NextDouble() * 16) - 8)))

                    'colMap(X, Y) = System.Drawing.Color.FromArgb(255, Math.Max(0, lR), Math.Max(0, lG), Math.Max(0, lB))
                    With colMap(X, Y)
                        .R = Math.Max(0, lR)
                        .G = Math.Max(0, lG)
                        .B = Math.Max(0, lB)
                    End With
                End If
            Next Y
        Next X

        For Y As Int32 = 0 To lHeight - 1
            For X As Int32 = 0 To lWidth - 1

                Dim lLeftX As Int32 = X - 1
                If lLeftX < 0 Then lLeftX += lWidth
                Dim lRightX As Int32 = X + 1
                If lRightX > lWidth - 1 Then lRightX -= lWidth
                Dim lTopY As Int32 = Y + 1
                If lTopY > lHeight - 1 Then lTopY = lHeight - 1
                Dim lBottomY As Int32 = Y - 1
                If lBottomY < 0 Then lBottomY = 0

                Dim fCenterR As Single = colMap(X, Y).R / 4.0F
                Dim fCenterG As Single = colMap(X, Y).G / 4.0F
                Dim fCenterB As Single = colMap(X, Y).B / 4.0F

                Dim fSidesR As Single = (CInt(colMap(lRightX, Y).R) + CInt(colMap(lLeftX, Y).R) + CInt(colMap(X, lTopY).R) + CInt(colMap(X, lBottomY).R)) / 8.0F
                Dim fSidesG As Single = (CInt(colMap(lRightX, Y).G) + CInt(colMap(lLeftX, Y).G) + CInt(colMap(X, lTopY).G) + CInt(colMap(X, lBottomY).G)) / 8.0F
                Dim fSidesB As Single = (CInt(colMap(lRightX, Y).B) + CInt(colMap(lLeftX, Y).B) + CInt(colMap(X, lTopY).B) + CInt(colMap(X, lBottomY).B)) / 8.0F

                Dim fCornerR As Single = (CInt(colMap(lRightX, lTopY).R) + CInt(colMap(lRightX, lBottomY).R) + CInt(colMap(lLeftX, lTopY).R) + CInt(colMap(lLeftX, lBottomY).R)) / 16.0F
                Dim fCornerG As Single = (CInt(colMap(lRightX, lTopY).G) + CInt(colMap(lRightX, lBottomY).G) + CInt(colMap(lLeftX, lTopY).G) + CInt(colMap(lLeftX, lBottomY).G)) / 16.0F
                Dim fCornerB As Single = (CInt(colMap(lRightX, lTopY).B) + CInt(colMap(lRightX, lBottomY).B) + CInt(colMap(lLeftX, lTopY).B) + CInt(colMap(lLeftX, lBottomY).B)) / 16.0F

                Dim lFinalR As Int32 = CInt((fCenterR + fSidesR + fCornerR)) ' * 0.9F)
                Dim lFinalG As Int32 = CInt((fCenterG + fSidesG + fCornerG)) ' * 0.9F)
                Dim lFinalB As Int32 = CInt((fCenterB + fSidesB + fCornerB)) ' * 0.9F)

                'colMap(X, Y) = System.Drawing.Color.FromArgb(255, lFinalR, lFinalG, lFinalB)
                With colMap(X, Y)
                    .R = lFinalR : .G = lFinalG : .B = lFinalB
                End With
            Next X
        Next Y



        moCosmoTex = New Texture(oDevice, lWidth, lHeight, 0, Usage.None, Format.A8R8G8B8, Pool.Managed)
        Dim lPitch As Int32
        Dim oStream As Microsoft.DirectX.GraphicsStream = moCosmoTex.LockRectangle(0, System.Drawing.Rectangle.Empty, Microsoft.DirectX.Direct3D.LockFlags.Discard, lPitch)

        Dim yData() As Byte

        Dim lIdx As Int32
        Dim lPos As Int32
        Dim lVal As Int32
        Dim lDataSize As Int32 = ((4 * lHeight) * lWidth) - 1

        Dim LocX As Int32
        Dim LocY As Int32

        ReDim yData(lDataSize)

        lPos = 0
        oStream.Position = lPos
        lVal = oStream.Read(yData, 0, yData.Length)

        For LocY = 0 To lHeight - 1
            For LocX = 0 To lWidth - 1
                lIdx = (LocY * (lWidth * 4)) + (LocX * 4)

                yData(lIdx) = CByte(colMap(LocX, LocY).B)
                yData(lIdx + 1) = CByte(colMap(LocX, LocY).G)
                yData(lIdx + 2) = CByte(colMap(LocX, LocY).R)
                yData(lIdx + 3) = 255
            Next LocX
        Next LocY

        oStream.Position = lPos
        oStream.Write(yData, 0, yData.Length)
        moCosmoTex.UnlockRectangle(0)

        'Now, set our Gamma and detail texture
        Dim lDetailID As Int32 = oRandom.Next(1, 11)
        Dim fMax As Single = 14
        If lDetailID = 1 Then fMax = 4
        mfGammaValue = CSng(oRandom.NextDouble() * fMax) + 1

        If moCosmoDetail Is Nothing = False Then moCosmoDetail.Dispose()
        moCosmoDetail = Nothing
        moCosmoDetail = goResMgr.LoadScratchTexture("Nebula" & lDetailID & ".bmp", "cosmos.pak")

    End Sub

    Public Shared Sub DoRandomCosmosGenerator(ByVal oDevice As Device)
        Dim lCurr As Int32 = mlCosmoID
        'Randomize()
        Dim lRndVal As Int32 = CInt(Rnd() * 10000)
        CreateCosmoSphere(lRndVal, oDevice)
        mlCosmoID = lCurr
    End Sub
#End Region

    Public Sub New()
        mlFarPlane = muSettings.FarClippingPlane - 1
        If mbSkyboxReady = False Then CreateSkybox()
        'If mbColorSphereReady = False Then CreateColorSphere()

        'If moCosmoDetail Is Nothing OrElse moCosmoDetail.Disposed = True Then
        '    moCosmoDetail = goResMgr.GetTexture("Nebula2.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
        'End If
    End Sub

    'Private Sub CreateColorSphere()
    '    moDevice.IsUsingEventHandlers = False
    '    Dim oTemp As Mesh = Mesh.Sphere(moDevice, muSettings.FarClippingPlane, 32, 32)

    '    moColorSphere = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionColored.Format, moDevice)

    '    ' Get the original mesh's vertex buffer.
    '    Dim ranks(0) As Integer
    '    ranks(0) = oTemp.NumberVertices
    '    Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

    '    ' Set the vertex buffer
    '    Dim data As System.Array = moColorSphere.VertexBuffer.Lock(0, (New CustomVertex.PositionColored()).GetType(), LockFlags.None, ranks)

    '    Dim phi As Single
    '    Dim u As Single
    '    Dim v As Single
    '    Dim i As Integer
    '    For i = 0 To arr.Length - 1
    '        Dim pn As Direct3D.CustomVertex.PositionNormal = arr.GetValue(i)
    '        Dim pnt As Direct3D.CustomVertex.PositionColored = data.GetValue(i)
    '        pnt.X = pn.X
    '        pnt.Y = pn.Y
    '        pnt.Z = pn.Z

    '        u = CSng(Math.Asin(pn.Nx) / Math.PI + 0.5)
    '        v = CSng(Math.Asin(pn.Ny) / Math.PI + 0.5)

    '        'The next line draws a gradient from the blue (40) to Black from Up and Down to the Horizon (blue to black to blue)
    '        'pnt.Color = System.Drawing.Color.FromArgb(255, 0, 0, (Math.Abs(0.5 - v) * 80)).ToArgb

    '        'The next line draws a gradient from blue (30) to Black from the Horizon to Up and Down (black to blue to black)
    '        pnt.Color = System.Drawing.Color.FromArgb(255, 0, 0, ((0.5 - Math.Abs(0.5 - v)) * 16) ^ 2).ToArgb

    '        data.SetValue(pnt, i)
    '    Next i

    '    moColorSphere.VertexBuffer.Unlock()
    '    oTemp.VertexBuffer.Unlock()

    '    ' Set the index buffer. 
    '    ranks(0) = oTemp.NumberFaces * 3
    '    arr = oTemp.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
    '    moColorSphere.IndexBuffer.SetData(arr, 0, LockFlags.None)

    '    mbColorSphereReady = True
    '    moDevice.IsUsingEventHandlers = True

    '    oTemp = Nothing
    '    arr = Nothing
    '    data = Nothing
    'End Sub

    Private mf2RSquared As Single = -1
    Private mf1Over2RSquared As Single = -1
    Private mf2R As Single = -1
    Public Function GetStarBaseYMod(ByVal fX As Single, ByVal fZ As Single) As Int32
        Try
            If mf2RSquared = -1 Then
                mf2R = 0
                If Me.StarType1Idx > -1 Then
                    mf2R += goStarTypes(StarType1Idx).StarRadius
                End If
                If StarType2Idx > -1 Then
                    mf2R += goStarTypes(StarType2Idx).StarRadius
                End If
                If StarType3Idx > -1 Then
                    mf2R += goStarTypes(StarType3Idx).StarRadius
                End If
                mf2R *= 2
                mf2RSquared = mf2R
                mf2RSquared *= mf2RSquared
                mf1Over2RSquared = 1.0F / mf2RSquared
            End If

            Dim fFastDist As Single = (fX * fX) + (fZ * fZ)
            Dim fRatio As Single = Math.Max(0, 1.0F - (fFastDist * mf1Over2RSquared))
            Return CInt(fRatio * mf2R)
        Catch
        End Try
        Return 0
    End Function

    Public Property CurrentPlanetIdx() As Int32
        Get
            Return mlCurrentPlanetIdx
        End Get
        Set(ByVal Value As Int32)
            If Value <> mlCurrentPlanetIdx Then
                If mlCurrentPlanetIdx > -1 Then
                    moPlanets(mlCurrentPlanetIdx).CleanResources(False)
                End If
            End If
            mlCurrentPlanetIdx = Value
        End Set
    End Property

    Private Sub CalculateCollectiveColors()
        Dim vecLight1Diff As Vector3
        Dim vecLight1Amb As Vector3
        Dim vecLight1Spec As Vector3
        Dim lIllum1Power As Int32 = 0
        Dim fIllum1Val As Single

        Dim vecLight2Diff As Vector3
        Dim vecLight2Amb As Vector3
        Dim vecLight2Spec As Vector3
        Dim lIllum2Power As Int32 = 0
        Dim fIllum2Val As Single

        Dim vecLight3Diff As Vector3
        Dim vecLight3Amb As Vector3
        Dim vecLight3Spec As Vector3
        Dim lIllum3Power As Int32 = 0
        Dim fIllum3Val As Single

        If StarType1Idx <> -1 Then
            Dim clrTemp As System.Drawing.Color = System.Drawing.Color.FromArgb(goStarTypes(StarType1Idx).LightDiffuse)
            vecLight1Diff = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)
            clrTemp = System.Drawing.Color.FromArgb(goStarTypes(StarType1Idx).LightAmbient)
            vecLight1Amb = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)
            clrTemp = System.Drawing.Color.FromArgb(goStarTypes(StarType1Idx).LightSpecular)
            vecLight1Spec = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)

            lIllum1Power = 5 - goStarTypes(StarType1Idx).lStarMapRectIdx
        End If
        If StarType2Idx <> -1 Then
            Dim clrTemp As System.Drawing.Color = System.Drawing.Color.FromArgb(goStarTypes(StarType2Idx).LightDiffuse)
            vecLight2Diff = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)
            clrTemp = System.Drawing.Color.FromArgb(goStarTypes(StarType2Idx).LightAmbient)
            vecLight2Amb = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)
            clrTemp = System.Drawing.Color.FromArgb(goStarTypes(StarType2Idx).LightSpecular)
            vecLight2Spec = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)

            lIllum2Power = 5 - goStarTypes(StarType2Idx).lStarMapRectIdx
        End If
        If StarType3Idx <> -1 Then
            Dim clrTemp As System.Drawing.Color = System.Drawing.Color.FromArgb(goStarTypes(StarType3Idx).LightDiffuse)
            vecLight3Diff = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)
            clrTemp = System.Drawing.Color.FromArgb(goStarTypes(StarType3Idx).LightAmbient)
            vecLight3Amb = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)
            clrTemp = System.Drawing.Color.FromArgb(goStarTypes(StarType3Idx).LightSpecular)
            vecLight3Spec = New Vector3(clrTemp.R, clrTemp.G, clrTemp.B)

            lIllum3Power = 5 - goStarTypes(StarType3Idx).lStarMapRectIdx
        End If

        'Now, for our illum vals
        Dim lSumIllumPower As Int32 = lIllum1Power + lIllum2Power + lIllum3Power
        If StarType1Idx <> -1 Then fIllum1Val = CSng(lIllum1Power / lSumIllumPower) Else fIllum1Val = Single.MinValue
        If StarType2Idx <> -1 Then fIllum2Val = CSng(lIllum2Power / lSumIllumPower) Else fIllum2Val = Single.MinValue
        If StarType3Idx <> -1 Then fIllum3Val = CSng(lIllum3Power / lSumIllumPower) Else fIllum3Val = Single.MinValue

        'Now, final light results...
        clrStarCollectiveSpecular = DoColorCombine(vecLight1Spec, fIllum1Val, vecLight2Spec, fIllum2Val, vecLight3Spec, fIllum3Val)
        clrStarCollectiveAmbient = DoColorCombine(vecLight1Amb, fIllum1Val, vecLight2Amb, fIllum2Val, vecLight3Amb, fIllum3Val)
        clrStarCollectiveDiffuse = DoColorCombine(vecLight1Diff, fIllum1Val, vecLight2Diff, fIllum2Val, vecLight3Diff, fIllum3Val)

        clrStarCollectiveSpecular = clrStarCollectiveDiffuse
    End Sub

    Private Function DoColorCombine(ByVal vec1 As Vector3, ByVal fIllumVal1 As Single, ByVal vec2 As Vector3, ByVal fIllumVal2 As Single, ByVal vec3 As Vector3, ByVal fIllumVal3 As Single) As System.Drawing.Color
        Dim fR As Single = 0
        Dim fG As Single = 0
        Dim fB As Single = 0

        If fIllumVal1 <> Single.MinValue Then
            fR += vec1.X * fIllumVal1 : fG += vec1.Y * fIllumVal1 : fB += vec1.Z * fIllumVal1
        End If
        If fIllumVal2 <> Single.MinValue Then
            fR += vec2.X * fIllumVal2 : fG += vec2.Y * fIllumVal2 : fB += vec2.Z * fIllumVal2
        End If
        If fIllumVal3 <> Single.MinValue Then
            fR += vec3.X * fIllumVal3 : fG += vec3.Y * fIllumVal3 : fB += vec3.Z * fIllumVal3
        End If

        fR /= 255.0F
        fG /= 255.0F
        fB /= 255.0F

        Dim vecTemp As Vector3 = New Vector3(fR, fG, fB)
        vecTemp.Normalize()

        fR = vecTemp.X * 255 : fG = vecTemp.Y * 255 : fB = vecTemp.Z * 255

        Dim fTemp As Single = 255.0F / Math.Max(Math.Max(Math.Max(fR, fG), fB), 1)
        fR *= fTemp
        fG *= fTemp
        fB *= fTemp

        'now, return our final color
        Return System.Drawing.Color.FromArgb(255, Math.Min(Math.Max(CInt(fR), 0), 255), Math.Min(Math.Max(CInt(fG), 0), 255), Math.Min(Math.Max(CInt(fB), 0), 255))
    End Function

    Public Sub Render(ByVal lEnvirType As Int32)
        Dim X As Int32
        Dim matWorld As Matrix = Matrix.Identity
        Dim matTemp As Matrix = Matrix.Identity
        Dim moDevice As Device = GFXEngine.moDevice

        If muSettings.RenderCosmos = True AndAlso (moCosmoTex Is Nothing OrElse moCosmoTex.Disposed = True OrElse mlCosmoID <> Me.ObjectID) Then
            CreateCosmoSphere(Me.ObjectID, moDevice)
        End If
        If mlCollectiveSystemID <> Me.ObjectID Then
            CalculateCollectiveColors()
            mlCollectiveSystemID = Me.ObjectID
        End If

        If lEnvirType = CurrentView.eSystemMapView1 Then
            Dim lIdx As Int32
            Dim lColor As Int32
            Dim oTmpMat As Material

            If (lEnvirType = CurrentView.eSystemMapView1 AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso NewTutorialManager.TutorialOn = False) Then
                If goUILib Is Nothing = False Then
                    Dim oFrm As UIWindow = goUILib.GetWindow("frmStrategicMapControl")
                    If oFrm Is Nothing Then
                        oFrm = New frmStrategicMapControl(goUILib)
                        oFrm.Visible = True
                    ElseIf oFrm.Visible = False Then
                        oFrm.Visible = True
                    End If
                    oFrm = Nothing
                End If
            End If

            moDevice.Clear(ClearFlags.ZBuffer Or ClearFlags.Target, System.Drawing.Color.Black, 1.0F, 0)

            moDevice.RenderState.FogEnable = False

            If muSettings.RenderCosmos = True Then
                'Ok, first, render our color sphere
                matTemp = Matrix.Identity
                matWorld = Matrix.Identity
                matTemp.RotateX(1.57079633)
                matWorld.Multiply(matTemp)
                matTemp = Matrix.Identity
                matTemp.Translate(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ)
                matWorld.Multiply(matTemp)
                moDevice.Transform.World = matWorld


                If moCosmoShader Is Nothing Then moCosmoShader = New CosmoShader()
                moCosmoShader.BeginRender(moCosmoTex, moCosmoDetail, mfGammaValue)

                moDevice.RenderState.Wrap0 = WrapCoordinates.Zero
                moDevice.RenderState.CullMode = Cull.None
                moDevice.RenderState.Lighting = False
                moDevice.RenderState.ZBufferWriteEnable = False


                If moCosmoSphere Is Nothing Then moCosmoSphere = goResMgr.CreateTexturedSphere(250000, 32, 32, 0, False)
                'moDevice.SetTexture(0, moCosmoTex)
                moCosmoSphere.DrawSubset(0)
                'moDevice.SetTexture(0, Nothing)

                moCosmoShader.EndRender()

                moDevice.RenderState.ZBufferWriteEnable = True
                moDevice.RenderState.Lighting = True
                moDevice.RenderState.CullMode = Cull.CounterClockwise
                moDevice.RenderState.Wrap0 = 0 ' WrapCoordinates.Two
            End If

            If moCelestialVecs Is Nothing Then
                X = -1 ' X = PlanetUB
                If moPlanets Is Nothing = False Then X = Math.Min(PlanetUB, moPlanets.GetUpperBound(0))
                If StarType1Idx <> -1 Then X += 1
                If StarType2Idx <> -1 Then X += 1
                If StarType3Idx <> -1 Then X += 1
                ReDim moCelestialVecs(X)

                'Now, create our celestial vec list
                lIdx = -1
                If StarType1Idx <> -1 Then
                    lIdx += 1
                    moCelestialVecs(lIdx) = New CustomVertex.PositionColored(0, 0, 0, goStarTypes(StarType1Idx).LightDiffuse)
                End If
                If StarType2Idx <> -1 Then
                    lIdx += 1
                    moCelestialVecs(lIdx) = New CustomVertex.PositionColored(CSng((goStarTypes(StarType1Idx).StarRadius + goStarTypes(StarType2Idx).StarRadius + 1000) / 10000), CSng(((goStarTypes(StarType2Idx).StarRadius / 2) - goStarTypes(StarType1Idx).StarRadius) / 10000), 0, goStarTypes(StarType2Idx).LightDiffuse)
                End If
                If StarType3Idx <> -1 Then
                    lIdx += 1
                    X = goStarTypes(StarType1Idx).StarRadius + goStarTypes(StarType2Idx).StarRadius + 1000
                    moCelestialVecs(lIdx) = New CustomVertex.PositionColored(CSng((-X) / 10000), 0, CSng(X / 10000), goStarTypes(StarType2Idx).LightDiffuse)
                End If

                lIdx += 1
                For X = 0 To PlanetUB
                    Select Case moPlanets(X).MapTypeID
                        Case PlanetType.eAcidic
                            lColor = System.Drawing.Color.Green.ToArgb
                        Case PlanetType.eAdaptable
                            lColor = System.Drawing.Color.FromArgb(255, 255, 192, 192).ToArgb
                        Case PlanetType.eBarren
                            lColor = System.Drawing.Color.FromArgb(255, 128, 128, 128).ToArgb
                        Case PlanetType.eDesert
                            lColor = System.Drawing.Color.Tan.ToArgb
                        Case PlanetType.eGasGiant
                            lColor = System.Drawing.Color.Orange.ToArgb
                        Case PlanetType.eGeoPlastic
                            lColor = System.Drawing.Color.Red.ToArgb
                        Case PlanetType.eTerran
                            lColor = System.Drawing.Color.Blue.ToArgb
                        Case PlanetType.eTundra
                            lColor = System.Drawing.Color.White.ToArgb
                        Case PlanetType.eWaterWorld
                            lColor = System.Drawing.Color.Teal.ToArgb
                    End Select
                    Dim lCelIdx As Int32 = X + lIdx
                    If lCelIdx <= moCelestialVecs.GetUpperBound(0) Then moCelestialVecs(lCelIdx) = New CustomVertex.PositionColored(CSng(moPlanets(X).LocX / 10000.0F), 0, CSng(moPlanets(X).LocZ / 10000.0F), lColor)
                Next X

            End If

            Try
                moDevice.Transform.World = Matrix.Identity
                Using oFont As New Font(moDevice, New System.Drawing.Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0))
                    For X = 0 To PlanetUB
                        Dim fTmpX As Single = moPlanets(X).LocX / 10000.0F
                        Dim fTmpZ As Single = moPlanets(X).LocZ / 10000.0F

                        Dim vec As Vector3 = New Vector3(fTmpX, 0, fTmpZ)
                        vec.Project(moDevice.Viewport, moDevice.Transform.Projection, moDevice.Transform.View, moDevice.Transform.World)
                        oFont.DrawText(Nothing, moPlanets(X).PlanetName, CInt(vec.X), CInt(vec.Y), System.Drawing.Color.FromArgb(128, 255, 255, 255))
                    Next X
                End Using
            Catch
            End Try

            'Ok, system map view is... special
            moDevice.Transform.World = Matrix.Identity
            moDevice.RenderState.ZBufferWriteEnable = False
            moDevice.RenderState.Lighting = False

            oTmpMat.Ambient = System.Drawing.Color.FromArgb(255, 64, 64, 64)
            oTmpMat.Diffuse = System.Drawing.Color.White
            oTmpMat.Emissive = System.Drawing.Color.Black
            oTmpMat.Specular = oTmpMat.Emissive
            moDevice.Material = oTmpMat

            moDevice.SetTexture(0, Nothing)

            moDevice.SetTexture(0, goResMgr.GetTexture("StratMap.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "Prax.pak"))

            Dim oVertsA(3) As CustomVertex.PositionColoredTextured
            oVertsA(0) = New CustomVertex.PositionColoredTextured(-650, -1, -650, System.Drawing.Color.FromArgb(64, 255, 255, 255).ToArgb, 0, 1)
            oVertsA(1) = New CustomVertex.PositionColoredTextured(-650, -1, 650, System.Drawing.Color.FromArgb(64, 255, 255, 255).ToArgb, 0, 0)
            oVertsA(2) = New CustomVertex.PositionColoredTextured(650, -1, -650, System.Drawing.Color.FromArgb(64, 255, 255, 255).ToArgb, 1, 1)
            oVertsA(3) = New CustomVertex.PositionColoredTextured(650, -1, 650, System.Drawing.Color.FromArgb(64, 255, 255, 255).ToArgb, 1, 0)
            moDevice.VertexFormat = CustomVertex.PositionColoredTextured.Format
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, oVertsA)
            moDevice.SetTexture(0, Nothing)

            'Now, get our zoom square where the mouse is located...
            moDevice.VertexFormat = CustomVertex.PositionColored.Format
            Dim oVerts(3) As CustomVertex.PositionColored
            oVerts(0) = New CustomVertex.PositionColored(SquareCenterX - 15, -0.5, SquareCenterZ - 15, System.Drawing.Color.FromArgb(128, 0, 128, 255).ToArgb)
            oVerts(1) = New CustomVertex.PositionColored(SquareCenterX - 15, -0.5, SquareCenterZ + 15, System.Drawing.Color.FromArgb(128, 0, 128, 255).ToArgb)
            oVerts(2) = New CustomVertex.PositionColored(SquareCenterX + 15, -0.5, SquareCenterZ - 15, System.Drawing.Color.FromArgb(128, 0, 128, 255).ToArgb)
            oVerts(3) = New CustomVertex.PositionColored(SquareCenterX + 15, -0.5, SquareCenterZ + 15, System.Drawing.Color.FromArgb(128, 0, 128, 255).ToArgb)
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, oVerts)

            'Now, get our zoom square where the current Sys Map View 2 is located...
            oVerts(0) = New CustomVertex.PositionColored(lSysMapView2CameraX - 15, -0.5, lSysMapView2CameraZ - 15, System.Drawing.Color.FromArgb(128, 128, 128, 128).ToArgb)
            oVerts(1) = New CustomVertex.PositionColored(lSysMapView2CameraX - 15, -0.5, lSysMapView2CameraZ + 15, System.Drawing.Color.FromArgb(128, 128, 128, 128).ToArgb)
            oVerts(2) = New CustomVertex.PositionColored(lSysMapView2CameraX + 15, -0.5, lSysMapView2CameraZ - 15, System.Drawing.Color.FromArgb(128, 128, 128, 128).ToArgb)
            oVerts(3) = New CustomVertex.PositionColored(lSysMapView2CameraX + 15, -0.5, lSysMapView2CameraZ + 15, System.Drawing.Color.FromArgb(128, 128, 128, 128).ToArgb)
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, oVerts)

            Dim lPrevSrcBlend As Blend
            Dim lPrevDestBlend As Blend

            With moDevice
                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

                lPrevSrcBlend = .RenderState.SourceBlend
                lPrevDestBlend = .RenderState.DestinationBlend

                .RenderState.SourceBlend = Blend.SourceAlpha
                .RenderState.DestinationBlend = Blend.One
                .RenderState.AlphaBlendEnable = True

                'Ok, if our device was created with mixed...
                If .CreationParameters.Behavior.MixedVertexProcessing = True Then
                    'Set us up for software vertex processing as point sprites always work in software
                    .SoftwareVertexProcessing = True
                End If

                .RenderState.PointSpriteEnable = True
                If .DeviceCaps.VertexFormatCaps.SupportsPointSize = True Then
                    .RenderState.PointScaleEnable = True
                    If .DeviceCaps.MaxPointSize > 16 Then
                        .RenderState.PointSize = 16
                    Else : .RenderState.PointSize = .DeviceCaps.MaxPointSize
                    End If
                    .RenderState.PointScaleA = 0
                    .RenderState.PointScaleB = 0
                    .RenderState.PointScaleC = 0.8F
                End If
            End With


            moDevice.SetTexture(0, goResMgr.GetTexture("Particle.dds", GFXResourceManager.eGetTextureType.NoSpecifics, ""))
            moDevice.DrawUserPrimitives(PrimitiveType.PointList, moCelestialVecs.Length, moCelestialVecs)

            With moDevice
                .SetTexture(0, Nothing)

                'Then, reset our device...  
                .RenderState.PointSpriteEnable = False
                .RenderState.PointScaleEnable = False

                'Ok, if our device was created with mixed...
                If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                    'Set us up for software vertex processing as point sprites always work in software
                    moDevice.SoftwareVertexProcessing = False
                End If

                .RenderState.SourceBlend = lPrevSrcBlend
                .RenderState.DestinationBlend = lPrevDestBlend

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With

            moDevice.RenderState.ZBufferWriteEnable = True
            moDevice.RenderState.Lighting = True
        ElseIf lEnvirType = CurrentView.eSystemView OrElse lEnvirType = CurrentView.eSystemMapView2 Then

            moDevice.RenderState.ZBufferEnable = True
            'set up our lights
            If StarType1Idx <> -1 Then
                With moDevice.Lights(ge_Light_Idx.eStar1Light)
                    .FromLight(goStarTypes(StarType1Idx).StarLight)
                    .Type = LightType.Point
                    .Position = Nothing
                    .Position = New Vector3(0, 0, 0)        '0,0,0 is 0,0,0 regardless of view
                    .Enabled = True
                    .Update()
                End With
            Else
                With moDevice.Lights(ge_Light_Idx.eStar1Light)
                    .Enabled = False
                    .Update()
                End With
            End If
            If StarType2Idx <> -1 Then
                With moDevice.Lights(ge_Light_Idx.eStar2Light)
                    .FromLight(goStarTypes(StarType2Idx).StarLight)
                    .Type = LightType.Point
                    .Position = Nothing
                    If lEnvirType = CurrentView.eSystemMapView2 Then
                        .Position = New Vector3(CSng((goStarTypes(StarType1Idx).StarRadius + goStarTypes(StarType2Idx).StarRadius + 1000) / 30), CSng(((goStarTypes(StarType2Idx).StarRadius / 2) - goStarTypes(StarType1Idx).StarRadius) / 30), 0)
                    Else
                        .Position = New Vector3(goStarTypes(StarType1Idx).StarRadius + goStarTypes(StarType2Idx).StarRadius + 1000, CSng(goStarTypes(StarType2Idx).StarRadius / 2) - goStarTypes(StarType1Idx).StarRadius, 0)
                    End If
                    .Enabled = True
                    .Update()
                End With
            Else
                With moDevice.Lights(ge_Light_Idx.eStar2Light)
                    .Enabled = False
                    .Update()
                End With
            End If
            If StarType3Idx <> -1 Then
                With moDevice.Lights(ge_Light_Idx.eStar3Light)
                    .FromLight(goStarTypes(StarType3Idx).StarLight)
                    .Type = LightType.Point
                    .Position = Nothing
                    If lEnvirType = CurrentView.eSystemMapView2 Then
                        X = goStarTypes(StarType1Idx).StarRadius + goStarTypes(StarType2Idx).StarRadius + 1000
                        .Position = New Vector3(CSng(-X / 30), CSng((goStarTypes(StarType1Idx).StarRadius) / 30), CSng(X / 30))
                    Else
                        X = goStarTypes(StarType1Idx).StarRadius + goStarTypes(StarType2Idx).StarRadius + 1000
                        .Position = New Vector3(-X, (goStarTypes(StarType1Idx).StarRadius), X)
                    End If
                    .Enabled = True
                    .Update()
                End With
            Else
                With moDevice.Lights(ge_Light_Idx.eStar3Light)
                    .Enabled = False
                    .Update()
                End With
            End If

            With moDevice

                If InANebula = True Then
                    .Clear(ClearFlags.Target Or ClearFlags.ZBuffer, NebulaCloudColor, 1.0F, 0)
                    .RenderState.FogColor = NebulaCloudColor
                    .RenderState.FogEnable = True
                    .RenderState.FogVertexMode = FogMode.Linear
                    .RenderState.FogStart = FogStart
                    .RenderState.FogEnd = FogEnd
                Else
                    .Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)
                    .RenderState.FogEnable = False
                End If

                'Render the skybox (stars, distant nebulae, etc...)
                If mbSkyboxReady = True Then '  (mbSkyboxReady And (Not InANebula)) Then 'AndAlso lEnvirType <> CurrentView.eSystemMapView2 Then
                    If InANebula = False Then
                        'Do this for everyone's sake

                        .SetTexture(0, Nothing)
                        If muSettings.RenderCosmos = True Then
                            'Ok, first, render our color sphere

                            'SurfaceLoader.Save("C:\CosmoTex.bmp", ImageFileFormat.Bmp, moCosmoTex.GetSurfaceLevel(0))
                            matTemp = Matrix.Identity
                            matWorld = Matrix.Identity
                            matTemp.RotateX(1.57079633)
                            matWorld.Multiply(matTemp)
                            matTemp = Matrix.Identity
                            matTemp.Translate(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ)
                            matWorld.Multiply(matTemp)
                            moDevice.Transform.World = matWorld

                            If moCosmoShader Is Nothing Then moCosmoShader = New CosmoShader()
                            moCosmoShader.BeginRender(moCosmoTex, moCosmoDetail, mfGammaValue)

                            'now, draw our spherem
                            moDevice.RenderState.Wrap0 = WrapCoordinates.Zero
                            moDevice.RenderState.Wrap1 = WrapCoordinates.Zero
                            moDevice.RenderState.CullMode = Cull.None
                            moDevice.RenderState.Lighting = False
                            moDevice.RenderState.ZBufferWriteEnable = False
                            'If moCosmoTex Is Nothing Then moCosmoTex = goResMgr.GetTexture("SpaceBox1.bmp", GFXResourceManager.eGetTextureType.NoSpecifics)
                            If moCosmoSphere Is Nothing Then moCosmoSphere = goResMgr.CreateTexturedSphere(250000, 32, 32, 0, False)
                            'moDevice.SetTexture(0, moCosmoTex)
                            moCosmoSphere.DrawSubset(0)
                            'moDevice.SetTexture(0, Nothing)
                            moDevice.RenderState.ZBufferWriteEnable = True
                            moDevice.RenderState.Lighting = True
                            moDevice.RenderState.CullMode = Cull.CounterClockwise
                            moDevice.RenderState.Wrap0 = 0 'WrapCoordinates.Two
                            moDevice.RenderState.Wrap1 = 0

                            moCosmoShader.EndRender()
                        End If

                        'DRAW OUR DISTANT NEBULAE
                        'BEGIN RENDER OF NEBULA
                        matWorld = Matrix.Identity
                        matTemp = Matrix.Identity
                        matTemp.Scale(2500, 2500, 2500)
                        matWorld.Multiply(matTemp)
                        matTemp = Matrix.Identity
                        matTemp.Translate(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ + 2000000)
                        matWorld.Multiply(matTemp)
                        matTemp = Nothing
                        moDevice.Transform.World = matWorld

                        'TODO: This is Nebula rendering, make it better it was causing performance issues
                        'Device.IsUsingEventHandlers = False
                        'Dim oSpr As Sprite = New Sprite(moDevice)
                        'Device.IsUsingEventHandlers = True
                        'moDevice.RenderState.ZBufferWriteEnable = False
                        'oSpr.SetWorldViewLH(moDevice.Transform.World, moDevice.Transform.View)
                        'oSpr.Begin(SpriteFlags.ObjectSpace)
                        'oSpr.Draw(moDistantNebula, System.Drawing.Rectangle.Empty, New Vector3(512, 512, 0), New Vector3(0, 0, 0), System.Drawing.Color.White)
                        'oSpr.End()
                        'moDevice.RenderState.ZBufferWriteEnable = True
                        'oSpr.Dispose()
                        'oSpr = Nothing

                        'Render star field... if we're not in system map view 2
                        If lEnvirType <> CurrentView.eSystemMapView2 Then
                            .Material = moSkyboxMat
                            'set up our renderstates
                            .RenderState.Lighting = False
                            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
                            .RenderState.SourceBlend = Blend.SourceAlpha
                            .RenderState.DestinationBlend = Blend.One
                            .RenderState.AlphaBlendEnable = True

                            'Ok, if our device was created with mixed...
                            If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                                'Set us up for software vertex processing as point sprites always work in software
                                moDevice.SoftwareVertexProcessing = True
                            End If

                            .RenderState.PointSpriteEnable = True
                            .RenderState.PointSize = 0.8
                            'render our points
                            Dim omatTemp As Matrix = Matrix.Identity
                            Dim omatTrans As Matrix = Matrix.Translation(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ)
                            omatTemp.Multiply(omatTrans)
                            .Transform.World = omatTemp
                            omatTemp = Nothing : omatTrans = Nothing

                            .SetStreamSource(0, mvbSkybox, 0)
                            .VertexFormat = CustomVertex.PositionColored.Format
                            .DrawPrimitives(PrimitiveType.PointList, 0, muSettings.StarfieldParticlesSpace)
                        End If

                        'now reset our renderstates
                        .RenderState.PointSpriteEnable = False

                        'Ok, if our device was created with mixed...
                        If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                            'Reset thesoftware vertex processing...
                            moDevice.SoftwareVertexProcessing = False
                        End If

                        .RenderState.Lighting = True

                        .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                        .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                        .RenderState.AlphaBlendEnable = True

                        .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                        .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                        .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                    End If
                Else
                    CreateSkybox()
                End If

                'set our texture wrapping coordinates
                .RenderState.Wrap0 = WrapCoordinates.Zero

                'OK, render our star(s)...
                'we could disable our Fog for the star to make it... SHINIER or whatever... especially if we
                '  put in a lens flare effect for it
                Dim lHtAdjust As Int32 = GetStarBaseYMod(goCamera.mlCameraX, goCamera.mlCameraZ)
                If StarType1Idx <> -1 Then goStarTypes(StarType1Idx).Render(lEnvirType, 1, 0, lHtAdjust) '1's radius is not needed
                If StarType2Idx <> -1 Then goStarTypes(StarType2Idx).Render(lEnvirType, 2, goStarTypes(StarType1Idx).StarRadius, lHtAdjust)
                If StarType3Idx <> -1 Then goStarTypes(StarType3Idx).Render(lEnvirType, 3, goStarTypes(StarType1Idx).StarRadius, lHtAdjust)

                'render our planets
                Try
                    For X = 0 To PlanetUB
                        moPlanets(X).Render(lEnvirType)
                    Next X
                Catch
                End Try

                'now reset texture wrapping coordinates
                .RenderState.Wrap0 = 0 ' WrapCoordinates.Two

                'Render our wormholes (if existant)
                If WormholeUB <> -1 Then
                    If goWormholeMgr Is Nothing Then goWormholeMgr = New WormholeManager()
                    If goWormholeMgr.lCurrentSystemID <> Me.ObjectID Then
                        goWormholeMgr.ClearAllWormholes()
                        Try
                            For X = 0 To WormholeUB
                                If moWormholes(X) Is Nothing = False Then
                                    goWormholeMgr.AddWormhole(moWormholes(X), Me.ObjectID)
                                End If
                            Next X
                        Catch
                        End Try
                        goWormholeMgr.lCurrentSystemID = Me.ObjectID
                    End If
                End If
            End With
        ElseIf lEnvirType = CurrentView.ePlanetMapView OrElse lEnvirType = CurrentView.ePlanetView Then
            If moPlanets Is Nothing = False AndAlso CurrentPlanetIdx <= moPlanets.GetUpperBound(0) AndAlso CurrentPlanetIdx > -1 Then moPlanets(CurrentPlanetIdx).Render(lEnvirType)
        End If
    End Sub

    'Public Sub Test()
    '	Dim matTemp As Matrix = Matrix.Identity
    '	Dim matWorld As Matrix = Matrix.Identity
    '	matTemp.RotateX(1.57079633)
    '	matWorld.Multiply(matTemp)
    '	matTemp = Matrix.Identity
    '	matTemp.Translate(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ)
    '	matWorld.Multiply(matTemp)
    '	moDevice.Transform.World = matWorld
    '	Dim oTmpMat As Material = moDevice.Material

    '	Dim oNewMat As Material
    '	With oNewMat
    '		.Ambient = System.Drawing.Color.FromArgb(255, 128, 128, 128)
    '		.Diffuse = .Ambient
    '		.Emissive = .Ambient
    '	End With
    '	moDevice.Material = oNewMat

    '	'now, draw our sphere
    '	moDevice.RenderState.Wrap0 = WrapCoordinates.Zero
    '	moDevice.RenderState.CullMode = Cull.None
    '	moDevice.RenderState.Lighting = False
    '	moDevice.RenderState.ZBufferEnable = False
    '	If moCosmoTex Is Nothing Then moCosmoTex = goResMgr.GetTexture("SpaceBox1.bmp", GFXResourceManager.eGetTextureType.NoSpecifics)
    '	If moCosmoSphere Is Nothing Then moCosmoSphere = goResMgr.CreateTexturedSphere(250000, 32, 32, 0, False)
    '	moDevice.SetTexture(0, moCosmoTex)
    '	moCosmoSphere.DrawSubset(0)
    '	moDevice.SetTexture(0, Nothing)
    '	moDevice.RenderState.ZBufferEnable = False
    '	moDevice.RenderState.Lighting = True
    '	moDevice.RenderState.CullMode = Cull.CounterClockwise
    '	moDevice.RenderState.Wrap0 = 0 'WrapCoordinates.Two

    '	moDevice.Material = oTmpMat
    'End Sub

    Public Function CheckSystemMapHover(ByVal lX As Int32, ByVal lZ As Int32) As Int32
        Dim lIdx As Int32

        If moCelestialVecs Is Nothing Then Return -1

        For lIdx = 0 To moCelestialVecs.Length - 1
            If moCelestialVecs(lIdx).X > lX - 3 AndAlso moCelestialVecs(lIdx).X < lX + 3 Then
                If moCelestialVecs(lIdx).Z > lZ - 3 AndAlso moCelestialVecs(lIdx).Z < lZ + 3 Then
                    Return lIdx
                End If
            End If
        Next lIdx

        Return -1

    End Function

    Public Sub AddPlanet(ByVal oPlanet As Planet)
        PlanetUB += 1
        ReDim Preserve moPlanets(PlanetUB)
        moPlanets(PlanetUB) = oPlanet
    End Sub

    Public Sub CleanResources(ByVal bComplete As Boolean)
        Dim X As Int32

        'Clean up every planet... because we are leaving the system...
        For X = 0 To PlanetUB
            If moPlanets(X) Is Nothing = False Then moPlanets(X).CleanResources(False)
        Next X

        'if bComplete, then we are leaving the game
        If bComplete = True Then
            If mvbSkybox Is Nothing = False Then
                mvbSkybox.Dispose()
                mvbSkybox = Nothing
                Erase moSkyboxVerts
            End If

            Erase moPlanets
            PlanetUB = -1
        End If

    End Sub

    Protected Overrides Sub Finalize()
        CleanResources(True)    'if we are finalizing a system then we are destroying it
        MyBase.Finalize()
    End Sub

    Private Sub CreateSkybox()
        Dim lIdx As Int32
        Dim fX As Single
        Dim fY As Single
        Dim fZ As Single
        Dim fD As Single

        Dim lColor As System.Drawing.Color
        Dim lCVal As Int32

        'now, generate our starfield...
        ReDim moSkyboxVerts(muSettings.StarfieldParticlesSpace - 1)
        For lIdx = 0 To muSettings.StarfieldParticlesSpace - 1
            'Now, we want to position our stars... anywhere along the sphere created by lFarPlane...
            fX = 0
            fZ = mlFarPlane - 1
            fY = 0

            fD = Int(Rnd() * 720) / 2
            RotatePoint(0, 0, fX, fZ, fD)
            fD = Int(Rnd() * 720) / 2
            RotatePoint(0, 0, fX, fY, fD)
            fD = Int(Rnd() * 720) / 2
            RotatePoint(0, 0, fZ, fY, fD)

            lCVal = CInt(Int(Rnd() * 192) + 63)
            lColor = System.Drawing.Color.FromArgb(255, lCVal, lCVal, lCVal)

            moSkyboxVerts(lIdx) = New CustomVertex.PositionColored(fX, fY, fZ, lColor.ToArgb)
        Next lIdx

        'Now, create our buffer
        mvbSkybox = New VertexBuffer(moSkyboxVerts(0).GetType, muSettings.StarfieldParticlesSpace, GFXEngine.moDevice, Usage.Points, CustomVertex.PositionColored.Format, Pool.Managed)
        mvbSkybox.SetData(moSkyboxVerts, 0, LockFlags.None)

        With moSkyboxMat
            .Ambient = System.Drawing.Color.White
            .Diffuse = .Ambient
            .Emissive = .Ambient
            .Specular = System.Drawing.Color.Black
        End With

        mbSkyboxReady = True
    End Sub

End Class
