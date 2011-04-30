Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class AOEExplosionMgr
    Private Shared moRandom As New Random
    Public Shared goMgr As New AOEExplosionMgr
    Private Shared moPlanet As Planet

    Private Class AOEExplosion
        Public vecGroundZero As Vector3
        Public lAOERange As Int32
        Public mlAOERangeSqrd As Int32

        Private Class AOEExpSeed
            Public vecLoc As Vector3
            Public vecNextSeed As Vector3
            Public yActive As Byte
        End Class

        Private moSeed() As AOEExpSeed
        Private mlSeedUB As Int32 = -1
        Private mlCnt As Int32 = 0
        Public Function Update() As Boolean
            Dim bResult As Boolean = True
            Dim lRndCnt As Int32 = 0

            Dim lCurUB As Int32 = mlSeedUB
            For X As Int32 = 0 To lCurUB
                If moSeed(X) Is Nothing = False AndAlso moSeed(X).yActive <> 0 Then

                    Dim fDistRatio As Single = ((moSeed(X).vecLoc - vecGroundZero).Length / (lAOERange * gl_FINAL_GRID_SQUARE_SIZE))
                    If fDistRatio < 1.0F Then 'AndAlso moRandom.Next(0, 100) < 90 Then 'fDistRatio * 2 < moRandom.NextDouble Then
                        moSeed(X).yActive = 0

                        mlSeedUB += 1
                        If moSeed.GetUpperBound(0) < mlSeedUB Then
                            ReDim Preserve moSeed(mlSeedUB + 100)
                        End If
                        moSeed(mlSeedUB) = New AOEExpSeed()
                        With moSeed(mlSeedUB)
                            mlCnt += 1

                            .yActive = 1
                            .vecLoc = moSeed(X).vecLoc + moSeed(X).vecNextSeed
                            .vecLoc.X += moRandom.Next(-300, 300) 'lAOERange, lAOERange)
                            .vecLoc.Z += moRandom.Next(-300, 300) 'lAOERange, lAOERange)

                            Dim fY As Single
                            If moPlanet Is Nothing = False Then
                                fY = CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(.vecLoc.X, .vecLoc.Z, True) + 500
                            Else
                                fY = vecGroundZero.Y
                            End If

                            .vecNextSeed = moSeed(X).vecNextSeed

                            Dim lVal As Int32 = Math.Min(lAOERange, 20) * gl_FINAL_GRID_SQUARE_SIZE
                            Dim fS As Single = lVal + ((1 - fDistRatio) * (lVal * 2))  '700 + ((1 - fDistRatio) * 500 * 2)
                            Dim fFS As Single = fS * 0.5F '1.5F

                            goExplMgr.Add(.vecLoc, moRandom.Next(0, 360), moRandom.Next(-25, 25) * 0.1F, moRandom.Next(0, 4), fS, moRandom.Next(-5, 5), Color.White, Color.White, 30, fFS, False)

                            lRndCnt += 1
                        End With
                        lRndCnt += 1
                        bResult = False
                    Else
                        bResult = False
                    End If

                End If
            Next X
            If bResult = False AndAlso lRndCnt = 0 Then
                bResult = True
            End If

            Return bResult
        End Function

        Public Sub SpawnSeeds()
            'Ok, the initial blast
            'Dim lRealSize As Int32 = lAOERange * gl_FINAL_GRID_SQUARE_SIZE
            'goExplMgr.Add(vecGroundZero, CSng(moRandom.NextDouble), CSng(moRandom.NextDouble), moRandom.Next(0, 4), lRealSize, CSng(moRandom.NextDouble), Color.White, Color.White, 30)

            Dim lOffshoots As Int32 = 0
            Dim bBothYDirs As Boolean = False
            If goCurrentEnvir Is Nothing = False Then
                If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    bBothYDirs = True
                End If
            End If
            If bBothYDirs = True Then
                lOffshoots = moRandom.Next(2, 7)
            Else
                lOffshoots = moRandom.Next(2, 4)
            End If 

            Dim lBaseSeedUB As Int32 = 32
            mlSeedUB = lBaseSeedUB + lOffshoots
            ReDim moSeed(mlSeedUB)
            moSeed(0) = New AOEExpSeed()
            With moSeed(0) : .vecLoc = vecGroundZero : .vecNextSeed = New Vector3(0, 0, 0) : .yActive = 0 : End With
            goExplMgr.Add(vecGroundZero, CSng(moRandom.NextDouble), CSng(moRandom.NextDouble), moRandom.Next(0, 4), lAOERange * gl_FINAL_GRID_SQUARE_SIZE, CSng(moRandom.NextDouble), Color.White, Color.White, 30, lAOERange * gl_FINAL_GRID_SQUARE_SIZE * 2, False)

            If goSound Is Nothing = False Then
                Dim sOverallSFX As String = "Explosions\MediumGroundDeath1.wav"
                If moRandom.Next(0, 100) < 50 Then
                    sOverallSFX = "Explosions\MediumGroundDeath2.wav"
                End If
                goSound.StartSound(sOverallSFX, False, SoundMgr.SoundUsage.eDeathExplosions, vecGroundZero, New Vector3(0, 0, 0))
            End If

            Dim fRot As Single = 0
            Dim fRotChg As Single = 360.0F / lBaseSeedUB
            Dim lMaxRng As Int32 = 300 ' Math.Min(300, lAOERange * gl_FINAL_GRID_SQUARE_SIZE)
            Dim lMinRng As Int32 = 300 'lMaxRng \ 2

            For X As Int32 = 1 To lBaseSeedUB
                moSeed(X) = New AOEExpSeed
                With moSeed(X)
                    Dim fX As Single = moRandom.Next(lMinRng, lMaxRng) 'lAOERange * 2 * gl_FINAL_GRID_SQUARE_SIZE
                    Dim fZ As Single = 0
                    RotatePoint(0, 0, fX, fZ, fRot)
                    fRot += fRotChg

                    .vecNextSeed = New Vector3(fX, 0, fZ)
                    .yActive = 1
                    fX += vecGroundZero.X
                    fZ += vecGroundZero.Z

                    Dim fY As Single
                    If moPlanet Is Nothing = False Then
                        fY = CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(fX, fZ, True) + 500
                    Else
                        fY = vecGroundZero.Y
                    End If
                    .vecLoc = New Vector3(fX, fY, fZ)
                End With
            Next X

            If lOffshoots > 0 Then
                For X As Int32 = 1 To lOffshoots

                    Dim fYVec As Single = moRandom.Next(0, lMaxRng)
                    If bBothYDirs = True Then
                        If moRandom.Next(0, 100) < 50 Then
                            fYVec = -fYVec
                        End If
                    End If

                    moSeed(X + lBaseSeedUB) = New AOEExpSeed
                    With moSeed(X + lBaseSeedUB)
                        Dim fX As Single = moRandom.Next(lMinRng, lMaxRng)
                        Dim fZ As Single = 0
                        RotatePoint(0, 0, fX, fZ, moRandom.Next(0, 360))

                        .vecNextSeed = New Vector3(fX, fYVec, fZ)
                        .yActive = 1
                        fX += vecGroundZero.X
                        fZ += vecGroundZero.Z

                        Dim fY As Single
                        If moPlanet Is Nothing = False Then
                            fY = CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(fX, fZ, True) + 500
                        Else
                            fY = vecGroundZero.Y
                        End If
                        .vecLoc = New Vector3(fX, fY + fYVec, fZ)
                    End With
                Next X
            End If 


        End Sub
    End Class

    Private moItems(-1) As AOEExplosion
    Private mlLastUpdate As Int32 = 0

    Public Sub AddNew(ByVal vecLoc As Vector3, ByVal lRng As Int32)
        Dim oPlanet As Planet = Nothing

        Try
            If goCurrentEnvir Is Nothing = False Then
                If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                    oPlanet = CType(goCurrentEnvir.oGeoObject, Planet)
                End If
            End If
        Catch
        End Try
        If oPlanet Is Nothing = False Then
            vecLoc.Y = oPlanet.GetHeightAtPoint(vecLoc.X, vecLoc.Z, True)
        End If

        For X As Int32 = 0 To moItems.GetUpperBound(0)
            If moItems(X) Is Nothing = True Then
                moItems(X) = New AOEExplosion()
                With moItems(X)
                    .vecGroundZero = vecLoc
                    .lAOERange = lRng
                    .mlAOERangeSqrd = lRng * lRng
                    .SpawnSeeds()
                End With
                Return
            End If
        Next X
        ReDim Preserve moItems(moItems.GetUpperBound(0) + 1)
        moItems(moItems.GetUpperBound(0)) = New AOEExplosion()
        With moItems(moItems.GetUpperBound(0))
            .vecGroundZero = vecLoc
            .lAOERange = lRng
            .mlAOERangeSqrd = lRng * lRng
            .SpawnSeeds()
        End With
    End Sub

    Public Sub Update()
        glAOEExplRendered = 0

        moPlanet = Nothing
        Try
            If goCurrentEnvir Is Nothing = False Then
                If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                    moPlanet = CType(goCurrentEnvir.oGeoObject, Planet)
                End If
            End If
        Catch
        End Try

        If glCurrentCycle = mlLastUpdate Then Return
        mlLastUpdate = glCurrentCycle
        For X As Int32 = 0 To moItems.GetUpperBound(0)
            If moItems(X) Is Nothing = False Then
                glAOEExplRendered += 1
                If moItems(X).Update() = True Then
                    moItems(X) = Nothing
                End If
            End If
        Next X
        If glAOEExplRendered = 0 AndAlso moItems.GetUpperBound(0) > -1 Then ReDim moItems(-1)
    End Sub
End Class
