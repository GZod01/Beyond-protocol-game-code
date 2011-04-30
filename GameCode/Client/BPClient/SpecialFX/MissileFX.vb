Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D


Public Class MissileMgr
    Public Class Missile
        Private Structure Particle
            Public vecLoc As Vector3
            Public vecAcc As Vector3
            Public vecSpeed As Vector3

            Public mfA As Single
            Public mfR As Single
            Public mfG As Single
            Public mfB As Single

            Public fAChg As Single
            Public fRChg As Single
            Public fGChg As Single
            Public fBChg As Single

            Public ParticleColor As System.Drawing.Color

            Public ParticleActive As Boolean

            Public Sub Update(ByVal fElapsedTime As Single)
                vecLoc.Add(Vector3.Multiply(vecSpeed, fElapsedTime))

                vecSpeed.Add(Vector3.Multiply(vecAcc, fElapsedTime))

                mfA += (fAChg * fElapsedTime)
                mfR += (fRChg * fElapsedTime)
                mfG += (fGChg * fElapsedTime)
                mfB += (fBChg * fElapsedTime)

                If mfA < 0 Then mfA = 0
                If mfA > 255 Then mfA = 255
                If mfR < 0 Then mfR = 0
                If mfR > 255 Then mfR = 255
                If mfG < 0 Then mfG = 0
                If mfG > 255 Then mfG = 255
                If mfB < 0 Then mfB = 0
                If mfB > 255 Then mfB = 255

                ParticleColor = System.Drawing.Color.FromArgb(CInt(mfA), CInt(mfR), CInt(mfG), CInt(mfB))
            End Sub

            Public Sub Reset(ByVal fX As Single, ByVal fY As Single, ByVal fZ As Single, ByVal fXSpeed As Single, ByVal fYSpeed As Single, ByVal fZSpeed As Single, ByVal fXAcc As Single, ByVal fYAcc As Single, ByVal fZAcc As Single, ByVal fR As Single, ByVal fG As Single, ByVal fB As Single, ByVal fA As Single)
                vecLoc.X = fX : vecLoc.Y = fY : vecLoc.Z = fZ
                vecAcc.X = fXAcc : vecAcc.Y = fYAcc : vecAcc.Z = fZAcc
                vecSpeed.X = fXSpeed : vecSpeed.Y = fYSpeed : vecSpeed.Z = fZSpeed
                mfA = fA : mfR = fR : mfG = fG : mfB = fB

                ParticleActive = True
            End Sub

        End Structure

        Public lID As Int32

        Private mvecPrevious As Vector3

        Public vecMissile As Vector3
        Public vecSpeed As Vector3

        Public MaxSpeed As Byte
        Private mfSpeed As Single = 0.0F

        Public oTarget As BaseEntity

        Public vecTarget As Vector3

        Private mbEmitterStopping As Boolean = False

        Private moParticles() As Particle

        Public moPoints() As CustomVertex.PositionColoredTextured
        Public mlParticleUB As Int32

        Public EmitterStopped As Boolean = False

        Public Event MissileExploded(ByRef oTarget As BaseEntity, ByVal vecLoc As Vector3, ByVal yHitType As Byte, ByVal vecOnModelTarget As Vector3)

        Private lBaseR As Int32 = 64
        Private lBaseG As Int32 = 255
        Private lBaseB As Int32 = 72

        Private iLocAngle As Int16

        Private mlSoundIdx As Int32 = -1
        Public bNoSound As Boolean = False

        Public yAOERange As Byte = 0

        Public yHit As Byte = 0     '0 = no hit, 1 = Hit Target, 2 = Missed Target

        Private mfAcceleration As Single = 0.0F

        Private mvecOnModelTarget As Vector3 = Vector3.Empty

        Private mvecPrevTarget As Vector3
        Private mlPrevTargetSet As Int32 = 0

        Public lGroundGlowFXID As Int32 = -1
        'Private mlStartCycle As Int32 = Int32.MinValue

        Protected Sub ResetParticle(ByVal lIndex As Int32)
            Dim fX As Single
            Dim fY As Single
            Dim fZ As Single
            Dim fXS As Single
            Dim fYS As Single
            Dim fZS As Single
            Dim fXA As Single
            Dim fYA As Single
            Dim fZA As Single
            Dim fOffsetX As Single
            Dim fOffsetY As Single

            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False
                moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lX).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                'Speed
                fXS = Rnd() * 1.0F - 0.1F
                fYS = -vecSpeed.Y '(GetNxtRnd() * 1.0F - 0.1F) * -vecSpeed.Y
                fZS = MaxSpeed  '10.0F

                'Accel
                If fXS < 0 Then fXA = Rnd() * -0.01F Else fXA = Rnd() * 0.01F
                If fYS < 0 Then fYA = Rnd() * -0.01F Else fYA = Rnd() * 0.01F
                fZA = 0

                fX = fOffsetX
                fY = fOffsetY
                fZ = 0

                'Now, rotate everyone
                RotateVals(fX, fY, fZ, fXS, fZS, fXA, fZA)

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, lBaseR, lBaseG, lBaseB, 128) '255 
                moParticles(lIndex).fAChg = -((0.1F + (0.1F * Rnd())) * 255)

            End If


        End Sub

        Public Sub Update(ByVal fElapsed As Single)
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Exit Sub

            If oTarget Is Nothing = False AndAlso yHit <> 2 Then
                If mvecOnModelTarget = Vector3.Empty Then
                    If oTarget.oMesh Is Nothing = False Then
                        Dim fAngle As Single = LineAngleDegrees(CInt(vecMissile.X), CInt(vecMissile.Z), CInt(oTarget.LocX), CInt(oTarget.LocZ)) - 180
                        Dim fMyAngle As Single = (oTarget.LocAngle / 10.0F)
                        fAngle -= fMyAngle
                        If fAngle > 360 Then fAngle -= 360
                        If fAngle < 0 Then fAngle += 360
                        Dim ySideHit As Byte = AngleToQuadrant(CInt(fAngle))
                        mvecOnModelTarget = oTarget.oMesh.GetRandomMeshPoint(ySideHit)
                    Else
                        mvecOnModelTarget = New Vector3(0, 0, 0)
                    End If
                End If
                'vecTarget.X = oTarget.fMapWrapLocX
                'vecTarget.Y = oTarget.LocY
                'vecTarget.Z = oTarget.LocZ
                If oTarget.oMesh Is Nothing = False Then
                    If glCurrentCycle - mlPrevTargetSet > 10 Then
                        With oTarget.GetWorldMatrix
                            Dim fX As Single = mvecOnModelTarget.X
                            Dim fY As Single = mvecOnModelTarget.Y
                            Dim fZ As Single = mvecOnModelTarget.Z
                            mvecPrevTarget.X = (fX * .M11) + (fY * .M21) + (fZ * .M31) + .M41
                            mvecPrevTarget.Y = (fX * .M12) + (fY * .M22) + (fZ * .M32) + .M42
                            mvecPrevTarget.Z = (fX * .M13) + (fY * .M23) + (fZ * .M33) + .M43
                        End With
                    End If
                    vecTarget = mvecPrevTarget 'Vector3.TransformCoordinate(mvecOnModelTarget, oTarget.GetWorldMatrix)
                Else
                    vecTarget.X = oTarget.LocX
                    vecTarget.Y = oTarget.LocY
                    vecTarget.Z = oTarget.LocZ
                End If
            End If

            'TODO: Determine where we will hit the target...
            'Dim matTemp As Matrix = oTarget.GetWorldMatrix()
            'matTemp.Invert()
            'Dim vecNear As Vector3 = Vector3.TransformCoordinate(vecMissile, matTemp)
            'Dim vecDir As Vector3 = Vector3.TransformNormal(Vector3.Subtract(vecMissile, vecTarget), matTemp)
            'Dim vecInt As IntersectInformation
            'oTarget.oMesh.oMesh.Intersect(vecMissile, vecDir, vecInt)

            'Dim oIntLoc As IntersectInformation
            'Dim vecFrom As Vector3 = vecMissile
            'Dim oDir As Vector3 = Vector3.Subtract(vecFrom, vecTarget)
            'oDir.Normalize()
            'oTarget.oMesh.oMesh.Intersect(vecFrom, oDir, oIntLoc)
            'oDir.Scale(oIntLoc.Dist)
            'vecTarget += Vector3.Add(vecFrom, oDir)

            ''=========================================

            'Move the Missile first
            'NOTE: if you experience issues with the movement of missiles syncing with the movement on the server, update this or the server
            mfSpeed += (mfAcceleration * fElapsed)
            If mfSpeed > MaxSpeed Then mfSpeed = MaxSpeed
            If mbEmitterStopping = False Then
                Dim vecAcc As Vector3 = Vector3.Subtract(vecTarget, vecMissile) 'vecMissile, vecTarget)

                If vecMissile.Y <> vecTarget.Y Then
                    vecSpeed.Y = Vector3.Multiply(Vector3.Normalize(vecAcc), mfSpeed).Y
                End If

                vecAcc.Y = 0
                vecAcc.Normalize()
                vecAcc.Multiply(mfAcceleration * fElapsed)
                vecSpeed.Add(vecAcc)
                Dim fTotalSpeed As Single = Math.Abs(vecSpeed.X) + Math.Abs(vecSpeed.Z)
                If fTotalSpeed > MaxSpeed Then
                    Dim fTmp As Single = MaxSpeed / fTotalSpeed
                    vecSpeed.X *= fTmp
                    vecSpeed.Z *= fTmp
                End If

                If yHit <> 0 OrElse Int32.MaxValue = lID Then
                    If Math.Abs(vecSpeed.X) > Math.Abs(vecMissile.X - vecTarget.X) Then
                        vecSpeed.X = 0
                        vecMissile.X = vecTarget.X
                    End If
                    If Math.Abs(vecSpeed.Z) > Math.Abs(vecMissile.Z - vecTarget.Z) Then
                        vecSpeed.Z = 0
                        vecMissile.Z = vecTarget.Z
                    End If
                End If

                If yHit = 1 AndAlso Math.Abs(vecTarget.X - vecMissile.X) + Math.Abs(vecTarget.Z - vecMissile.Z) < (fTotalSpeed * 0.5F) Then
                    mbEmitterStopping = True
                    If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
                    RaiseEvent MissileExploded(oTarget, vecMissile, yHit, mvecOnModelTarget)
                ElseIf yHit = 2 OrElse fTotalSpeed < 0.00001F Then
                    mbEmitterStopping = True
                    If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
                    RaiseEvent MissileExploded(oTarget, vecMissile, yHit, mvecOnModelTarget)
                ElseIf Math.Abs(vecTarget.X - vecMissile.X) + Math.Abs(vecTarget.Z - vecMissile.Z) < (fTotalSpeed * 0.5F) Then
                    mbEmitterStopping = True
                    If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
                    RaiseEvent MissileExploded(oTarget, vecMissile, yHit, mvecOnModelTarget)
                Else
                    vecMissile.Add(Vector3.Multiply(vecSpeed, fElapsed))
                    If bNoSound = False AndAlso goSound Is Nothing = False Then

                        Select Case Me.lID Mod 3
                            Case 0
                                mlSoundIdx = goSound.HandleMovingSound(mlSoundIdx, "Hiss1.wav", SoundMgr.SoundUsage.eWeaponsFireMissiles, vecMissile, vecSpeed)
                            Case 1
                                mlSoundIdx = goSound.HandleMovingSound(mlSoundIdx, "Hiss2.wav", SoundMgr.SoundUsage.eWeaponsFireMissiles, vecMissile, vecSpeed)
                            Case 2
                                mlSoundIdx = goSound.HandleMovingSound(mlSoundIdx, "Hiss3.wav", SoundMgr.SoundUsage.eWeaponsFireMissiles, vecMissile, vecSpeed)
                        End Select
                        

                    End If
                End If
                'If mbEmitterStopping = True Then
                '    If lGroundGlowFXID <> -1 Then TerrainClass.RemoveGroundLightAccent(lGroundGlowFXID)
                '    lGroundGlowFXID = -1
                'End If
 
            End If

            'Set our locAngle
            iLocAngle = CShort(LineAngleDegrees(CInt(vecMissile.X), CInt(vecMissile.Z), CInt(vecMissile.X + vecSpeed.X), CInt(vecMissile.Z + vecSpeed.Z)) * 10)

            'If mbEmitterStopping = False AndAlso (yHit <> 0 OrElse glCurrentCycle > lFltTimeEnd) Then
            '	mbEmitterStopping = True
            '	If mlSoundIdx <> -1 Then goSound.StopSound(mlSoundIdx)
            '	RaiseEvent MissileExploded(oTarget, vecMissile, yHit)
            'End If

            fElapsed = 0.2F
            Dim bActive As Boolean = False
            For X = 0 To mlParticleUB
                With moParticles(X)
                    If .ParticleActive = True Then
                        .Update(fElapsed)
                        If .mfA <= 0 Then
                            ResetParticle(X)
                        End If

                        moPoints(X).Color = .ParticleColor.ToArgb
                        moPoints(X).Position = Vector3.Add(.vecLoc, vecMissile)
                        bActive = True
                    Else
                        moPoints(X).Color = Color.FromArgb(0, 0, 0, 0).ToArgb
                    End If
                End With
            Next X
            If bActive = False Then EmitterStopped = True

            'If mlStartCycle = Int32.MinValue Then mlStartCycle = glCurrentCycle
            'If mbEmitterStopping = False AndAlso (glCurrentCycle - mlStartCycle > 9000) Then
            '    mbEmitterStopping = True
            'End If
        End Sub

        Private Sub RotateVals(ByRef fXLoc As Single, ByRef fYLoc As Single, ByRef fZLoc As Single, ByRef fXSpeed As Single, ByRef fZSpeed As Single, ByRef fXAcc As Single, ByRef fZAcc As Single)
            Dim fDX As Single
            Dim fDZ As Single

            Dim fRads As Single = 0
            Dim fCosR As Single = CSng(Math.Cos(fRads))
            Dim fSinR As Single = CSng(Math.Sin(fRads))

            'Yaw...
            fDX = fXLoc
            fDZ = fYLoc

            fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

            'Now set up for standard rotation...
            fRads = ((iLocAngle / 10.0F) - 90) * gdRadPerDegree
            fCosR = CSng(Math.Cos(fRads))
            fSinR = CSng(Math.Sin(fRads))

            'Loc
            fDX = fXLoc
            fDZ = fZLoc
            fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

            'Speed
            fDX = fXSpeed
            fDZ = fZSpeed
            fXSpeed = +((fDX * fCosR) + (fDZ * fSinR))
            fZSpeed = -((fDX * fSinR) - (fDZ * fCosR))

            'Acc
            fDX = fXAcc
            fDZ = fZAcc
            fXAcc = +((fDX * fCosR) + (fDZ * fSinR))
            fZAcc = -((fDX * fSinR) - (fDZ * fCosR))
        End Sub

        Public Sub New(ByVal oVecLoc As Vector3, ByRef poTarget As BaseEntity, ByVal yMaxSpeed As Byte, ByVal yManeuver As Byte, ByVal clrVal As System.Drawing.Color, ByVal plID As Int32)
            'lFltTimeEnd = glCurrentCycle + iFltTime
            lID = plID
            lBaseR = clrVal.R
            lBaseG = clrVal.G
            lBaseB = clrVal.B

            oTarget = poTarget
            MaxSpeed = yMaxSpeed
            vecMissile = oVecLoc
            mvecPrevious = oVecLoc
            vecSpeed = New Vector3(0, 0, 0)

            mfAcceleration = yManeuver / 10.0F

            'If oTarget.fMapWrapLocX = Single.MinValue Then oTarget.fMapWrapLocX = oTarget.LocX
            iLocAngle = CShort(LineAngleDegrees(CInt(vecMissile.X), CInt(vecMissile.Z), CInt(oTarget.LocX), CInt(oTarget.LocZ)) * 10)

            Dim lParticleCnt As Int32 = CInt(50 * (muSettings.BurnFXParticles / 4.0F))  ' 200 'Math.Max(200 + (CInt(yMaxSpeed) - 50), 200)
            If lParticleCnt < 15 Then lParticleCnt = 15

            mlParticleUB = lParticleCnt - 1
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB)

            For X As Int32 = 0 To mlParticleUB
                ResetParticle(X)
            Next X

        End Sub

        Public Sub KillMe()
            EmitterStopped = True
            If mlSoundIdx <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(mlSoundIdx)
            'If lGroundGlowFXID <> -1 Then TerrainClass.RemoveGroundLightAccent(lGroundGlowFXID)
            'lGroundGlowFXID = -1
        End Sub

        Public Function GetCullBox() As CullBox
            Return CullBox.GetCullBox(vecMissile.X, vecMissile.Y, vecMissile.Z, -30, -30, -30, 30, 30, 30)
        End Function

        'Public Sub RenderToTerrain()
        '    Try
        '        If mbEmitterStopping = True Then
        '            If lGroundGlowFXID <> -1 Then TerrainClass.RemoveGroundLightAccent(lGroundGlowFXID)
        '            lGroundGlowFXID = -1
        '            Return
        '        End If
        '        Dim oPlanet As Planet = CType(goCurrentEnvir.oGeoObject, Planet)
        '        If oPlanet Is Nothing = False Then
        '            Dim fTHt As Single = oPlanet.GetHeightAtPoint(vecMissile.X, vecMissile.Z, True)
        '            Dim fMissileHtFromGround As Single = Math.Abs(vecMissile.Y - fTHt)
        '            If fMissileHtFromGround < 1500 Then
        '                Dim fRad As Single = 1.0F - (fMissileHtFromGround / 1500.0F)
        '                fRad *= 100
        '                'fRad = 100
        '                fRad = 500

        '                Dim yA As Byte = CByte((Rnd() * 64) + 180)
        '                If lGroundGlowFXID = -1 Then
        '                    lGroundGlowFXID = TerrainClass.AddGroundLightAccent(vecMissile.X, vecMissile.Z, fRad, System.Drawing.Color.FromArgb(yA, lBaseR, lBaseG, lBaseB), 0)
        '                Else
        '                    TerrainClass.UpdateGroundLightAccent(lGroundGlowFXID, vecMissile.X, vecMissile.Z, fRad, System.Drawing.Color.FromArgb(yA, lBaseR, lBaseG, lBaseB), 0)
        '                End If
        '            Else
        '                If lGroundGlowFXID <> -1 Then TerrainClass.RemoveGroundLightAccent(lGroundGlowFXID)
        '                lGroundGlowFXID = -1
        '            End If
        '        End If
        '    Catch
        '    End Try
        'End Sub
    End Class

    Private moMissiles() As Missile
    Private mlMissileIdx() As Int32
    Private mlMissileUB As Int32 = -1

    Private moTex As Texture

    Private clrVals() As System.Drawing.Color
    Private lClrIdx As Int32 = -1

    Private swUpdate As Stopwatch

    Public Sub MissileBreakApart(ByVal lMissileID As Int32)
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) = lMissileID Then
                MissileExploded(Nothing, moMissiles(X).vecMissile, 0, New Vector3(0, 0, 0))
                moMissiles(X).EmitterStopped = True
                moMissiles(X).KillMe()
                Exit For
            End If
        Next X
    End Sub

    Public Sub MissileDetonated(ByVal lMissileID As Int32, ByVal lX As Int32, ByVal lZ As Int32, ByVal yAOERange As Byte)
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) = lMissileID Then
                Dim oMissile As Missile = moMissiles(X)
                If oMissile Is Nothing Then Return
                With oMissile
                    .yHit = 2
                    .oTarget = Nothing
                    .vecTarget.X = lX
                    .vecTarget.Z = lZ
                    .KillMe()
                End With

                Exit For
            End If
        Next X
    End Sub

    'Public Sub MissileImpact(ByVal lMissileID As Int32, ByVal yAOERange As Byte)
    '    'TODO: Not sure what to do here... is impact still needed?
    'End Sub

    Public Function GetMissileLoc(ByVal lMissileID As Int32) As Vector3
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) = lMissileID Then Return moMissiles(X).vecMissile
        Next X
        Return Vector3.Empty
    End Function

    Public Function AddWpnBuilderMissile(ByVal vecFrom As Vector3, ByVal vecTo As Vector3, ByVal MaxSpeed As Byte, ByVal Maneuver As Byte, ByVal lClrIdx As Int32, ByVal lID As Int32) As Int32
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) = lID Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlMissileUB += 1
            ReDim Preserve moMissiles(mlMissileUB)
            ReDim Preserve mlMissileIdx(mlMissileUB)
            mlMissileIdx(mlMissileUB) = -1
            lIdx = mlMissileUB
        End If

        If lClrIdx > clrVals.GetUpperBound(0) Then lClrIdx = 0

        Dim oTarget As New BaseEntity
        oTarget.LocX = vecTo.X
        oTarget.LocY = CInt(vecTo.Y)
        oTarget.LocZ = vecTo.Z
        'oTarget.fMapWrapLocX = oTarget.LocX

        moMissiles(lIdx) = New Missile(vecFrom, oTarget, MaxSpeed, Maneuver, clrVals(lClrIdx), lID)
        moMissiles(lIdx).bNoSound = True
        mlMissileIdx(lIdx) = lID

        AddHandler moMissiles(lIdx).MissileExploded, AddressOf MissileExploded

        Return lID
    End Function

    Public Sub AddMissile(ByVal vecFrom As Vector3, ByRef oTarget As BaseEntity, ByVal MaxSpeed As Byte, ByVal Maneuver As Byte, ByVal lClrIdx As Int32, ByVal lID As Int32, ByVal bHit As Boolean)

        'If oTarget Is Nothing Then Return

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) = -1 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlMissileUB += 1
            ReDim Preserve moMissiles(mlMissileUB)
            ReDim Preserve mlMissileIdx(mlMissileUB)
            mlMissileIdx(mlMissileUB) = -1
            lIdx = mlMissileUB
        End If

        If lClrIdx > clrVals.GetUpperBound(0) Then lClrIdx = 0

        moMissiles(lIdx) = New Missile(vecFrom, oTarget, MaxSpeed, Maneuver, clrVals(lClrIdx), lID)
        If bHit = True Then moMissiles(lIdx).yHit = 1 Else moMissiles(lIdx).yHit = 0
        mlMissileIdx(lIdx) = lID

        If goSound Is Nothing = False Then
            If lClrIdx < 4 Then
                goSound.StartSound("Launch1.wav", False, SoundMgr.SoundUsage.eWeaponsFireMissiles, vecFrom, Nothing)
            Else
                goSound.StartSound("Launch2.wav", False, SoundMgr.SoundUsage.eWeaponsFireMissiles, vecFrom, Nothing)
            End If

        End If

        AddHandler moMissiles(lIdx).MissileExploded, AddressOf MissileExploded
    End Sub

    Public Sub New()
        ReDim mlMissileIdx(0)
        mlMissileIdx(0) = -1

        ReDim clrVals(8)
        'Now, set the Engines values
        clrVals(0) = System.Drawing.Color.FromArgb(255, 64, 128, 255)
        clrVals(1) = System.Drawing.Color.FromArgb(255, 64, 255, 64)
        clrVals(2) = System.Drawing.Color.FromArgb(255, 255, 128, 64)
        clrVals(3) = System.Drawing.Color.FromArgb(255, 64, 64, 255)
        clrVals(4) = System.Drawing.Color.FromArgb(255, 192, 64, 255)
        clrVals(5) = System.Drawing.Color.FromArgb(255, 64, 64, 92)
        clrVals(6) = System.Drawing.Color.FromArgb(255, 255, 64, 64)
        clrVals(7) = System.Drawing.Color.FromArgb(255, 255, 255, 32)
        clrVals(8) = System.Drawing.Color.FromArgb(255, 32, 128, 64)

        swUpdate = Stopwatch.StartNew()
    End Sub

    Public Sub Render(ByVal bUpdateNoRender As Boolean)
        'Dim matWorld As Matrix
        Dim X As Int32

        If moTex Is Nothing OrElse moTex.Disposed = True Then
            moTex = goResMgr.GetTexture("Flare2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak")
        End If

        Dim fElapsed As Single = swUpdate.ElapsedMilliseconds / 30.0F
        swUpdate.Reset()
        swUpdate.Start()
        'Dim fElapsed As Single = 1.0F

        If mlMissileUB = -1 Then Exit Sub

        For X = 0 To mlMissileUB
            If mlMissileIdx(X) > -1 Then
                moMissiles(X).Update(fElapsed)
                If moMissiles(X).EmitterStopped = True Then mlMissileIdx(X) = -1
            End If
        Next X

        Dim moDevice As Device = GFXEngine.moDevice

        If bUpdateNoRender = False Then
            'Dim bRenderTerrain As Boolean = False
            'If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then bRenderTerrain = True

            'And now render them... first set up our device for renders...
            With moDevice
                .Transform.World = Matrix.Identity

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

                'Ok, if our device was created with mixed...
                If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                    'Set us up for software vertex processing as point sprites always work in software
                    moDevice.SoftwareVertexProcessing = True
                End If

                .RenderState.PointSpriteEnable = True

                If .DeviceCaps.VertexFormatCaps.SupportsPointSize = True Then
                    .RenderState.PointScaleEnable = True
                    If .DeviceCaps.MaxPointSize > 16 Then        '32
                        .RenderState.PointSize = 16              '32
                    Else : .RenderState.PointSize = .DeviceCaps.MaxPointSize
                    End If
                    .RenderState.PointScaleA = 0
                    .RenderState.PointScaleB = 0
                    .RenderState.PointScaleC = 0.8F
                End If

                .RenderState.SourceBlend = Blend.SourceAlpha
                .RenderState.DestinationBlend = Blend.One
                .RenderState.AlphaBlendEnable = True

                .RenderState.ZBufferWriteEnable = False

                .RenderState.Lighting = False

                .VertexFormat = CustomVertex.PositionColoredTextured.Format
                '.VertexFormat = moPoints(0).Format
                .SetTexture(0, moTex)

                'Now render everything 
                If glCurrentEnvirView < CurrentView.eFullScreenInterface Then
                    For X = 0 To mlMissileUB
                        If mlMissileIdx(X) > -1 Then
                            'If goCamera.CullObject(moMissiles(X).GetCullBox) = False Then
                            .DrawUserPrimitives(PrimitiveType.PointList, moMissiles(X).mlParticleUB + 1, moMissiles(X).moPoints)
                            'End If


                            'If bRenderTerrain = True Then
                            '    moMissiles(X).RenderToTerrain()
                            'End If
                        End If
                    Next X
                ElseIf glCurrentEnvirView = CurrentView.eStartupLogin Then
                    For X = 0 To mlMissileUB
                        If mlMissileIdx(X) > -1 Then
                            .DrawUserPrimitives(PrimitiveType.PointList, moMissiles(X).mlParticleUB + 1, moMissiles(X).moPoints)
                        End If
                    Next X
                ElseIf glCurrentEnvirView = CurrentView.eWeaponBuilder Then
                    For X = 0 To mlMissileUB
                        If mlMissileIdx(X) > -1 Then
                            .DrawUserPrimitives(PrimitiveType.PointList, moMissiles(X).mlParticleUB + 1, moMissiles(X).moPoints)
                        End If
                    Next X
                End If

                'Then, reset our device...
                .RenderState.ZBufferWriteEnable = True
                .RenderState.Lighting = True
                .RenderState.PointSpriteEnable = False

                'Ok, if our device was created with mixed...
                If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                    'Set us up for software vertex processing as point sprites always work in software
                    moDevice.SoftwareVertexProcessing = False
                End If

                .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                .RenderState.AlphaBlendEnable = True

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With
        End If
    End Sub

    Public Sub KillAll()
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) <> -1 Then
                moMissiles(X).KillMe()
            End If
        Next X
        mlMissileUB = -1
        Erase moMissiles
        Erase mlMissileIdx
    End Sub

    Public Sub ClearTex()
        moTex = Nothing
    End Sub

    Public Sub MissileExploded(ByRef oTarget As BaseEntity, ByVal vecLoc As Vector3, ByVal yHitType As Byte, ByVal vecOnModelTarget As Vector3)
        'for now, just add a screenshake
        Dim fXT As Single = vecLoc.X - goCamera.mlCameraAtX
        Dim fYT As Single = vecLoc.Z - goCamera.mlCameraAtZ
        fXT *= fXT
        fYT *= fYT
        Dim dDist As Double = Math.Sqrt(fXT + fYT)
        If dDist < 10000 Then goCamera.ScreenShake(100 * CSng(dDist / 10000))

        If goSound Is Nothing = False Then
            goSound.StartSound("Explosions\MissileHit.wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, vecLoc, New Vector3(0, 0, 0))
        End If

        If goExplMgr Is Nothing = False Then
            'goExplMgr.AddMinorExplosion(vecLoc, 1)
            goExplMgr.Add(vecLoc, Rnd() * 360, Rnd(), CInt(Rnd() * 3), 200, 0, Color.White, Color.Red, 60, 400, True)
            Dim vecTemp As Vector3 = New Vector3(vecLoc.X + (Rnd() * 50 - 100), vecLoc.Y + (Rnd() * 50 - 100), vecLoc.Z + (Rnd() * 50 - 100))
            goExplMgr.Add(vecTemp, Rnd() * 360, Rnd(), CInt(Rnd() * 3), 220, 0, Color.White, Color.Red, 30, 440, True)
            vecTemp = New Vector3(vecLoc.X + (Rnd() * 50 - 100), vecLoc.Y + (Rnd() * 50 - 100), vecLoc.Z + (Rnd() * 50 - 100))
            goExplMgr.Add(vecTemp, Rnd() * 360, Rnd(), CInt(Rnd() * 3), 220, 0, Color.White, Color.Red, 45, 440, True)
        End If
        If oTarget Is Nothing = False AndAlso yHitType <> 0 Then
            'Now, determine what side I hit...
            Dim fAngle As Single = LineAngleDegrees(CInt(vecLoc.X), CInt(vecLoc.Z), CInt(oTarget.LocX), CInt(oTarget.LocZ))
            fAngle -= (oTarget.LocAngle * 0.1F)
            If fAngle > 360 Then fAngle -= 360
            If fAngle < 0 Then fAngle += 360
            Dim ySideHit As Byte = AngleToQuadrant(CInt(fAngle))
            oTarget.AddBurnMark(vecOnModelTarget, ySideHit, 60 + Rnd() * 32)

            If glCurrentEnvirView = CurrentView.eStartupLogin Then
                Dim bFound As Boolean = False
                For X As Int32 = 0 To 3
                    If oTarget.yArmorHP(X) > 10 Then
                        oTarget.yArmorHP(X) = CByte(oTarget.yArmorHP(X) - 10)
                        bFound = True
                    Else
                        oTarget.yArmorHP(X) = 0
                    End If
                Next X
                If bFound = False Then
                    If oTarget.yStructureHP > 10 Then oTarget.yStructureHP = CByte(oTarget.yStructureHP - 10) Else oTarget.yStructureHP = 0
                End If

                oTarget.TestForBurnFX()
                goShldMgr.AddNewEffect(CType(oTarget, RenderObject), oTarget.clrShieldFX)
            Else
                '0 = no hit, 1 = hit shield, 2 = detonate, 3 = hit armor
                'If yHitType = 1 Then        'hit shield
                '    If goShldMgr Is Nothing = False Then
                '        goShldMgr.AddNewEffect(CType(oTarget, RenderObject), oTarget.clrShieldFX)
                '    End If

                '    If goSound Is Nothing = False Then
                '        If goSound.MusicOn = True AndAlso oTarget.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(20) 'hit shields +20

                '        'Two shield sound FX files
                '        Dim lTemp As Int32 = CInt(Int(GetNxtRnd() * 2) + 1)
                '        Dim oTmpVec As Microsoft.DirectX.Vector3 = New Microsoft.DirectX.Vector3(oTarget.LocX, oTarget.LocY, oTarget.LocZ)
                '        'Once and done SFX, we dont care about retreiving the id
                '        goSound.StartSound("Explosions\ShieldHit" & lTemp & ".wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, oTmpVec, New Microsoft.DirectX.Vector3(0, 0, 0))
                '    End If
                'ElseIf yHitType = 3 Then    'hit armor
                If goSound Is Nothing = False Then
                    If goSound.MusicOn = True AndAlso oTarget.OwnerID = glPlayerID Then goSound.IncrementExcitementLevel(30) 'hit armor +30
                    Dim lTemp As Int32 = CInt(Int(Rnd() * 2) + 1)
                    Dim oTmpVec As Microsoft.DirectX.Vector3 = New Microsoft.DirectX.Vector3(oTarget.LocX, oTarget.LocY, oTarget.LocZ)
                    'TODO: There are various explosions, but right now, there's only 4 and are all of the same type
                    lTemp = CInt(Int(Rnd() * 4) + 1)
                    'Once and done SFX, we don't care about the ID
                    'goSound.StartSound("Explosions\HullHit" & lTemp & ".wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, oTmpVec, New Microsoft.DirectX.Vector3(0, 0, 0))
                End If
                'End If
            End If
        End If
    End Sub

    Public Function GetMissile(ByVal lMissileID As Int32) As Missile
        For X As Int32 = 0 To mlMissileUB
            If mlMissileIdx(X) = lMissileID Then Return moMissiles(X)
        Next X
        Return Nothing
    End Function
End Class
