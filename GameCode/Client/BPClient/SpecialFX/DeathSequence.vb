Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class DeathSequenceMgr
    Private moSeq() As DeathSequence
    Private mlSeqUB As Int32 = -1
    Private mySeqUsed() As Byte

    Public Sub AddNewDeathSequence(ByRef oEntity As BaseEntity, ByVal lEntityIndex As Int32)
        'ByVal vecLoc As Vector3, ByVal plExpCnt As Int32, ByVal vecSize As Vector3, ByVal pbDoFinalExp As Boolean, ByVal plEntityIndex As Int32, ByVal sWAVFile As String
        If oEntity Is Nothing Then Return

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlSeqUB
            If mySeqUsed(X) = 0 Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            lIdx = mlSeqUB + 1
            ReDim Preserve moSeq(lIdx)
            ReDim Preserve mySeqUsed(lIdx)
            mySeqUsed(lIdx) = 0
            mlSeqUB += 1
        End If

        moSeq(lIdx) = Nothing
        Dim vecLoc As Vector3

        With oEntity
            vecLoc = New Vector3(.LocX, .LocY + (.oMesh.YMidPoint \ 2), .LocZ)       '???? not sure aout adding yMidPoint
            moSeq(lIdx) = New DeathSequence(vecLoc, .oMesh.lDeathSeqExpCnt, .oMesh.vecDeathSeqSize, .oMesh.bDeathSeqFinale, lEntityIndex, .LocAngle, .LocYaw, oEntity)
            mySeqUsed(lIdx) = 255

            If .oMesh.sDeathSeqWav Is Nothing = False AndAlso .oMesh.sDeathSeqWav.GetUpperBound(0) > -1 Then
                Dim lVal As Int32 = CInt(Rnd() * .oMesh.sDeathSeqWav.GetUpperBound(0))
                If goSound Is Nothing = False Then 'AndAlso (.OwnerID = glPlayerID OrElse (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.GetPlayerRelScore(.OwnerID) < elRelTypes.eWar)) Then
                    goSound.StartSound(.oMesh.sDeathSeqWav(lVal), False, SoundMgr.SoundUsage.eDeathExplosions, vecLoc, New Vector3(0, 0, 0))
                    If goSound.MusicOn = True Then goSound.IncrementExcitementLevel(100)
                End If
            End If
        End With
    End Sub

    Public Sub ClearAll()
        mlSeqUB = -1
        Erase moSeq
        Erase mySeqUsed
    End Sub

    Public Sub RenderSequences(ByVal bUpdateNoRender As Boolean)

        'Dim yMapWrapSituation As Byte = 0           '0 = no map wrap, 1 = left edge, 2 = right edge
        'Dim lLocXMapWrapCheck As Int32 = 0
        'Dim lTmpMapWrapVal As Int32 = 0

        'If goCurrentEnvir Is Nothing = False Then
        '    Try
        '        lTmpMapWrapVal = Math.Min((goCurrentEnvir.lMaxXPos * 2) \ 3, muSettings.EntityClipPlane)
        '        If goCamera.mlCameraX < goCurrentEnvir.lMinXPos + lTmpMapWrapVal Then
        '            yMapWrapSituation = 1
        '            lLocXMapWrapCheck = goCurrentEnvir.lMaxXPos - lTmpMapWrapVal
        '        ElseIf goCamera.mlCameraX > goCurrentEnvir.lMaxXPos - lTmpMapWrapVal Then
        '            yMapWrapSituation = 2
        '            lLocXMapWrapCheck = goCurrentEnvir.lMinXPos + lTmpMapWrapVal
        '        End If
        '    Catch
        '        'do nothing
        '    End Try
        'End If

        Try
            For X As Int32 = 0 To mlSeqUB
                If mySeqUsed(X) <> 0 Then
                    'moSeq(X).UpdateMapWrapSituation(yMapWrapSituation, lLocXMapWrapCheck)
                    moSeq(X).Render(bUpdateNoRender)
                    If moSeq(X).bSequenceEnded = True Then
                        mySeqUsed(X) = 0
                        moSeq(X) = Nothing
                    End If
                End If
            Next X
        Catch
            'do nothing
        End Try
    End Sub

End Class

Public Class DeathSequence
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Private Const ml_MINOR_EXP_PCNT As Int32 = 200
    Private Const ml_FINALE_EXP_PCNT As Int32 = 500

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

    Private Structure MinorExplosion
        'Where this minor explosion is located
        Public mvecEmitter As Vector3
        Public vecOffset As Vector3
        Private moParticles() As Particle
        Private mlPrevFrame As Int32

        Public mbEmitterStopping As Boolean
        Public EmitterStopped As Boolean

        'REFERENCE To the Actual array!!!
        Private moPoints() As CustomVertex.PositionColoredTextured
        Private mlPointStartIndex As Int32      'Tells us where to put the points

        Public vecFlashPoint As Vector3
        Public fFlashAlpha As Single

        Public mlParticleUB As Int32

        Public bInitialized As Boolean
        Private mlSmallestSize As Int32
        Private mlHalfSmallestSize As Int32

        Public Sub Initialize(ByVal vecLoc As Vector3, ByVal lStartIndex As Int32, ByRef poPoints() As CustomVertex.PositionColoredTextured, ByVal lSmallestSizeVal As Int32, ByVal vecAdjLoc As Vector3)
            mlSmallestSize = lSmallestSizeVal \ 4
            If mlSmallestSize > 70 Then mlSmallestSize = 70
            mlHalfSmallestSize = mlSmallestSize \ 2

            mbEmitterStopping = False
            EmitterStopped = False
            mlPointStartIndex = lStartIndex
            moPoints = poPoints

            Dim X As Int32

            vecOffset = vecLoc
            mvecEmitter = vecAdjLoc

            mlParticleUB = ml_MINOR_EXP_PCNT - 1
            ReDim moParticles(mlParticleUB)

            For X = 0 To mlParticleUB
                ResetParticle(X)
            Next X

            vecFlashPoint = New Vector3(vecAdjLoc.X, vecAdjLoc.Y, vecAdjLoc.Z)
            fFlashAlpha = 255

            bInitialized = True
        End Sub

        Private Sub ResetParticle(ByVal lIndex As Int32)
            Dim fX As Single
            Dim fY As Single
            Dim fZ As Single
            Dim fXS As Single
            Dim fYS As Single
            Dim fZS As Single
            Dim fXA As Single
            Dim fYA As Single
            Dim fZA As Single

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
                fX = mvecEmitter.X + (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fY = mvecEmitter.Y + (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fZ = mvecEmitter.Z + (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fXS = (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fYS = (Rnd() * mlSmallestSize - mlHalfSmallestSize)
                fZS = (Rnd() * mlSmallestSize - mlHalfSmallestSize)

                fXA = Rnd() * 0.2F : fYA = Rnd() * 0.2F : fZA = Rnd() * 0.2F
                If fXS > 0 Then fXA = -fXA
                If fYS > 0 Then fYA = -fYA
                If fZS > 0 Then fZA = -fZA

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 200 + (Rnd() * 50), 90 + (Rnd() * 30), 32 + (Rnd() * 32), 255)
                moParticles(lIndex).fAChg = -((Rnd() * 16) + 8)
            End If


        End Sub

        Public Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Exit Sub

            If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    If .ParticleActive = True Then
                        .Update(fElapsed)
                        If .mfA <= 0 Then
                            ResetParticle(X)
                        End If

                        moPoints(X + mlPointStartIndex).Color = .ParticleColor.ToArgb
                        moPoints(X + mlPointStartIndex).Position = .vecLoc
                    End If
                End With
            Next X

            fFlashAlpha -= 15 * fElapsed
            If fFlashAlpha < 0 Then fFlashAlpha = 0

            If fFlashAlpha = 0 AndAlso mbEmitterStopping = False Then
                vecFlashPoint = New Vector3(mvecEmitter.X + (Rnd() * 50 - 25), mvecEmitter.Y + (Rnd() * 50 - 25), mvecEmitter.Z + (Rnd() * 50 - 25))
                fFlashAlpha = 255
            End If
        End Sub
    End Structure

    Private Structure FinaleExplosion
        Public mvecEmitter As Vector3

        Private moParticles() As Particle
        Private mlPrevFrame As Int32
        Public mbEmitterStopping As Boolean

        Public moPoints() As CustomVertex.PositionColoredTextured
        Public mlParticleUB As Int32

        Public EmitterStopped As Boolean

        Public bInitialized As Boolean

        Private mfPraxisAlpha As Single
        Private mfPraxisRed As Single
        Private mfPraxisGreen As Single
        Private mfPraxisBlue As Single
        Public mfPraxisScale As Single
        Public moPraxisColor As System.Drawing.Color

        Private mfMaxPraxisSize As Single

        Public Event FinaleComplete()

        Public Sub Initialize(ByVal vecLoc As Vector3, ByVal lMaxSize As Int32)
            mfMaxPraxisSize = lMaxSize / 30.0F
            mbEmitterStopping = False
            EmitterStopped = False

            Dim X As Int32

            mvecEmitter = vecLoc

            mlParticleUB = ml_FINALE_EXP_PCNT - 1
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB + 1)

            For X = 0 To mlParticleUB
                ResetParticle(X)
            Next X

            If Rnd() * 100 < 50 Then mfPraxisAlpha = 0.0F Else mfPraxisAlpha = 255.0F
            mfPraxisRed = 255.0F
            mfPraxisGreen = 255.0F
            mfPraxisBlue = 255.0F
            mfPraxisScale = 0.0F

            bInitialized = True
            mbEmitterStopping = True
        End Sub

        Private Sub ResetParticle(ByVal lIndex As Int32)
            Dim fX As Single
            Dim fY As Single
            Dim fZ As Single
            Dim fXS As Single
            Dim fYS As Single
            Dim fZS As Single
            Dim fXA As Single
            Dim fYA As Single
            Dim fZA As Single

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

                If EmitterStopped = True Then RaiseEvent FinaleComplete()
            Else
                fX = mvecEmitter.X + (Rnd() * 10 - 5)  '(GetNxtRnd() - 0.5)
                fY = mvecEmitter.Y + (Rnd() * 10 - 5) '(GetNxtRnd() - 0.5)
                fZ = mvecEmitter.Z + (Rnd() * 10 - 5) ' (GetNxtRnd() - 0.5)
                fXS = (Rnd() * 10) - 5
                fYS = (Rnd() * 10) - 5
                fZS = (Rnd() * 10) - 5

                fXA = Rnd() * 6 - 3 : fYA = Rnd() * 6 - 3 : fZA = Rnd() * 6 - 3
                If fXS > 0 Then fXA = -fXA
                If fYS > 0 Then fYA = -fYA
                If fZS > 0 Then fZA = -fZA

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 200 + (Rnd() * 50), 90 + (Rnd() * 30), 32 + (Rnd() * 32), 255)
                moParticles(lIndex).fAChg = -(Rnd() + 3.0F)
            End If


        End Sub

        Public Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Return

            If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            mlPrevFrame = timeGetTime

            'Update the praxis
            'TODO: Make the Praxis effect dependant on the size of the entity exploding
            mfPraxisAlpha -= (10.0F * fElapsed)
            mfPraxisRed -= (0.001F * fElapsed)
            mfPraxisGreen -= (0.5F * fElapsed)
            mfPraxisBlue -= (1.0F * fElapsed)


            mfPraxisScale = (1.0F - (mfPraxisAlpha / 255.0F)) * mfMaxPraxisSize
            'mfPraxisScale += (0.7F * fElapsed)
            If mfPraxisAlpha < 0 Then mfPraxisAlpha = 0
            If mfPraxisRed < 0 Then mfPraxisRed = 0
            If mfPraxisGreen < 0 Then mfPraxisGreen = 0
            If mfPraxisBlue < 0 Then mfPraxisBlue = 0
            moPraxisColor = Color.FromArgb(CInt(mfPraxisAlpha), CInt(mfPraxisRed), CInt(mfPraxisGreen), CInt(mfPraxisBlue))

            If mfPraxisAlpha < 0 Then mfPraxisAlpha = 0
            If mfPraxisRed < 200 Then mfPraxisRed = 200
            If mfPraxisGreen < 120 Then mfPraxisGreen = 120
            If mfPraxisBlue < 0 Then mfPraxisBlue = 0

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    If .ParticleActive = True Then
                        .Update(fElapsed)
                        If .mfA <= 0 Then
                            ResetParticle(X)
                        End If

                        moPoints(X).Color = .ParticleColor.ToArgb
                        moPoints(X).Position = .vecLoc
                    End If
                End With
            Next X

        End Sub

    End Structure

    'Particle texture...
    Public Shared moTex As Texture
    Public Shared moPraxisTex As Texture
    Public Shared moSprite As Sprite
    Private moLoc As Vector3
    Private moSize As Vector3

    'Helper for half of moSize
    Private moHalfSize As Vector3

    'Based on parameter in New routine
    Private mbDoFinalExplosion As Boolean = False

    'The individual explosions
    Private moExp() As MinorExplosion
    'That lead up to the finale
    Private moFinale As FinaleExplosion

    'Last time an explosion was added
    Private mlLastExpAdd As Int32 = 0

    'Indicates that this sequence has ended in the event that any subsequent calls are made to this sequence after termination
    Public bSequenceEnded As Boolean = False

    'The final explosion causes a bright light, this manages that bright light
    Private mlFlashPointAlpha As Int32 = 0

    'The actual point list used for all children
    Private moPoints() As CustomVertex.PositionColoredTextured
    Private mlPointUB As Int32 = -1

    Private moFlashPoints() As CustomVertex.PositionColoredTextured
    Private mlCurrentFlashPointUB As Int32 = -1

    Private mlCurrentMaxPointUB As Int32 = -1

    Private mbRenderExplosions As Boolean = True
    Private mlEntityIndex As Int32 = -1

    Private miLocAngle As Int16
    Private miLocYaw As Int16

    Private moEntity As BaseEntity = Nothing

    Public Sub New(ByVal oLoc As Vector3, ByVal lExpCnt As Int32, ByVal oSize As Vector3, ByVal bDoFinalExp As Boolean, ByVal lEntityIndex As Int32, ByVal iLocAngle As Int16, ByVal iLocYaw As Int16, ByVal oEntity As BaseEntity)
        moEntity = oEntity
        mlEntityIndex = lEntityIndex
        miLocAngle = iLocAngle
        miLocYaw = iLocYaw

        If moTex Is Nothing OrElse moTex.Disposed = True Then moTex = goResMgr.GetTexture("Flare2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak")
        If moPraxisTex Is Nothing OrElse moPraxisTex.Disposed = True Then moPraxisTex = goResMgr.GetTexture("Praxis.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "")

        moLoc = oLoc
        moSize = oSize
        moHalfSize = New Vector3(oSize.X / 2.0F, oSize.Y / 2.0F, oSize.Z / 2.0F)
        mbDoFinalExplosion = bDoFinalExp

        'Ok, our points
        ReDim moPoints(lExpCnt * ml_MINOR_EXP_PCNT - 1)
        ReDim moFlashPoints(lExpCnt - 1)
        ReDim moExp(lExpCnt - 1)
        InitializeExplosion(0)

    End Sub

    Private Sub InitializeExplosion(ByVal lIdx As Int32)
        If moExp Is Nothing Then Return
        If lIdx > moExp.GetUpperBound(0) OrElse lIdx < 0 Then Return

        Dim fX As Single = (Rnd() * moSize.X) - moHalfSize.X
        Dim fY As Single = (Rnd() * moSize.Y) - moHalfSize.Y
        Dim fZ As Single = (Rnd() * moSize.Z) - moHalfSize.Z

        RotateVals(fX, fY, fZ)

        Dim lSmallest As Int32 = CInt(Math.Min(Math.Min(moSize.X, moSize.Y), moSize.Z))
        Dim lLargest As Int32 = CInt(Math.Max(Math.Max(moSize.X, moSize.Y), moSize.Z))

        'Dim vecTmp As Vector3 = New Vector3(moLoc.X + fX, moLoc.Y + fY, moLoc.Z + fZ)
        Dim vecTmp As Vector3 = New Vector3(fX, fY, fZ)
        Dim vecAdj As Vector3 = New Vector3(moLoc.X + fX, moLoc.Y + fY, moLoc.Z + fZ)

        moExp(lIdx).Initialize(vecTmp, lIdx * ml_MINOR_EXP_PCNT, moPoints, lSmallest, vecAdj)

        If goExplMgr Is Nothing = False Then
            'Dim fSize As Single = Math.Max(Math.Max(moHalfSize.X, moHalfSize.Y), moHalfSize.Z)
            Dim lDiff As Int32 = lLargest - lSmallest
            lDiff \= 2
            lDiff += lSmallest
            goExplMgr.Add(vecAdj, Rnd() * 360, Rnd() * 4, CInt(Rnd() * 2), lDiff, 0, Color.White, System.Drawing.Color.FromArgb(255, 64, 64, 64), 30, lDiff * 2, True)
            'Dim lSize As Int32 = lSmallest '\ 2
            'Dim lHalfDiff As Int32 = lSmallest \ 2
            'Dim vecTemp2 As Vector3 = New Vector3(vecTmp.X + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Y + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Z + (lSmallest * GetNxtRnd()) - lHalfDiff)
            'goExplMgr.Add(vecTemp2, GetNxtRnd() * 360, GetNxtRnd() * 4, CInt(GetNxtRnd() * 2), lSize, 0, Color.White, System.Drawing.Color.FromArgb(255, 64, 64, 64), 30)
            'vecTemp2 = New Vector3(vecTmp.X + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Y + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Z + (lSmallest * GetNxtRnd()) - lHalfDiff)
            'goExplMgr.Add(vecTemp2, GetNxtRnd() * 360, GetNxtRnd() * 4, CInt(GetNxtRnd() * 2), lSize, 0, Color.White, System.Drawing.Color.FromArgb(255, 64, 64, 64), 30)
            'vecTemp2 = New Vector3(vecTmp.X + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Y + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Z + (lSmallest * GetNxtRnd()) - lHalfDiff)
            'goExplMgr.Add(vecTemp2, GetNxtRnd() * 360, GetNxtRnd() * 4, CInt(GetNxtRnd() * 2), lSize, 0, Color.White, System.Drawing.Color.FromArgb(255, 64, 64, 64), 30)
            'vecTemp2 = New Vector3(vecTmp.X + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Y + (lSmallest * GetNxtRnd()) - lHalfDiff, vecTmp.Z + (lSmallest * GetNxtRnd()) - lHalfDiff)
            'goExplMgr.Add(vecTemp2, GetNxtRnd() * 360, GetNxtRnd() * 4, CInt(GetNxtRnd() * 2), lSize, 0, Color.White, System.Drawing.Color.FromArgb(255, 64, 64, 64), 30)

        End If

        mlCurrentMaxPointUB = ((lIdx + 1) * ml_MINOR_EXP_PCNT) - 1
        mlCurrentFlashPointUB = lIdx
    End Sub

    Public Sub Render(ByVal bUpdateNoRender As Boolean)
        If bSequenceEnded = True Then Return
        Dim moDevice As Device = GFXEngine.moDevice
        If moExp Is Nothing = False AndAlso mbRenderExplosions = True Then

            If moEntity Is Nothing = False Then
                moLoc = New Vector3(moEntity.LocX, moEntity.LocY, moEntity.LocZ)
            End If

            mbRenderExplosions = False
            For X As Int32 = 0 To moExp.GetUpperBound(0)
                If moExp(X).bInitialized = True Then
                    moExp(X).mvecEmitter = moLoc + moExp(X).vecOffset
                    moExp(X).Update()
                    If moExp(X).EmitterStopped = False Then mbRenderExplosions = True
                ElseIf timeGetTime - mlLastExpAdd > 75 Then '50
                    InitializeExplosion(X)
                    mlLastExpAdd = timeGetTime
                    Exit For
                End If
            Next X

            Dim oEnvir As BaseEnvironment
            oEnvir = goCurrentEnvir

            If mlEntityIndex <> -1 Then
                If moExp(moExp.GetUpperBound(0)).bInitialized = True Then
                    If oEnvir Is Nothing = False Then
                        If oEnvir.lEntityUB >= mlEntityIndex AndAlso mlEntityIndex > -1 Then
                            oEnvir.oEntity(mlEntityIndex).ClearParticleFX()

                            'Do our screen shake here...
                            If glCurrentEnvirView = CurrentView.ePlanetView OrElse glCurrentEnvirView = CurrentView.eSystemView Then
                                Try
                                    With oEnvir.oEntity(mlEntityIndex)
                                        Dim fTX As Single
                                        Dim fTZ As Single
                                        Dim fDist As Single
                                        fTX = .LocX - goCamera.mlCameraX
                                        fTZ = .LocZ - goCamera.mlCameraZ
                                        fTX *= fTX
                                        fTZ *= fTZ
                                        fDist = CSng(Math.Sqrt(fTX + fTZ))
                                        If fDist < 10000.0F Then
                                            fDist /= 10000.0F
                                            If .oUnitDef.HullSize > 0 Then fDist *= .oUnitDef.HullSize
                                            goCamera.ScreenShake(1000 * fDist)
                                        End If
                                    End With
                                Catch
                                    'do nothing
                                End Try
                            End If

                            oEnvir.lEntityIdx(mlEntityIndex) = -1
                            oEnvir.oEntity(mlEntityIndex) = Nothing
                        End If
                    Else : LoginScreen.BSDestroyed += 1
                    End If
                    mlEntityIndex = -1
                End If

                If mlEntityIndex = -1 Then

                    For X As Int32 = 0 To moExp.GetUpperBound(0)
                        moExp(X).mbEmitterStopping = True
                    Next

                    'Our final flash, regardless of the big final explosion
                    mlFlashPointAlpha = 255

                    If mbDoFinalExplosion = True Then
                        Dim lMax As Int32 = CInt(Math.Max(Math.Max(moSize.X, moSize.Y), moSize.Z))
                        moFinale.Initialize(moLoc, lMax)
                        AddHandler moFinale.FinaleComplete, AddressOf BigFinale_Ended
                    End If
                End If
            End If

        End If

        If mbRenderExplosions = False AndAlso mbDoFinalExplosion = False Then
            BigFinale_Ended()
            Return
        End If

        If bUpdateNoRender = False Then

            'Now, do our rendering of everything...
            With moDevice
                .Transform.World = Matrix.Identity

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

                .RenderState.PointSpriteEnable = True
                .RenderState.PointScaleEnable = True
                .RenderState.PointSize = 64
                .RenderState.PointScaleA = 0
                .RenderState.PointScaleB = 0
                .RenderState.PointScaleC = 1
                .RenderState.SourceBlend = Blend.SourceAlpha
                .RenderState.DestinationBlend = Blend.One
                .RenderState.AlphaBlendEnable = True

                .RenderState.ZBufferWriteEnable = False

                .RenderState.Lighting = False

                .VertexFormat = CustomVertex.PositionColoredTextured.Format
                .SetTexture(0, moTex)

                If mbRenderExplosions = True Then
                    'Render the Minor Explosion Points in one call...
                    .DrawUserPrimitives(PrimitiveType.PointList, mlCurrentMaxPointUB + 1, moPoints)

                    'Prepare to render flashpoints
                    .RenderState.PointSize = 512

                    'set up our flashpoints
                    For X As Int32 = 0 To mlCurrentFlashPointUB
                        moFlashPoints(X).Color = System.Drawing.Color.FromArgb(CInt(moExp(X).fFlashAlpha), 255, 255, 255).ToArgb
                        moFlashPoints(X).Position = moExp(X).vecFlashPoint
                    Next X

                    'Render the MinorExplosion's FlashPoints
                    .DrawUserPrimitives(PrimitiveType.PointList, mlCurrentFlashPointUB + 1, moFlashPoints)

                End If

                'Render the Finale Explosion Points in one call...
                If mbDoFinalExplosion = True AndAlso moFinale.bInitialized = True Then
                    .RenderState.PointSize = 128
                    moFinale.Update()
                    .DrawUserPrimitives(PrimitiveType.PointList, moFinale.moPoints.Length, moFinale.moPoints)

                    'Then, render the Praxis effect
                    If moFinale.moPraxisColor.A > 0 Then
                        moDevice.Transform.World = Matrix.Multiply(Matrix.Identity, Matrix.Scaling(moFinale.mfPraxisScale, moFinale.mfPraxisScale, moFinale.mfPraxisScale))
                        If moSprite Is Nothing Then
                            Device.IsUsingEventHandlers = False
                            moSprite = New Sprite(moDevice)
                            Device.IsUsingEventHandlers = True
                        End If
                        With moSprite
                            .SetWorldViewLH(moDevice.Transform.World, moDevice.Transform.View)
                            .Begin(SpriteFlags.AlphaBlend Or SpriteFlags.ObjectSpace)
                            .Draw(moPraxisTex, System.Drawing.Rectangle.Empty, New Vector3(64, 64, 0), Vector3.Multiply(Me.moLoc, 1 / moFinale.mfPraxisScale), moFinale.moPraxisColor)
                            .End()
                        End With
                        'Using moPraxis As Sprite = New Sprite(moDevice)
                        '    moPraxis.SetWorldViewLH(moDevice.Transform.World, moDevice.Transform.View)
                        '    moPraxis.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.ObjectSpace)
                        '    moPraxis.Draw(moPraxisTex, System.Drawing.Rectangle.Empty, New Vector3(64, 64, 0), Vector3.Multiply(Me.moLoc, 1 / moFinale.mfPraxisScale), moFinale.moPraxisColor)
                        '    moPraxis.End()
                        'End Using
                    End If
                End If

                'Then, reset our device...
                .RenderState.ZBufferWriteEnable = True
                .RenderState.Lighting = True

                .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                .RenderState.AlphaBlendEnable = True

                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With
        End If

        If mlFlashPointAlpha <> 0 Then
            If bUpdateNoRender = False Then
                If moSprite Is Nothing Then
                    Device.IsUsingEventHandlers = False
                    moSprite = New Sprite(moDevice)
                    Device.IsUsingEventHandlers = True
                End If
                With moSprite
                    Dim oTmpMat As Matrix = Matrix.Identity
                    Dim oTmpVec As Vector3 = New Vector3(moLoc.X / moHalfSize.X, moLoc.Y / moHalfSize.Y, moLoc.Z / moHalfSize.Z)
                    oTmpMat.Multiply(Matrix.Translation(oTmpVec))
                    oTmpMat.Multiply(Matrix.Scaling(moHalfSize))
                    moDevice.Transform.World = oTmpMat
                    moDevice.RenderState.ZBufferWriteEnable = False

                    .SetWorldViewLH(oTmpMat, moDevice.Transform.View)
                    .Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)
                    .Draw(moTex, System.Drawing.Rectangle.Empty, New Vector3(16, 16, 0), New Vector3(0, 0, 0), Color.FromArgb(mlFlashPointAlpha, 255, 255, 255))
                    .End()

                    moDevice.RenderState.ZBufferWriteEnable = True
                    moDevice.Transform.World = Matrix.Identity
                End With
                'Using oSpr As New Sprite(moDevice)
                '    Dim oTmpMat As Matrix = Matrix.Identity
                '    Dim oTmpVec As Vector3 = New Vector3(moLoc.X / moHalfSize.X, moLoc.Y / moHalfSize.Y, moLoc.Z / moHalfSize.Z)
                '    oTmpMat.Multiply(Matrix.Translation(oTmpVec))
                '    oTmpMat.Multiply(Matrix.Scaling(moHalfSize))
                '    moDevice.Transform.World = oTmpMat
                '    moDevice.RenderState.ZBufferWriteEnable = False
                '    oSpr.SetWorldViewLH(oTmpMat, moDevice.Transform.View)
                '    oSpr.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)
                '    oSpr.Draw(moTex, System.Drawing.Rectangle.Empty, New Vector3(16, 16, 0), New Vector3(0, 0, 0), Color.FromArgb(mlFlashPointAlpha, 255, 255, 255))
                '    oSpr.End()
                '    oSpr.Dispose()
                '    moDevice.RenderState.ZBufferWriteEnable = True
                '    moDevice.Transform.World = Matrix.Identity
                'End Using
            End If

            mlFlashPointAlpha -= 10
            If mlFlashPointAlpha < 0 Then mlFlashPointAlpha = 0
        End If
    End Sub

    Private Sub BigFinale_Ended()
        bSequenceEnded = True
    End Sub

    Private Sub RotateVals(ByRef fXLoc As Single, ByRef fYLoc As Single, ByRef fZLoc As Single)
        Dim fDX As Single
        Dim fDZ As Single

        Const gdRadPerDegree As Single = Math.PI / 180.0F

        Dim fRads As Single = -((miLocYaw / 10.0F)) * gdRadPerDegree
        Dim fCosR As Single = CSng(Math.Cos(fRads))
        Dim fSinR As Single = CSng(Math.Sin(fRads))

        'Yaw...
        fDX = fXLoc
        fDZ = fYLoc
        fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
        fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

        'Now set up for standard rotation...
        fRads = ((miLocAngle / 10.0F) - 90) * gdRadPerDegree
        fCosR = CSng(Math.Cos(fRads))
        fSinR = CSng(Math.Sin(fRads))

        'Loc
        fDX = fXLoc
        fDZ = fZLoc
        fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
        fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

    End Sub

    Public Shared Sub ClearShared()
        ' DeathSequence.moTex, moPraxisTex, moSprite (shared)
        If moTex Is Nothing = False Then moTex.Dispose()
        moTex = Nothing
        If moPraxisTex Is Nothing = False Then moPraxisTex.Dispose()
        moPraxisTex = Nothing
        If moSprite Is Nothing = False Then moSprite.Dispose()
        moSprite = Nothing
    End Sub

    'Private myPrevMapWrapSituation As Byte = 0
    'Public Sub UpdateMapWrapSituation(ByVal yMapWrapSituation As Byte, ByVal lLocXMapWrapCheck As Integer)
    '    If yMapWrapSituation <> myPrevMapWrapSituation Then
    '        'Reset our loc
    '        Dim lLocXMod As Int32 = 0
    '        If moLoc.X < goCurrentEnvir.lMinXPos Then
    '            moLoc.X += goCurrentEnvir.lMapWrapAdjustX
    '            lLocXMod += goCurrentEnvir.lMapWrapAdjustX
    '        ElseIf moLoc.X > goCurrentEnvir.lMaxXPos Then
    '            moLoc.X -= goCurrentEnvir.lMapWrapAdjustX
    '            lLocXMod -= goCurrentEnvir.lMapWrapAdjustX
    '        End If

    '        If yMapWrapSituation = 1 Then
    '            'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
    '            If moLoc.X > lLocXMapWrapCheck Then
    '                moLoc.X -= goCurrentEnvir.lMapWrapAdjustX
    '                lLocXMod -= goCurrentEnvir.lMapWrapAdjustX
    '            End If
    '        ElseIf yMapWrapSituation = 2 Then
    '            'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
    '            If moLoc.X < lLocXMapWrapCheck Then
    '                moLoc.X += goCurrentEnvir.lMapWrapAdjustX
    '                lLocXMod += goCurrentEnvir.lMapWrapAdjustX
    '            End If
    '        End If

    '        If lLocXMod <> 0 Then
    '            If moPoints Is Nothing = False Then
    '                For X As Int32 = 0 To mlPointUB
    '                    moPoints(X).X += lLocXMod
    '                Next X
    '            End If
    '            If moExp Is Nothing = False Then
    '                For X As Int32 = 0 To moExp.GetUpperBound(0)
    '                    moExp(X).mvecEmitter.X += lLocXMod
    '                Next
    '            End If
    '        End If
    '    End If
    '    myPrevMapWrapSituation = yMapWrapSituation
    'End Sub
End Class