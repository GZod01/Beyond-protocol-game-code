Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Namespace PlanetFX
    Public Class Particle
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

    End Class

    Public Class ParticleEngine

        Public Enum EmitterType As Integer
            eVolcano = 0
            eSlowSteam = 1
            eAcidMist = 2
            eLavaColumn = 3
        End Enum

        Private moChildren() As FXEmitter
        Private myChildUsed() As Byte
        Public mlChildrenUB As Int32 = -1

        'The standard texture to use for emitters using this engine object
        Public Shared moTex As Texture

        Private moTerrain As TerrainClass

        Private mlPreviousPFX As Int32
        'Private myPrevMapWrapSituation As Byte

        Public Sub New(ByRef oParent As TerrainClass)
            moTerrain = oParent
            ReDim myChildUsed(0)
            myChildUsed(0) = 0
            mlPreviousPFX = muSettings.PlanetFXParticles
        End Sub

        Public Sub Render(ByVal bUpdateNoRender As Boolean) 
            Dim X As Int32

            If muSettings.PlanetFXParticles = 0 Then Return

            If moTex Is Nothing OrElse moTex.Disposed = True Then
				moTex = goResMgr.GetTexture("ExpParticle2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "Misc.pak")
            End If

            If mlChildrenUB = -1 Then Return

            If mlPreviousPFX <> muSettings.PlanetFXParticles Then
                For X = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 Then
                        moChildren(X).ResetParticleCount()
                    End If
                Next X
                mlPreviousPFX = muSettings.PlanetFXParticles
            End If

            Dim moDevice As Device = GFXEngine.moDevice

            'Dim yMapWrapSituation As Byte = 0           '0 = no map wrap, 1 = left edge, 2 = right edge
            'Dim lLocXMapWrapCheck As Int32 = 0
            'Dim lTmpMapWrapVal As Int32 = Math.Min((goCurrentEnvir.lMaxXPos * 2) \ 3, muSettings.EntityClipPlane)
            'If goCamera.mlCameraX < goCurrentEnvir.lMinXPos + lTmpMapWrapVal Then
            '    yMapWrapSituation = 1
            '    lLocXMapWrapCheck = goCurrentEnvir.lMaxXPos - lTmpMapWrapVal
            'ElseIf goCamera.mlCameraX > goCurrentEnvir.lMaxXPos - lTmpMapWrapVal Then
            '    yMapWrapSituation = 2
            '    lLocXMapWrapCheck = goCurrentEnvir.lMinXPos + lTmpMapWrapVal
            'End If

            With moDevice
                If bUpdateNoRender = False Then
                    'And now render them... first set up our device for renders...
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
                        If .DeviceCaps.MaxPointSize > 128 Then
                            .RenderState.PointSize = 128
                        Else : .RenderState.PointSize = .DeviceCaps.MaxPointSize
                        End If
                        .RenderState.PointScaleA = 0
                        .RenderState.PointScaleB = 0
                        .RenderState.PointScaleC = 0.1F
                    End If

                    .RenderState.SourceBlend = Blend.SourceAlpha
                    .RenderState.DestinationBlend = Blend.One
                    .RenderState.AlphaBlendEnable = True

                    .RenderState.ZBufferWriteEnable = False

                    .RenderState.Lighting = False

                    .VertexFormat = CustomVertex.PositionColoredTextured.Format
                    '.VertexFormat = moPoints(0).Format
                    .SetTexture(0, moTex)

                    'glPlanetFXRendered = 0
                    For X = 0 To mlChildrenUB
                        If myChildUsed(X) > 0 Then
                            'If yMapWrapSituation <> myPrevMapWrapSituation Then moChildren(X).UpdateMapWrapSituation(yMapWrapSituation, lLocXMapWrapCheck)
                            If goCamera.CullObject(moChildren(X).GetCullBox) = False Then
                                'Now, check for whether the emitter can be seen
                                moChildren(X).Update()
                                If moChildren(X).EmitterStopped = True OrElse moChildren(X).mlParticleUB = -1 Then
                                    myChildUsed(X) = 0
                                Else
                                    .DrawUserPrimitives(PrimitiveType.PointList, moChildren(X).mlParticleUB + 1, moChildren(X).moPoints)
                                    'glPlanetFXRendered += 1
                                End If
                            End If
                        End If
                    Next X

                    'Then, reset our device...
                    .RenderState.ZBufferWriteEnable = True
                    .RenderState.Lighting = True
                    .RenderState.PointSpriteEnable = False
                    .RenderState.PointScaleEnable = False

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
                End If
            End With

            'myPrevMapWrapSituation = yMapWrapSituation
        End Sub

        Public Function AddEmitter(ByVal iType As EmitterType, ByVal oLoc As Vector3, ByVal lPCnt As Int32) As Int32
            Dim X As Int32
            Dim lIdx As Int32 = -1
            ' Try
            For X = 0 To mlChildrenUB
                If myChildUsed(X) = 0 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                mlChildrenUB += 1
                ReDim Preserve moChildren(mlChildrenUB)
                ReDim Preserve myChildUsed(mlChildrenUB)
                lIdx = mlChildrenUB
            End If

            Select Case iType
                Case EmitterType.eLavaColumn
                    If lPCnt <= 0 Then lPCnt = 30
                    moChildren(lIdx) = New LavaColumn(oLoc, lPCnt, moTerrain.CellSpacing)
                Case EmitterType.eVolcano
                    If lPCnt <= 0 Then lPCnt = 400
                    moChildren(lIdx) = New VolcanoEmitter(oLoc, moTerrain.CellSpacing, lPCnt, moTerrain.ml_Y_Mult)
            End Select

            myChildUsed(lIdx) = 255
            ' Catch
            ' End Try

            Return lIdx
        End Function

        Public Function AddEmitter(ByVal iType As EmitterType, ByVal oLoc() As Vector3, ByVal lPCnt As Int32) As Int32
            Dim X As Int32
            Dim lIdx As Int32 = -1

            For X = 0 To mlChildrenUB
                If myChildUsed(X) = 0 Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                mlChildrenUB += 1
                ReDim Preserve moChildren(mlChildrenUB)
                ReDim Preserve myChildUsed(mlChildrenUB)
                lIdx = mlChildrenUB
            End If

            Select Case iType
                Case EmitterType.eSlowSteam
                    If lPCnt <= 0 Then lPCnt = 300 
                    moChildren(lIdx) = New SlowSteam(oLoc, moTerrain.CellSpacing, lPCnt)
                Case EmitterType.eAcidMist
                    If lPCnt <= 0 Then lPCnt = 300 
                    moChildren(lIdx) = New AcidMist(oLoc, moTerrain.CellSpacing, lPCnt)
            End Select

            myChildUsed(lIdx) = 255

            Return lIdx
        End Function

        Public Sub StopEmitter(ByVal lIndex As Int32)
            If lIndex > -1 AndAlso lIndex <= mlChildrenUB Then
                If myChildUsed(lIndex) > 0 Then moChildren(lIndex).StopEmitter()
            End If
        End Sub

        Public Sub ClearAllEmitters()
            Erase moChildren
            Erase myChildUsed
            mlChildrenUB = -1
        End Sub
    End Class

    Public MustInherit Class FXEmitter
        Protected Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

        Protected Shared moInvisColor As System.Drawing.Color = System.Drawing.Color.FromArgb(0, 0, 0, 0)

        Protected moParticles() As Particle
        Protected mlPrevFrame As Int32
        Protected mbEmitterStopping As Boolean = False

        Public moPoints() As CustomVertex.PositionColoredTextured
        Public mlParticleUB As Int32
        Public EmitterStopped As Boolean = False

        Protected MustOverride Sub ResetParticle(ByVal lIndex As Int32)
        Public MustOverride Sub Update()
        Public MustOverride Function GetCullBox() As CullBox
        'Public MustOverride Sub UpdateMapWrapSituation(ByVal yMapWrapSituation As Byte, ByVal lLocXMapWrapCheck As Int32)
        Protected myPrevMapWrapSituation As Byte = 0

        Public lEmitterType As ParticleEngine.EmitterType

        Public ParticleCount As Int32

        Private mlOriginalPCnt As Int32

        Protected mbInitialized As Boolean = False

        Public Sub Initialize(ByVal lParticleCnt As Int32)
            Dim X As Int32

            mlOriginalPCnt = lParticleCnt

            Dim fMult As Single = 1.0F
            Select Case muSettings.PlanetFXParticles
                Case 0
                    Return
                Case 1
                    fMult = 0.25F
                Case 2
                    fMult = 0.5F
                Case 3
                    fMult = 0.75F
            End Select
            ParticleCount = CInt(lParticleCnt * fMult)

            mlParticleUB = ParticleCount - 1
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB)

            For X = 0 To mlParticleUB
                moParticles(X) = New Particle()
                ResetParticle(X)
            Next X
            mbInitialized = True
        End Sub

        Public Sub StopEmitter()
            mbEmitterStopping = True
        End Sub

        Public Sub StartEmitter()
            Dim X As Int32
            mbEmitterStopping = False
            For X = 0 To mlParticleUB
                moParticles(X).ParticleActive = True
                ResetParticle(X)
            Next X
        End Sub

        Public Sub ResetParticleCount()
            If mbInitialized = False Then Return
            Dim lTemp As Int32 = mlOriginalPCnt
            Initialize(mlOriginalPCnt)
            mlOriginalPCnt = lTemp
        End Sub
    End Class

    Public Class VolcanoEmitter
        Inherits FXEmitter

        Public vecEmitter As Vector3

        Private mlCellSpacing As Int32
        Private mlFull As Int32
        Private mlHalf As Int32

        Private ml_Y_Mult As Single

        Public Sub New(ByVal oVecLoc As Vector3, ByVal lCellSpacing As Int32, ByVal lParticleCnt As Int32, ByVal y_Mult As Single)
            mlCellSpacing = lCellSpacing
            mlFull = lCellSpacing \ 2
            mlHalf = mlFull \ 2
            ml_Y_Mult = y_Mult
            vecEmitter = oVecLoc
            ParticleCount = lParticleCnt
        End Sub

        Public Sub DisposeMe()
            Erase moParticles
            Erase moPoints
        End Sub

        Protected Overrides Sub ResetParticle(ByVal lIndex As Integer)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lX).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX

                If EmitterStopped = True Then DisposeMe()
            Else
                'moParticles(lIndex).Reset(vecEmitter.X + ((GetNxtRnd() * mlFull) - mlHalf), vecEmitter.Y, vecEmitter.Z + ((GetNxtRnd() * mlFull) - mlHalf), GetNxtRnd() * 6.0F - 3.0F, GetNxtRnd() * 15.0F, GetNxtRnd() * 6.0F - 3.0F, 0, GetNxtRnd() * -0.001F, 0, 128, 128, 128, 255)
                'moParticles(lIndex).fAChg = -(GetNxtRnd() * 6.0F + 4)

                Dim fX As Single = vecEmitter.X + ((Rnd() * mlFull) - mlHalf)
                Dim fY As Single = vecEmitter.Y + ((Rnd() * mlFull) - mlHalf)
                Dim fZ As Single = vecEmitter.Z + ((Rnd() * mlFull) - mlHalf)
                Dim fXS As Single = Rnd() - 0.5F
                Dim fZS As Single = Rnd() - 0.5F
                Dim fYS As Single = (Rnd() * 2.0F) + 3.6F

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, 0, (Rnd() * 0.006F) + 0.001F, 0, 255, 128, Rnd() * 50, (0.8F + (0.2F * Rnd())) * 255)
                moParticles(lIndex).fAChg = -((Rnd() * 0.6F) + 0.2F)

                moParticles(lIndex).fRChg = moParticles(lIndex).fAChg * 2
                moParticles(lIndex).fGChg = moParticles(lIndex).fRChg / 2
                moParticles(lIndex).fBChg = moParticles(lIndex).fGChg / 2

            End If
        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Exit Sub
            If mbInitialized = False Then MyBase.Initialize(ParticleCount)

            If mlPrevFrame = 0 Then fElapsed = 1.0F Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            If fElapsed > 1.0F Then fElapsed = 1.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)

                    If .mfA <= 0 Then
                        ResetParticle(X)
                    Else
                        If .vecLoc.Y > 260 * ml_Y_Mult AndAlso .vecAcc.Y > 0.0F Then
                            .vecAcc.Y *= -3
                        End If

                        If .mfB < 52 Then
                            .mfB = 52
                            .fBChg = 0
                        End If
                        If .mfR < 64 Then
                            .mfR = 64
                            .fRChg = 0
                        End If
                        If .mfG < 52 Then
                            .mfG = 52
                            .fGChg = 0
                        End If
                    End If


                    If EmitterStopped = True Then Exit Sub

                    moPoints(X).Color = .ParticleColor.ToArgb
                    moPoints(X).Position = .vecLoc

                End With
            Next X
        End Sub

        Private mbCullBoxSet As Boolean = False
        Private muCullBox As CullBox
        Public Overrides Function GetCullBox() As CullBox
            If mbCullBoxSet = False Then
                mbCullBoxSet = True
                muCullBox = CullBox.GetCullBox(vecEmitter.X, vecEmitter.Y, vecEmitter.Z, -mlCellSpacing, 0, -mlCellSpacing, mlCellSpacing, 100000, mlCellSpacing)
            End If
            Return muCullBox
        End Function

        'Public Overrides Sub UpdateMapWrapSituation(ByVal yMapWrapSituation As Byte, ByVal lLocXMapWrapCheck As Integer)
        '    If yMapWrapSituation <> myPrevMapWrapSituation Then
        '        mbCullBoxSet = False
        '        'Reset our loc
        '        If vecEmitter.X < goCurrentEnvir.lMinXPos Then
        '            vecEmitter.X += goCurrentEnvir.lMapWrapAdjustX
        '        ElseIf vecEmitter.X > goCurrentEnvir.lMaxXPos Then
        '            vecEmitter.X -= goCurrentEnvir.lMapWrapAdjustX
        '        End If

        '        If yMapWrapSituation = 1 Then
        '            'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
        '            If vecEmitter.X > lLocXMapWrapCheck Then
        '                vecEmitter.X -= goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        ElseIf yMapWrapSituation = 2 Then
        '            'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
        '            If vecEmitter.X < lLocXMapWrapCheck Then
        '                vecEmitter.X += goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        End If
        '    End If
        '    myPrevMapWrapSituation = yMapWrapSituation
        'End Sub
    End Class

    Public Class SlowSteam
        Inherits FXEmitter

        Public vecEmitter() As Vector3

        Private mlFull As Int32
        Private mlHalf As Int32

        Public Sub New(ByVal oVecLoc() As Vector3, ByVal lCellSpacing As Int32, ByVal lParticleCnt As Int32)
            mlFull = lCellSpacing
            mlHalf = mlFull \ 2

            vecEmitter = oVecLoc
            ParticleCount = lParticleCnt
        End Sub

        Public Sub DisposeMe()
            Erase moParticles
            Erase moPoints
        End Sub

        Private mbCullBoxSet As Boolean = False
        Private muCullBox As CullBox
        Public Overrides Function GetCullBox() As CullBox
            If mbCullBoxSet = False Then
                mbCullBoxSet = True

                Dim vecMin As Vector3 = New Vector3(Int32.MaxValue, Int32.MaxValue, Int32.MaxValue)
                Dim vecMax As Vector3 = New Vector3(Int32.MinValue, Int32.MinValue, Int32.MinValue)

                For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
                    With vecEmitter(X)
                        If .X < vecMin.X Then vecMin.X = .X
                        If .X > vecMax.X Then vecMax.X = .X
                        If .Y < vecMin.Y Then vecMin.Y = .Y
                        If .Y > vecMax.Y Then vecMax.Y = .Y
                        If .Z < vecMin.Z Then vecMin.Z = .Z
                        If .Z > vecMax.Z Then vecMax.Z = .Z
                    End With
                Next X

                vecMin.X += -mlFull
                vecMin.Y += -mlFull
                vecMin.Z += -mlFull
                vecMax.X += mlFull
                vecMax.Y += mlFull
                vecMax.Z += mlFull

                Dim vecSize As Vector3 = Vector3.Multiply(Vector3.Subtract(vecMax, vecMin), 0.5F)
                Dim vecCenter As Vector3 = Vector3.Add(vecMin, vecSize)

                muCullBox = CullBox.GetCullBox(vecCenter.X, vecCenter.Y, vecCenter.Z, -vecSize.X, -vecSize.Y, -vecSize.Z, vecSize.X, vecSize.Y, vecSize.Z)
            End If
            Return muCullBox
        End Function

        Protected Overrides Sub ResetParticle(ByVal lIndex As Integer)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lX).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX

                If EmitterStopped = True Then DisposeMe()
            Else
                'moParticles(lIndex).Reset(vecEmitter.X + ((GetNxtRnd() * mlFull) - mlHalf), vecEmitter.Y, vecEmitter.Z + ((GetNxtRnd() * mlFull) - mlHalf), GetNxtRnd() * 6.0F - 3.0F, GetNxtRnd() * 15.0F, GetNxtRnd() * 6.0F - 3.0F, 0, GetNxtRnd() * -0.001F, 0, 128, 128, 128, 255)
                'moParticles(lIndex).fAChg = -(GetNxtRnd() * 6.0F + 4)

                Dim lEmitterIdx As Int32 = CInt(Rnd() * vecEmitter.GetUpperBound(0))

                Dim fX As Single = vecEmitter(lEmitterIdx).X + ((Rnd() * mlFull) - mlHalf)
                Dim fY As Single = vecEmitter(lEmitterIdx).Y + ((Rnd() * mlFull) - mlHalf)
                Dim fZ As Single = vecEmitter(lEmitterIdx).Z + ((Rnd() * mlFull) - mlHalf)
                Dim fXS As Single = Rnd() * 2 - 1
                Dim fZS As Single = Rnd() * 2 - 1
                Dim fYS As Single = (Rnd() * 0.3F) + 0.1F

                'moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, 0, -0.001F, 0, 255, 128, GetNxtRnd() * 50, (0.8F + (0.2F * GetNxtRnd())) * 255)
                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, 0, 0, 0, 128, 128, 128, Rnd() * 16)
                moParticles(lIndex).fAChg = (Rnd() * 0.6F) + 0.2F

            End If
        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Exit Sub
            If mbInitialized = False Then MyBase.Initialize(ParticleCount)

            If mlPrevFrame = 0 Then fElapsed = 1.0F Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            If fElapsed > 1.0F Then fElapsed = 1.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)

                    If .mfA > 128 Then
                        .mfA = 128
                        .fAChg = -(Math.Abs(.fAChg))
                    End If
                    If .mfA <= 0 Then ResetParticle(X)
                    If .vecSpeed.Y < 0 Then
                        .vecSpeed.Y = 0 : .vecAcc.Y = 0
                    End If

                    If EmitterStopped = True Then Exit Sub

                    moPoints(X).Color = .ParticleColor.ToArgb
                    moPoints(X).Position = .vecLoc

                End With
            Next X
        End Sub

        'Public Overrides Sub UpdateMapWrapSituation(ByVal yMapWrapSituation As Byte, ByVal lLocXMapWrapCheck As Integer)
        '    If yMapWrapSituation <> myPrevMapWrapSituation Then
        '        mbCullBoxSet = False
        '        'Reset our locs
        '        For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
        '            If vecEmitter(X).X < goCurrentEnvir.lMinXPos Then
        '                vecEmitter(X).X += goCurrentEnvir.lMapWrapAdjustX
        '            ElseIf vecEmitter(X).X > goCurrentEnvir.lMaxXPos Then
        '                vecEmitter(X).X -= goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        Next X


        '        If yMapWrapSituation = 1 Then
        '            'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
        '            For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
        '                If vecEmitter(X).X > lLocXMapWrapCheck Then
        '                    vecEmitter(X).X -= goCurrentEnvir.lMapWrapAdjustX
        '                End If
        '            Next X
        '        ElseIf yMapWrapSituation = 2 Then
        '            'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
        '            For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
        '                If vecEmitter(X).X < lLocXMapWrapCheck Then
        '                    vecEmitter(X).X += goCurrentEnvir.lMapWrapAdjustX
        '                End If
        '            Next X
        '        End If

        '    End If
        '    myPrevMapWrapSituation = yMapWrapSituation
        'End Sub
    End Class

    Public Class AcidMist
        Inherits FXEmitter

        Public vecEmitter() As Vector3

        Private mlFull As Int32
        Private mlHalf As Int32

        Public Sub New(ByVal oVecLoc() As Vector3, ByVal lCellSpacing As Int32, ByVal lParticleCnt As Int32)
            mlFull = lCellSpacing
            mlHalf = mlFull \ 2
            vecEmitter = oVecLoc
            ParticleCount = lParticleCnt
        End Sub

        Public Sub DisposeMe()
            Erase moParticles
            Erase moPoints
        End Sub

        Private mbCullBoxSet As Boolean = False
        Private muCullBox As CullBox
        Public Overrides Function GetCullBox() As CullBox
            If mbCullBoxSet = False Then
                mbCullBoxSet = True

                Dim vecMin As Vector3 = New Vector3(Int32.MaxValue, Int32.MaxValue, Int32.MaxValue)
                Dim vecMax As Vector3 = New Vector3(Int32.MinValue, Int32.MinValue, Int32.MinValue)

                For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
                    With vecEmitter(X)
                        If .X < vecMin.X Then vecMin.X = .X
                        If .X > vecMax.X Then vecMax.X = .X
                        If .Y < vecMin.Y Then vecMin.Y = .Y
                        If .Y > vecMax.Y Then vecMax.Y = .Y
                        If .Z < vecMin.Z Then vecMin.Z = .Z
                        If .Z > vecMax.Z Then vecMax.Z = .Z
                    End With
                Next X

                vecMin.X += -mlFull
                vecMin.Y += -mlFull
                vecMin.Z += -mlFull
                vecMax.X += mlFull
                vecMax.Y += mlFull
                vecMax.Z += mlFull

                Dim vecSize As Vector3 = Vector3.Multiply(Vector3.Subtract(vecMax, vecMin), 0.5F)
                Dim vecCenter As Vector3 = Vector3.Add(vecMin, vecSize)

                muCullBox = CullBox.GetCullBox(vecCenter.X, vecCenter.Y, vecCenter.Z, -vecSize.X, -vecSize.Y, -vecSize.Z, vecSize.X, vecSize.Y, vecSize.Z)
            End If
            Return muCullBox
        End Function

        Protected Overrides Sub ResetParticle(ByVal lIndex As Integer)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lX).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX

                If EmitterStopped = True Then DisposeMe()
            Else
                Dim lEmitterIdx As Int32 = CInt(Rnd() * vecEmitter.GetUpperBound(0))

                Dim fX As Single = vecEmitter(lEmitterIdx).X + ((Rnd() * mlFull) - mlHalf)
                Dim fY As Single = vecEmitter(lEmitterIdx).Y + ((Rnd() * mlFull)) ' - mlHalf)
                Dim fZ As Single = vecEmitter(lEmitterIdx).Z + ((Rnd() * mlFull) - mlHalf)
                Dim fXS As Single = Rnd() * 2 - 1
                Dim fZS As Single = Rnd() * 2 - 1
                Dim fYS As Single = Rnd()
                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, 0, 0, 0, 24, 255, 24, Rnd() * 16)
                moParticles(lIndex).fAChg = ((Rnd() * 5.0F) + 2.0F)

            End If
        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If EmitterStopped = True Then Exit Sub
            If mbInitialized = False Then MyBase.Initialize(ParticleCount)

            If mlPrevFrame = 0 Then fElapsed = 1.0F Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            If fElapsed > 1.0F Then fElapsed = 1.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)

                    If .mfA <= 0 Then ResetParticle(X)
                    If .mfA > 128 Then
                        .mfA = 128
                        .fAChg = -(Math.Abs(.fAChg))
                    End If
                    If .vecSpeed.Y < 0 Then
                        .vecSpeed.Y = 0 : .vecAcc.Y = 0
                    End If

                    If EmitterStopped = True Then Exit Sub

                    moPoints(X).Color = .ParticleColor.ToArgb
                    moPoints(X).Position = .vecLoc

                End With
            Next X
        End Sub

        'Public Overrides Sub UpdateMapWrapSituation(ByVal yMapWrapSituation As Byte, ByVal lLocXMapWrapCheck As Integer)
        '    If yMapWrapSituation <> myPrevMapWrapSituation Then
        '        mbCullBoxSet = False
        '        'Reset our locs
        '        For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
        '            If vecEmitter(X).X < goCurrentEnvir.lMinXPos Then
        '                vecEmitter(X).X += goCurrentEnvir.lMapWrapAdjustX
        '            ElseIf vecEmitter(X).X > goCurrentEnvir.lMaxXPos Then
        '                vecEmitter(X).X -= goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        Next X


        '        If yMapWrapSituation = 1 Then
        '            'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
        '            For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
        '                If vecEmitter(X).X > lLocXMapWrapCheck Then
        '                    vecEmitter(X).X -= goCurrentEnvir.lMapWrapAdjustX
        '                End If
        '            Next X
        '        ElseIf yMapWrapSituation = 2 Then
        '            'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
        '            For X As Int32 = 0 To vecEmitter.GetUpperBound(0)
        '                If vecEmitter(X).X < lLocXMapWrapCheck Then
        '                    vecEmitter(X).X += goCurrentEnvir.lMapWrapAdjustX
        '                End If
        '            Next X
        '        End If

        '    End If
        '    myPrevMapWrapSituation = yMapWrapSituation
        'End Sub
    End Class

    Public Class LavaColumn
        Inherits FXEmitter

        Public vecEmitter As Vector3

        Private mlFull As Int32
        Private mlHalf As Int32

        Public Sub New(ByVal oVecLoc As Vector3, ByVal lParticleCnt As Int32, ByVal lCellSpacing As Int32)
            mlFull = lCellSpacing
            mlHalf = mlFull \ 2
            vecEmitter = oVecLoc
            ParticleCount = lParticleCnt
        End Sub

        Public Sub DisposeMe()
            Erase moParticles
            Erase moPoints
        End Sub

        Private mbCullBoxSet As Boolean = False
        Private muCullBox As CullBox
        Public Overrides Function GetCullBox() As CullBox
            If mbCullBoxSet = False Then
                mbCullBoxSet = True
                muCullBox = CullBox.GetCullBox(vecEmitter.X, vecEmitter.Y, vecEmitter.Z, -mlFull, 0, -mlFull, mlFull, 100000, mlFull)
            End If
            Return muCullBox
        End Function

        Protected Overrides Sub ResetParticle(ByVal lIndex As Integer)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                moParticles(lIndex).mfA = 0
                moParticles(lIndex).ParticleActive = False

                EmitterStopped = True
                For lX = 0 To mlParticleUB
                    If moParticles(lX).ParticleActive = True Then
                        EmitterStopped = False
                        Exit For
                    End If
                Next lX

                If EmitterStopped = True Then DisposeMe()
            Else
                Dim fX As Single = vecEmitter.X
                Dim fY As Single = vecEmitter.Y
                Dim fZ As Single = vecEmitter.Z
                Dim fXS As Single = (Rnd() * 6) - 3
                Dim fZS As Single = (Rnd() * 6) - 3
                Dim fYS As Single = Rnd() * 20 + 10

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, 0, 0, 0, 255, 128, Rnd() * 50, (0.8F + (0.2F * Rnd())) * 255)
                moParticles(lIndex).fAChg = -((Rnd() * 5.0F) + 2.0F)
                moParticles(lIndex).fRChg = (0.01F + (Rnd() * 0.01F)) * -255
                moParticles(lIndex).fGChg = moParticles(lIndex).fRChg / 2
                moParticles(lIndex).fBChg = Math.Abs((64 - moParticles(lIndex).mfB) / ((moParticles(lIndex).mfR - 64) / moParticles(lIndex).fRChg))
            End If
        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
            'Dim matWorld As Matrix

            If EmitterStopped = True Then Return
            If mbInitialized = False Then MyBase.Initialize(ParticleCount)

            If mlPrevFrame = 0 Then fElapsed = 1.0F Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
            If fElapsed > 1.0F Then fElapsed = 1.0F
            mlPrevFrame = timeGetTime

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    .Update(fElapsed)
                    '.vecAcc.Y = .vecSpeed.Y / (-10.0F * (fElapsed / 30.0F))
                    If .mfA <= 0 Then ResetParticle(X)
                    If .vecSpeed.Y < 0 Then
                        .vecSpeed.Y = 0 : .vecAcc.Y = 0
                    End If

                    If EmitterStopped = True Then Return

                    moPoints(X).Color = .ParticleColor.ToArgb
                    moPoints(X).Position = .vecLoc

                End With
            Next X
        End Sub

        'Public Overrides Sub UpdateMapWrapSituation(ByVal yMapWrapSituation As Byte, ByVal lLocXMapWrapCheck As Integer)
        '    If yMapWrapSituation <> myPrevMapWrapSituation Then
        '        mbCullBoxSet = False
        '        'Reset our loc
        '        If vecEmitter.X < goCurrentEnvir.lMinXPos Then
        '            vecEmitter.X += goCurrentEnvir.lMapWrapAdjustX
        '        ElseIf vecEmitter.X > goCurrentEnvir.lMaxXPos Then
        '            vecEmitter.X -= goCurrentEnvir.lMapWrapAdjustX
        '        End If

        '        If yMapWrapSituation = 1 Then
        '            'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
        '            If vecEmitter.X > lLocXMapWrapCheck Then
        '                vecEmitter.X -= goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        ElseIf yMapWrapSituation = 2 Then
        '            'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
        '            If vecEmitter.X < lLocXMapWrapCheck Then
        '                vecEmitter.X += goCurrentEnvir.lMapWrapAdjustX
        '            End If
        '        End If
        '    End If
        '    myPrevMapWrapSituation = yMapWrapSituation
        'End Sub
    End Class
End Namespace