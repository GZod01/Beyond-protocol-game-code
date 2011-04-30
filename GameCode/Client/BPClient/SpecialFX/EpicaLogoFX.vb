Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Public Class EpicaLogoFX
'    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32
'    Private Class Particle
'        Public vecLoc As Vector3
'        Public vecAcc As Vector3
'        Public vecSpeed As Vector3

'        Public mfA As Single
'        Public mfR As Single
'        Public mfG As Single
'        Public mfB As Single

'        Public fAChg As Single
'        Public fRChg As Single
'        Public fGChg As Single
'        Public fBChg As Single

'        Public ParticleColor As System.Drawing.Color

'        Public ParticleActive As Boolean

'        Public Sub Update(ByVal fElapsedTime As Single)
'            vecLoc.Add(Vector3.Multiply(vecSpeed, fElapsedTime))
'            vecSpeed.Add(Vector3.Multiply(vecAcc, fElapsedTime))

'            mfA += (fAChg * fElapsedTime)
'            mfR += (fRChg * fElapsedTime)
'            mfG += (fGChg * fElapsedTime)
'            mfB += (fBChg * fElapsedTime)

'            If mfA < 0 Then mfA = 0
'            If mfA > 255 Then mfA = 255
'            If mfR < 0 Then mfR = 0
'            If mfR > 255 Then mfR = 255
'            If mfG < 0 Then mfG = 0
'            If mfG > 255 Then mfG = 255
'            If mfB < 0 Then mfB = 0
'            If mfB > 255 Then mfB = 255

'            ParticleColor = System.Drawing.Color.FromArgb(CInt(mfA), CInt(mfR), CInt(mfG), CInt(mfB))
'        End Sub

'        Public Sub Reset(ByVal fX As Single, ByVal fY As Single, ByVal fZ As Single, ByVal fXSpeed As Single, ByVal fYSpeed As Single, ByVal fZSpeed As Single, ByVal fXAcc As Single, ByVal fYAcc As Single, ByVal fZAcc As Single, ByVal fR As Single, ByVal fG As Single, ByVal fB As Single, ByVal fA As Single)

'            Dim fZOffset As Single = goCamera.mlCameraZ + 1000

'            vecLoc.X = fX : vecLoc.Y = fY : vecLoc.Z = fZ + fZOffset
'            vecAcc.X = fXAcc : vecAcc.Y = fYAcc : vecAcc.Z = fZAcc
'            vecSpeed.X = fXSpeed : vecSpeed.Y = fYSpeed : vecSpeed.Z = fZSpeed
'            mfA = fA : mfR = fR : mfG = fG : mfB = fB

'            ParticleActive = True
'        End Sub

'    End Class

'    Public vecEmitter As Vector3

'    Private moParticles() As Particle
'    Private mlParticleUB As Int32
'    Private mlPrevFrame As Int32
'    Public mbEmitterStopping As Boolean = False

'    Private EmitterStopped As Boolean = False

'    Private mfMinExt() As Single
'    Private mfMaxExt() As Single
'    Private mbUseHorizontal() As Boolean

'    Private moDevice As Device
'    Private moParticle As Texture

'    Private mbFadedIn As Boolean = False

'    Public moPoints() As CustomVertex.PositionColoredTextured

'    Public Event Disposed()

'    Public Sub New(ByRef oDevice As Device, ByVal oVecLoc As Vector3, ByVal lParticleCnt As Int32)
'        Dim X As Int32

'        moDevice = oDevice

'        vecEmitter = oVecLoc
'        mlParticleUB = lParticleCnt - 1

'        ReDim moParticles(mlParticleUB)
'        ReDim mfMinExt(mlParticleUB)
'        ReDim mfMaxExt(mlParticleUB)
'        ReDim mbUseHorizontal(mlParticleUB)
'        ReDim moPoints(mlParticleUB)

'        For X = 0 To mlParticleUB
'            moParticles(X) = New Particle()
'            ResetParticle(X)
'        Next X

'    End Sub

'    Public Sub StopEmitter()
'        Dim X As Int32
'        mbEmitterStopping = True

'        For X = 0 To 1200   '680
'            moParticles(X).fAChg = -10
'        Next X
'    End Sub

'    Public Sub StartEmitter()
'        Dim X As Int32
'        mbEmitterStopping = False
'        For X = 0 To mlParticleUB
'            moParticles(X).ParticleActive = True
'            ResetParticle(X)
'        Next X
'    End Sub

'    Private Sub ResetParticle(ByVal lIndex As Int32)
'        Dim lX As Int32

'        If mbEmitterStopping = True Then
'            moParticles(lIndex).mfA = 0
'            moParticles(lIndex).ParticleActive = False

'            EmitterStopped = True
'            For lX = 0 To mlParticleUB
'                If moParticles(lX).ParticleActive = True Then
'                    EmitterStopped = False
'                    Exit For
'                End If
'            Next lX

'            If EmitterStopped = True Then DisposeMe()
'        Else
'            'IIIIIIIIIIIIIIIIIIIIIIIIIIIIIIIII
'            'DAAAA  HEEE   TTTTT  NLLLL   OOO
'            'D      H   G    K    N      Q   R
'            'DCC    HFFF     K    N      QPPPR
'            'D      H        K    N      Q   R
'            'DBBBB  H      UUUUU  NMMMM  Q   R
'            'JJJJJJJJJJJJJJJJJJJJJJJJJJJJJJJJJ

'            Dim fXS As Single
'            Dim fYS As Single
'            Dim lTemp As Int32
'            Dim lYLoc As Int32
'            Dim lXLoc As Int32

'            'Particles are 32 in size, middle is the middle K
'            If lIndex < 80 Then
'                'A group or B Group -15

'                If lIndex > 39 Then
'                    lTemp = lIndex - 40
'                    lYLoc = CInt(vecEmitter.Y + 64)
'                Else : lTemp = lIndex : lYLoc = CInt(vecEmitter.Y - 64)
'                End If

'                fXS = (GetNxtRnd() * 6) - 3
'                fYS = 0

'                'Ok, set up its extent...
'                mfMinExt(lIndex) = vecEmitter.X - 480
'                mfMaxExt(lIndex) = vecEmitter.X - 352
'                mbUseHorizontal(lIndex) = True

'                moParticles(lIndex).Reset(vecEmitter.X - 480 + (3.2F * lTemp), lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 100 Then
'                'C Group -15
'                lTemp = lIndex - 80
'                fXS = (GetNxtRnd() * 6) - 3
'                fYS = 0

'                'ok, set up its extent
'                mfMinExt(lIndex) = vecEmitter.X - 480
'                mfMaxExt(lIndex) = vecEmitter.X - 416
'                mbUseHorizontal(lIndex) = True

'                moParticles(lIndex).Reset(vecEmitter.X - 480 + (3.2F * lTemp), vecEmitter.Y, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 150 Then
'                'D Group -16
'                lTemp = lIndex - 100
'                fXS = 0
'                fYS = (GetNxtRnd() * 6) - 3

'                'ok, set up its extent
'                mfMinExt(lIndex) = vecEmitter.Y - 64
'                mfMaxExt(lIndex) = vecEmitter.Y + 64
'                mbUseHorizontal(lIndex) = False

'                moParticles(lIndex).Reset(vecEmitter.X - 480, vecEmitter.Y - 64 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 210 Then
'                'E and F Group
'                lTemp = lIndex - 150
'                fXS = GetNxtRnd() * 6 - 3
'                fYS = 0

'                If lTemp > 29 Then
'                    lTemp -= 30
'                    lYLoc = CInt(vecEmitter.Y + 64)
'                Else : lYLoc = CInt(vecEmitter.Y) '- 64
'                End If

'                mfMinExt(lIndex) = vecEmitter.X - 256
'                mfMaxExt(lIndex) = vecEmitter.X - 160
'                mbUseHorizontal(lIndex) = True

'                moParticles(lIndex).Reset(vecEmitter.X - 256 + (3.2F * lTemp), lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 220 Then
'                'G Group
'                lTemp = lIndex - 210
'                fXS = 0 : fYS = GetNxtRnd() * 6 - 3

'                mfMinExt(lIndex) = vecEmitter.Y : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False
'                moParticles(lIndex).Reset(vecEmitter.X - 160, vecEmitter.Y + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 270 Then
'                'H Group
'                lTemp = lIndex - 220
'                fXS = 0 : fYS = GetNxtRnd() * 6 - 3
'                mfMinExt(lIndex) = vecEmitter.Y - 64 : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False
'                moParticles(lIndex).Reset(vecEmitter.X - 256, vecEmitter.Y - 64 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 370 Then
'                'I or J group
'                lTemp = lIndex - 270
'                fXS = GetNxtRnd() * 6 - 3 : fYS = 0
'                If lTemp > 49 Then
'                    lTemp -= 50
'                    lYLoc = CInt(vecEmitter.Y - 100)   '64
'                Else : lYLoc = CInt(vecEmitter.Y + 100) '64    
'                End If

'                'mfMinExt(lIndex) = vecEmitter.X - 64 : mfMaxExt(lIndex) = vecEmitter.X + 96 : mbUseHorizontal(lIndex) = True
'                mfMinExt(lIndex) = vecEmitter.X - 512 : mfMaxExt(lIndex) = vecEmitter.X + 512 : mbUseHorizontal(lIndex) = True
'                moParticles(lIndex).Reset(vecEmitter.X - 64 + (3.2F * lTemp), lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 410 Then
'                'K Group
'                lTemp = lIndex - 370
'                fXS = 0 : fYS = GetNxtRnd() * 6 - 3
'                mfMinExt(lIndex) = vecEmitter.Y - 64 : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False '64
'                moParticles(lIndex).Reset(vecEmitter.X + 16, vecEmitter.Y - 32 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 490 Then
'                'L and M group
'                lTemp = lIndex - 410
'                fXS = GetNxtRnd() * 6 - 3 : fYS = 0
'                If lTemp > 39 Then
'                    lTemp -= 40
'                    lYLoc = CInt(vecEmitter.Y - 64)
'                Else : lYLoc = CInt(vecEmitter.Y + 64)
'                End If
'                mfMinExt(lIndex) = vecEmitter.X + 192 : mfMaxExt(lIndex) = vecEmitter.X + 320 : mbUseHorizontal(lIndex) = True
'                moParticles(lIndex).Reset(vecEmitter.X + 192 + (3.2F * lTemp), lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 540 Then
'                'N Group
'                lTemp = lIndex - 490
'                fXS = 0 : fYS = GetNxtRnd() * 6 - 3
'                mfMinExt(lIndex) = vecEmitter.Y - 64 : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False
'                moParticles(lIndex).Reset(vecEmitter.X + 192, vecEmitter.Y - 64 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 600 Then
'                'O and P Group
'                lTemp = lIndex - 540
'                fXS = GetNxtRnd() * 6 - 3 : fYS = 0
'                If lTemp > 29 Then
'                    lTemp -= 30
'                    lYLoc = CInt(vecEmitter.Y)
'                Else : lYLoc = CInt(vecEmitter.Y + 64)
'                End If
'                mfMinExt(lIndex) = vecEmitter.X + 416 : mfMaxExt(lIndex) = vecEmitter.X + 512 : mbUseHorizontal(lIndex) = True
'                moParticles(lIndex).Reset(vecEmitter.X + 416 + (3.2F * lTemp), lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 680 Then
'                'Q and R Group
'                lTemp = lIndex - 600
'                fXS = 0 : fYS = GetNxtRnd() * 6 - 3
'                If lTemp > 39 Then
'                    lTemp -= 40
'                    lXLoc = CInt(vecEmitter.X + 416)
'                Else : lXLoc = CInt(vecEmitter.X + 512)
'                End If
'                mfMinExt(lIndex) = vecEmitter.Y - 64 : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False
'                moParticles(lIndex).Reset(lXLoc, vecEmitter.Y - 64 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 760 Then
'                'T and U Group
'                lTemp = lIndex - 680
'                fXS = GetNxtRnd() * 6 - 3 : fYS = 0
'                If lTemp > 39 Then
'                    lTemp -= 40
'                    lYLoc = CInt(vecEmitter.Y - 64)
'                Else : lYLoc = CInt(vecEmitter.Y + 64)
'                End If
'                mfMinExt(lIndex) = vecEmitter.X - 64 : mfMaxExt(lIndex) = vecEmitter.X + 92 : mbUseHorizontal(lIndex) = True
'                moParticles(lIndex).Reset(vecEmitter.X - 64 + (3.2F * lTemp), lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 1200 Then
'                'S Group
'                lTemp = lIndex - 760
'                fXS = GetNxtRnd() * 6 - 3 : fYS = 0
'                If lTemp > 220 Then
'                    lTemp -= 110
'                    lYLoc = CInt(vecEmitter.Y + 100)
'                Else : lYLoc = CInt(vecEmitter.Y - 100)
'                End If
'                lXLoc = CInt((GetNxtRnd() * 1024) - 512) '+ vecEmitter.X
'                If lXLoc < 0 Then fXS = (GetNxtRnd() * 5) + 1 Else fXS = -((GetNxtRnd() * 5) + 1)
'                mfMinExt(lIndex) = vecEmitter.X - 512 : mfMaxExt(lIndex) = vecEmitter.X + 512 : mbUseHorizontal(lIndex) = True
'                moParticles(lIndex).Reset(lXLoc, lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            Else
'                lTemp = CInt(((lIndex - 1200) / (mlParticleUB - 1200)) * 1200)
'                lXLoc = CInt(moParticles(lTemp).vecLoc.X)
'                lYLoc = CInt(moParticles(lTemp).vecLoc.Y)

'                Dim fXA As Single
'                Dim fYA As Single

'                If GetNxtRnd() * 100 < 50 Then
'                    fXS = -((GetNxtRnd() * 0.25F) + 0.3F) : fXA = GetNxtRnd() * 0.01F
'                Else : fXS = GetNxtRnd() * 0.25F + 0.3F : fXA = GetNxtRnd() * -0.01F
'                End If
'                If GetNxtRnd() * 100 < 50 Then
'                    fYS = -((GetNxtRnd() * 0.25F) + 0.3F) : fYA = GetNxtRnd() * 0.001F
'                Else : fYS = GetNxtRnd() * 0.25F + 0.3F : fYA = GetNxtRnd() * -0.001F
'                End If

'                moParticles(lIndex).Reset(lXLoc, lYLoc, 0, fXS, fYS, 0, fXA, fYA, 0, 255, 255, 255, 255)
'                moParticles(lIndex).fAChg = -((GetNxtRnd() * 3) + 2)

'                If mbFadedIn = False Then
'                    moParticles(lIndex).mfA = 5
'                End If

'                'moParticles(lIndex).fRChg = -((GetNxtRnd() * 5) + 25)
'                moParticles(lIndex).fGChg = -((GetNxtRnd() * 5) + 1)
'                moParticles(lIndex).fBChg = -((GetNxtRnd() * 1) + 1)
'            End If

'            If lIndex < 1200 Then       '680
'                moParticles(lIndex).mfA = 0
'                moParticles(lIndex).fAChg = 1
'                moParticles(lIndex).fRChg = 0
'                moParticles(lIndex).fGChg = GetNxtRnd() * -5
'                moParticles(lIndex).fBChg = GetNxtRnd() * -5
'            End If

'            moParticles(lIndex).mfR = 64
'            moParticles(lIndex).mfG = 255
'            moParticles(lIndex).mfB = 255
'        End If


'    End Sub

'    Private Sub Update()
'        Dim fElapsed As Single
'        Dim X As Int32
'        Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
'        'Dim matWorld As Matrix

'        If EmitterStopped = True Then Exit Sub

'        If mlPrevFrame = 0 Then fElapsed = 300 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
'        mlPrevFrame = timeGetTime

'        'Update the particles
'        For X = 0 To mlParticleUB
'            With moParticles(X)
'                .Update(fElapsed)

'                If X < 1200 Then ' 680 Then

'                    If mbUseHorizontal(X) = True Then
'                        If .vecLoc.X <= mfMinExt(X) Then
'                            .vecSpeed.X = Math.Abs(.vecSpeed.X)
'                        ElseIf .vecLoc.X >= mfMaxExt(X) Then
'                            .vecSpeed.X = -(Math.Abs(.vecSpeed.X))
'                        End If
'                    Else
'                        If .vecLoc.Y <= mfMinExt(X) Then
'                            .vecSpeed.Y = Math.Abs(.vecSpeed.Y)
'                        ElseIf .vecLoc.Y >= mfMaxExt(X) Then
'                            .vecSpeed.Y = -(Math.Abs(.vecSpeed.Y))
'                        End If
'                    End If
'                    'If .mfR <= 64 Then .fRChg = GetNxtRnd() * 5 : .mfR = 64
'                    If .mfG <= 64 Then .fGChg = GetNxtRnd() * 5 : .mfG = 64
'                    If .mfB <= 64 Then .fBChg = GetNxtRnd() * 5 : .mfB = 64

'                    'Ok, check if we have an alpha at 255
'                    If mbFadedIn = False AndAlso .mfA = 255 Then
'                        mbFadedIn = True
'                        'Debug.Write("Faded In")
'                    End If

'                    If mbEmitterStopping = True AndAlso .mfA <= 0 Then ResetParticle(X)

'                ElseIf .mfA <= 0 Then
'                    ResetParticle(X)
'                End If

'                If EmitterStopped = True Then Exit Sub

'                moPoints(X).Color = .ParticleColor.ToArgb
'                moPoints(X).Position = .vecLoc


'            End With
'        Next X

'    End Sub

'    Public Sub Render(ByVal bUpdateNoRender As Boolean)
'        'Dim X As Int32

'        Update()

'        If EmitterStopped = True OrElse bUpdateNoRender = True Then Exit Sub

'        'now, render
'        Device.IsUsingEventHandlers = False
'        If moParticle Is Nothing OrElse moParticle.Disposed = True Then
'            moParticle = goResMgr.GetTexture("Particle.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
'        End If
'        Device.IsUsingEventHandlers = True

'        With moDevice
'            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
'            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
'            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

'            'Ok, if our device was created with mixed...
'            If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
'                'Set us up for software vertex processing as point sprites always work in software
'                moDevice.SoftwareVertexProcessing = True
'            End If

'            .RenderState.PointSpriteEnable = True

'            If .DeviceCaps.VertexFormatCaps.SupportsPointSize = True Then
'                .RenderState.PointScaleEnable = True
'                If .DeviceCaps.MaxPointSize > 48 Then
'                    .RenderState.PointSize = 48 '16
'                Else : .RenderState.PointSize = .DeviceCaps.MaxPointSize
'                End If
'                .RenderState.PointScaleA = 0
'                .RenderState.PointScaleB = 0
'                .RenderState.PointScaleC = 1
'            End If

'            .RenderState.SourceBlend = Blend.SourceAlpha
'            .RenderState.DestinationBlend = Blend.One
'            .RenderState.AlphaBlendEnable = True

'            .RenderState.ZBufferWriteEnable = False

'            .RenderState.Lighting = False

'            .VertexFormat = CustomVertex.PositionColoredTextured.Format
'            .SetTexture(0, moParticle)
'            .DrawUserPrimitives(PrimitiveType.PointList, mlParticleUB + 1, moPoints)
'            .RenderState.ZBufferWriteEnable = True
'            .RenderState.Lighting = True
'            .RenderState.PointSpriteEnable = False

'            'Ok, if our device was created with mixed...
'            If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
'                'Set us up for software vertex processing as point sprites always work in software
'                moDevice.SoftwareVertexProcessing = False
'            End If

'            .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
'            .RenderState.DestinationBlend = Blend.InvSourceAlpha ' Blend.Zero
'            .RenderState.AlphaBlendEnable = True

'            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
'            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
'            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
'        End With


'    End Sub

'    Public Sub DisposeMe()
'        Erase moParticles
'        Erase moPoints
'        moParticle = Nothing
'        moDevice = Nothing

'        RaiseEvent Disposed()
'    End Sub
'End Class