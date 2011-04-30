Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Namespace BurnFX
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
            eEngineEmitter = 0
            eSmokeyFireEmitter
            eFireEmitter                                'Defaults to horizontal (x-z) going positive Y

            eFireEmitterMod_MinusY = 33554432           'horizontal (x-z) emitter going Negative in the Y column
            eFireEmitterMod_PlusX = 67108864            'vertical (y-z) emitter going positive in the X
            eFireEmitterMod_MinusX = 134217728          'vertical (y-z) emitter going Negative in the X Column
            eFireEmitterMod_PlusZ = 268435456           'vertical (x-y) emitter going positive z
            eFireEmitterMod_MinusZ = 536870912          'vertical (x-y) emitter going negative z
        End Enum

        Private moChildren() As FXEmitter
        Private myChildUsed() As Byte
        Public mlChildrenUB As Int32 = -1

        'The standard texture to use for emitters using this engine object
        Private moTex As Texture
        Private mfParticleSize As Single

        Private mlLimitedParticleLvl As Int32

		Public Shared l_PARTICLE_CNT As Int32

        Public Sub New(ByVal fParticleSize As Single)
            mfParticleSize = fParticleSize
            ReDim myChildUsed(0)
            myChildUsed(0) = 0

            mlLimitedParticleLvl = muSettings.BurnFXParticles
        End Sub

        Public Sub Render(ByVal bUpdateNoRender As Boolean)
            'Dim matWorld As Matrix
            Dim X As Int32

            On Error Resume Next

            If moTex Is Nothing OrElse moTex.Disposed = True Then
				moTex = goResMgr.GetTexture("ExpParticle2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "Misc.pak")
            End If

            If mlChildrenUB = -1 Then Return
            If mlLimitedParticleLvl = 0 Then Return

            Dim moDevice As Device = GFXEngine.moDevice

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
                        If .DeviceCaps.MaxPointSize > mfParticleSize Then
                            .RenderState.PointSize = mfParticleSize
                        Else : .RenderState.PointSize = .DeviceCaps.MaxPointSize
                        End If
                        .RenderState.PointScaleA = 0
                        .RenderState.PointScaleB = 0
                        .RenderState.PointScaleC = 1
                    End If

					.RenderState.SourceBlend = Blend.SourceAlpha
					.RenderState.DestinationBlend = Blend.One
					.RenderState.AlphaBlendEnable = True

					.RenderState.ZBufferWriteEnable = False

                    .RenderState.Lighting = False

                    .VertexFormat = CustomVertex.PositionColoredTextured.Format
                    '.VertexFormat = moPoints(0).Format
                    .SetTexture(0, moTex)

					glBurnFXRendered = 0
					l_PARTICLE_CNT = 0
                    For X = 0 To mlChildrenUB
                        If myChildUsed(X) > 0 AndAlso moChildren(X).AttachedObject.bCulled = False AndAlso moChildren(X).AttachedObject.yVisibility = eVisibilityType.Visible Then
                            'Now, check for whether the emitter can be seen
                            moChildren(X).Update()
                            If moChildren(X).EmitterStopped = True Then
                                myChildUsed(X) = 0
                            Else
                                l_PARTICLE_CNT += moChildren(X).mlCurrentParticleUB + 1
                                .DrawUserPrimitives(PrimitiveType.PointList, moChildren(X).mlCurrentParticleUB + 1, moChildren(X).moPoints)
                                glBurnFXRendered += 1
                            End If
                        End If
                    Next X
                    '.DrawUserPrimitives(PrimitiveType.PointList, uTmpList.GetUpperBound(0) + 1, uTmpList)

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
                End If
            End With

        End Sub

        Public Function AddEmitter(ByVal iType As EmitterType, ByVal oLoc As Vector3, ByVal lPCnt As Int32, ByVal oAttachedObject As RenderObject) As Int32
            Dim X As Int32
            Dim lIdx As Int32 = -1

            Dim lBaseType As Int32
            Dim lModValue As Int32

            If (iType And EmitterType.eEngineEmitter) <> 0 Then
                lBaseType = EmitterType.eEngineEmitter
            ElseIf (iType And EmitterType.eFireEmitter) <> 0 Then
                lBaseType = EmitterType.eFireEmitter
            ElseIf (iType And EmitterType.eSmokeyFireEmitter) <> 0 Then
                lBaseType = EmitterType.eSmokeyFireEmitter
            End If
            lModValue = iType - lBaseType

            If lBaseType = EmitterType.eEngineEmitter AndAlso muSettings.EngineFXParticles = 0 Then Return -1
            If lBaseType = EmitterType.eFireEmitter AndAlso muSettings.BurnFXParticles = 0 Then Return -1
            If lBaseType = EmitterType.eSmokeyFireEmitter AndAlso muSettings.BurnFXParticles = 0 Then Return -1

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

            Select Case lBaseType
                Case EmitterType.eEngineEmitter
                    If lPCnt <= 0 Then lPCnt = 30
                    Dim lActPCnt As Int32 = lPCnt
                    'Now, adjust the pcnt
                    Select Case muSettings.EngineFXParticles
                        Case 1      '25%
                            lActPCnt = CInt(lPCnt * 0.25F)
                        Case 2      '50%
                            lActPCnt = CInt(lPCnt * 0.5F)
                        Case 3      '75%
                            lActPCnt = CInt(lPCnt * 0.75F)
                    End Select

                    moChildren(lIdx) = New EngineFX(oLoc, Me, lPCnt, lActPCnt, oAttachedObject.clrEngineFX)
                    moChildren(lIdx).lEmitterType = EmitterType.eEngineEmitter
                Case EmitterType.eFireEmitter
                    If lPCnt <= 0 Then lPCnt = 50
                    Dim lActPCnt As Int32 = lPCnt
                    'Now, adjust the pcnt
                    Select Case muSettings.BurnFXParticles
                        Case 1      '25%
                            lActPCnt = CInt(lPCnt * 0.25F)
                        Case 2      '50%
                            lActPCnt = CInt(lPCnt * 0.5F)
                        Case 3      '75%
                            lActPCnt = CInt(lPCnt * 0.75F)
                    End Select

                    moChildren(lIdx) = New FireFX(oLoc, Me, lPCnt, lActPCnt)
                    CType(moChildren(lIdx), FireFX).lEmitterMod = CType(lModValue, ParticleEngine.EmitterType)
                    moChildren(lIdx).lEmitterType = EmitterType.eFireEmitter
                Case EmitterType.eSmokeyFireEmitter
                    If lPCnt <= 0 Then lPCnt = 100
                    Dim lActPCnt As Int32 = lPCnt
                    'Now, adjust the pcnt
                    Select Case muSettings.BurnFXParticles
                        Case 1      '25%
                            lActPCnt = CInt(lPCnt * 0.25F)
                        Case 2      '50%
                            lActPCnt = CInt(lPCnt * 0.5F)
                        Case 3      '75%
                            lActPCnt = CInt(lPCnt * 0.75F)
                    End Select

                    moChildren(lIdx) = New SmokeyFireFX(oLoc, Me, lPCnt, lActPCnt)
                    CType(moChildren(lIdx), SmokeyFireFX).lEmitterMod = CType(lModValue, ParticleEngine.EmitterType)
                    moChildren(lIdx).lEmitterType = EmitterType.eSmokeyFireEmitter
            End Select

            myChildUsed(lIdx) = 255

            moChildren(lIdx).AttachedObject = oAttachedObject
            moChildren(lIdx).Initialize()

            Return lIdx
        End Function

        Public Sub StopEmitter(ByVal lIndex As Int32)
            If lIndex > -1 AndAlso lIndex <= mlChildrenUB Then
                If myChildUsed(lIndex) > 0 Then moChildren(lIndex).StopEmitter()
            End If
        End Sub

        Public Sub MoveEmitter(ByVal lIndex As Int32, ByVal fVecX As Single, ByVal fVecY As Single, ByVal fVecZ As Single)
            If lIndex > -1 AndAlso lIndex <= mlChildrenUB Then
                If myChildUsed(lIndex) > 0 Then moChildren(lIndex).MoveEmitter(fVecX, fVecY, fVecZ)
            End If
        End Sub

        Public Sub ClearAllEmitters()
            mlChildrenUB = -1
            Erase moChildren
            Erase myChildUsed
        End Sub

        Public Sub ResetEngineFXPCnts()
            If muSettings.EngineFXParticles = 0 Then
                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 AndAlso moChildren(X).lEmitterType = EmitterType.eEngineEmitter Then
                        myChildUsed(X) = 0
                        moChildren(X) = Nothing
                    End If
                Next X
            Else
                Dim fMult As Single = 1.0F

                Select Case muSettings.EngineFXParticles
                    Case 1
                        fMult = 0.25F
                    Case 2
                        fMult = 0.5F
                    Case 3
                        fMult = 0.75
                End Select

                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 AndAlso moChildren(X).lEmitterType = EmitterType.eEngineEmitter Then
                        Dim lCnt As Int32 = moChildren(X).ParticleCount
                        Dim lActPCnt As Int32 = CInt(lCnt * fMult)
                        moChildren(X).AdjustActualPCnt(lActPCnt)
                    End If
                Next X
            End If
        End Sub

        Public Sub ResetBurnFXPCnts()
            If muSettings.BurnFXParticles = 0 Then
                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 AndAlso (moChildren(X).lEmitterType = EmitterType.eFireEmitter OrElse moChildren(X).lEmitterType = EmitterType.eSmokeyFireEmitter) Then
                        myChildUsed(X) = 0
                        moChildren(X) = Nothing
                    End If
                Next X
            Else
                Dim fMult As Single = 1.0F

                Select Case muSettings.BurnFXParticles
                    Case 1
                        fMult = 0.25F
                    Case 2
                        fMult = 0.5F
                    Case 3
                        fMult = 0.75
                End Select

                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 AndAlso (moChildren(X).lEmitterType = EmitterType.eFireEmitter OrElse moChildren(X).lEmitterType = EmitterType.eSmokeyFireEmitter) Then
                        Dim lCnt As Int32 = moChildren(X).ParticleCount
                        Dim lActPCnt As Int32 = CInt(lCnt * fMult)
                        moChildren(X).AdjustActualPCnt(lActPCnt)
                    End If
                Next X
            End If
        End Sub

        Public Sub IncreaseLimitOnParticles(ByVal lFPS As Int32)
            'telling us to reduce our quantity of particles for fire and engine fx down one lvl...
            '  this is a temporary setting until frame rate increases...
            mlLimitedParticleLvl -= 1
            If mlLimitedParticleLvl < 1 Then
                If lFPS < 22 Then mlLimitedParticleLvl = 0 Else mlLimitedParticleLvl = 1
                Return
            End If

            If mlLimitedParticleLvl = 4 Then
                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 Then
                        moChildren(X).mlCurrentParticleUB = moChildren(X).mlParticleUB
                    End If
                Next X
            Else
                Dim fMult As Single = 1.0F
                Select Case mlLimitedParticleLvl
                    Case 1
                        fMult = 0.25F
                    Case 2
                        fMult = 0.5F
                    Case 3
                        fMult = 0.75
                End Select
                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 Then
                        Dim lCnt As Int32 = CInt(moChildren(X).mlParticleUB * fMult)
                        moChildren(X).mlCurrentParticleUB = lCnt
                    End If
                Next X
            End If

        End Sub
        Public Sub DecreaseLimitOnParticles()
            'telling us that we can increase our quantity of particles for fire and engine fx up one lvl...
            '  this is a temporary setting... once we reach the original particle quantity, we are done...
            mlLimitedParticleLvl += 1
            If mlLimitedParticleLvl > 4 Then
                mlLimitedParticleLvl = 4
                Return
            End If

            If mlLimitedParticleLvl = 4 Then
                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 Then
                        moChildren(X).mlCurrentParticleUB = moChildren(X).mlParticleUB
                    End If
                Next X
            Else
                Dim fMult As Single = 1.0F
                Select Case mlLimitedParticleLvl
                    Case 1
                        fMult = 0.25F
                    Case 2
                        fMult = 0.5F
                    Case 3
                        fMult = 0.75
                End Select
                For X As Int32 = 0 To mlChildrenUB
                    If myChildUsed(X) <> 0 Then
                        Dim lCnt As Int32 = CInt(moChildren(X).mlParticleUB * fMult)
                        moChildren(X).mlCurrentParticleUB = lCnt
                    End If
                Next X
            End If
        End Sub

    End Class

    Public MustInherit Class FXEmitter
        Protected Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

        Protected mvecEmitter As Vector3

        Protected Const mlInvisColor As Int32 = 0 'System.Drawing.Color = System.Drawing.Color.FromArgb(0, 0, 0, 0)

        Protected Shared msw_Timer As Stopwatch

        Protected moParticles() As Particle
        Protected moParentEngine As ParticleEngine
        Protected mblPrevFrame As Int64
        Protected mbEmitterStopping As Boolean = False

        Public moPoints() As CustomVertex.PositionColoredTextured
        Public mlParticleUB As Int32
        Public mlCurrentParticleUB As Int32
        Public EmitterStopped As Boolean = False
        Public AttachedObject As RenderObject

        Protected MustOverride Sub ResetParticle(ByVal lIndex As Int32)
        Public MustOverride Sub Update()

        Public lEmitterType As ParticleEngine.EmitterType

        Private mlPCnt As Int32

        Public Sub MoveEmitter(ByVal fVecX As Single, ByVal fVecY As Single, ByVal fVecZ As Single)
            mvecEmitter.Add(New Vector3(fVecX, fVecY, fVecZ))
        End Sub

        Public Function GetAbsolutePosition() As Vector3
            If AttachedObject Is Nothing Then
                Return mvecEmitter
            Else : Return New Vector3(mvecEmitter.X + AttachedObject.LocX, mvecEmitter.Y + AttachedObject.LocY, mvecEmitter.Z + AttachedObject.LocZ)
            End If
        End Function

        Public ReadOnly Property ParticleCount() As Int32
            Get
                Return mlPCnt
            End Get
        End Property

        Public Sub New(ByVal oVecLoc As Vector3, ByRef oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32, ByVal lActualPCnt As Int32)
            If msw_Timer Is Nothing Then msw_Timer = Stopwatch.StartNew
            mlPCnt = lParticleCnt
            moParentEngine = oParentEngine
            mvecEmitter = oVecLoc
            mlParticleUB = lActualPCnt - 1
            mlCurrentParticleUB = mlParticleUB
        End Sub

        Public Sub Initialize()
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB)

            For X As Int32 = 0 To mlCurrentParticleUB
                moParticles(X) = New Particle()
                ResetParticle(X)
            Next X
        End Sub

        Public Sub StopEmitter()
            mbEmitterStopping = True
        End Sub

        Public Sub StartEmitter()
            Dim X As Int32
            mbEmitterStopping = False
            For X = 0 To mlCurrentParticleUB
                moParticles(X).ParticleActive = True
                ResetParticle(X)
            Next X
        End Sub

        Public Sub AdjustActualPCnt(ByVal lNewCnt As Int32)
            Dim lNewUB As Int32 = lNewCnt - 1
            If lNewUB = mlParticleUB Then Return

            If lNewUB < mlParticleUB Then
                mlParticleUB = lNewUB
                ReDim Preserve moParticles(mlParticleUB)
                ReDim Preserve moPoints(mlParticleUB)
            Else
                ReDim Preserve moParticles(lNewUB)
                ReDim Preserve moPoints(lNewUB)
                For X As Int32 = mlParticleUB + 1 To lNewUB
                    moParticles(X) = New Particle()
                    ResetParticle(X)
                Next X
            End If
        End Sub
    End Class

    Public Class EngineFX
        Inherits FXEmitter

        Private mfAlphaChgMult As Single

        Public lBaseR As Int32 = 64
        Public lBaseG As Int32 = 128
        Public lBaseB As Int32 = 255

        Private mvecPrevLoc As Vector3      'of our attached object

        Public Sub New(ByVal oVecLoc As Vector3, ByVal oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32, ByVal lActPCnt As Int32, ByVal clrFX As System.Drawing.Color)
            MyBase.New(oVecLoc, oParentEngine, lParticleCnt, lActPCnt)
            lBaseR = clrFX.R : lBaseG = clrFX.G : lBaseB = clrFX.B
        End Sub

        Protected Overrides Sub ResetParticle(ByVal lIndex As Int32)
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
            Dim fTemp As Single

            Dim lX As Int32

            If mbEmitterStopping = True Then
                MyBase.moParticles(lIndex).mfA = 0
                MyBase.moParticles(lIndex).ParticleActive = False
                MyBase.moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                MyBase.EmitterStopped = True
                For lX = 0 To MyBase.mlCurrentParticleUB
                    If MyBase.moParticles(lX).ParticleActive = True Then
                        MyBase.EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                'Speed
                fXS = Rnd() * 1.0F - 0.1F
                fYS = Rnd() * 1.0F - 0.1F
                fZS = 10.0F

                '6/13/07 - removed cuz it didn't work
                ''MSC - 9/8/6 - Added this to see if the engine fx would look a little better
                'If MyBase.AttachedObject Is Nothing = False Then
                '    If CType(MyBase.AttachedObject, EpicaEntity).TotalVelocity > 10.0F Then
                '        fZS = 2.0F
                '    End If
                'End If

                'Accel
                If fXS < 0 Then fXA = Rnd() * -0.01F Else fXA = Rnd() * 0.01F
                If fYS < 0 Then fYA = Rnd() * -0.01F Else fYA = Rnd() * 0.01F
                fZA = 0

                'Now, offset our starting X and Y based on the particle count... at 50, we use the X and Y
                If MyBase.ParticleCount > 50 Then
                    fTemp = MyBase.ParticleCount / 10.0F
                    fOffsetX = ((fTemp * 2) * Rnd()) - fTemp
                    fOffsetY = ((fTemp * 2) * Rnd()) - fTemp
                End If


                fX = MyBase.mvecEmitter.X + fOffsetX
                fY = MyBase.mvecEmitter.Y + fOffsetY
                fZ = MyBase.mvecEmitter.Z

                'Now, rotate everyone
                RotateVals(fX, fY, fZ, fXS, fYS, fZS, fXA, fZA)

                'Now, offset all of that by the location of our attached object
                If MyBase.AttachedObject Is Nothing = False Then
                    'fX += MyBase.AttachedObject.LocX
                    fX += MyBase.AttachedObject.LocX
                    fY += MyBase.AttachedObject.LocY
                    fZ += MyBase.AttachedObject.LocZ
                End If

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, lBaseR, lBaseG, lBaseB, 255)
                If mfAlphaChgMult = 0 Then
                    If MyBase.ParticleCount < 50 Then
                        mfAlphaChgMult = 0.1F * (100.0F / (MyBase.ParticleCount * 2))
                    Else
                        mfAlphaChgMult = 0.1F
                    End If
                End If
                moParticles(lIndex).fAChg = -((mfAlphaChgMult + (mfAlphaChgMult * Rnd())) * 255)
            End If


        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
            'Dim matWorld As Matrix

            If MyBase.EmitterStopped = True Then Exit Sub

            Try

                Dim blCurrMS As Int64 = msw_Timer.ElapsedMilliseconds
                If mblPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (blcurrms - mblPrevFrame) * 0.033F
                If fElapsed > 1.0F Then fElapsed = 1.0F
                mblPrevFrame = blCurrMS


                If MyBase.AttachedObject Is Nothing = False Then
                    mvecPrevLoc.X = MyBase.AttachedObject.LocX
                    mvecPrevLoc.Y = MyBase.AttachedObject.LocY
                    mvecPrevLoc.Z = MyBase.AttachedObject.LocZ
                End If

                'Dim fVecX As Single = 0.0F
                'Dim fVecZ As Single = 0.0F

                'If AttachedObject.ObjTypeID = ObjectType.eUnit OrElse AttachedObject.ObjTypeID = ObjectType.eFacility Then
                '    With CType(AttachedObject, EpicaEntity)
                '        fVecX = .VelX * fElapsed * 0.5F
                '        fVecZ = .VelZ * fElapsed * 0.5F
                '    End With
                'End If

                'Update the particles

                Dim bSetInvisColor As Boolean = MyBase.AttachedObject Is Nothing OrElse MyBase.AttachedObject.yVisibility <> eVisibilityType.Visible

                For X = 0 To mlCurrentParticleUB
                    With moParticles(X)
                        If .ParticleActive = True Then
                            .Update(fElapsed)
                            If .mfA <= 0 Then
                                ResetParticle(X)
                            End If

                            '.vecLoc.X += fVecX
                            '.vecLoc.Z += fVecZ

                            If bSetInvisColor = True Then
                                MyBase.moPoints(X).Color = mlInvisColor
                            Else
                                MyBase.moPoints(X).Color = .ParticleColor.ToArgb
                            End If

                            moPoints(X).Position = .vecLoc
                        End If
                    End With
                Next X
            Catch

            End Try
        End Sub

        Private Sub RotateVals(ByRef fXLoc As Single, ByRef fYLoc As Single, ByRef fZLoc As Single, ByRef fXSpeed As Single, ByRef fYSpeed As Single, ByRef fZSpeed As Single, ByRef fXAcc As Single, ByRef fZAcc As Single)
            'Dim fDX As Single
            'Dim fDZ As Single

            If MyBase.AttachedObject Is Nothing Then Return

            'vecCol1 = m11, m21, m31
            'vecCol2 = m12, m22, m32
            'vecCol3 = m13, m23, m33

            'Dim vecLoc As Vector3 = New Vector3(fXLoc, fYLoc, fZLoc)
            Dim matWorld As Matrix = MyBase.AttachedObject.GetWorldMatrix()

            With matWorld
                Dim fX As Single = fXLoc
                Dim fY As Single = fYLoc
                Dim fZ As Single = fZLoc
                fXLoc = (fX * .M11) + (fY * .M21) + (fZ * .M31) '+ .M41
                fYLoc = (fX * .M12) + (fY * .M22) + (fZ * .M32) '+ .M42
                fZLoc = (fX * .M13) + (fY * .M23) + (fZ * .M33) '+ .M43
                'fXLoc = fX : fYLoc = fY : fZLoc = fZ

                fX = fXSpeed
                fY = fYSpeed
                fZ = fZSpeed
                fXSpeed = (fX * .M11) + (fY * .M21) + (fZ * .M31) '+ .M41
                fYSpeed = (fX * .M12) + (fY * .M22) + (fZ * .M32) '+ .M42
                fZSpeed = (fX * .M13) + (fY * .M23) + (fZ * .M33) '+ .M43
                'fX = (fXSpeed * .M11) + (fYSpeed * .M21) + (fZSpeed * .M31)
                'fY = (fXSpeed * .M12) + (fYSpeed * .M22) + (fZSpeed * .M32)
                'fZ = (fXSpeed * .M13) + (fYSpeed * .M23) + (fZSpeed * .M33)
                'fXSpeed = fX : fYSpeed = fY : fZSpeed = fZ

                fX = fXAcc
                fZ = fZAcc
                fXAcc = (fX * .M11) + (fZ * .M31)
                fZAcc = (fX * .M13) + (fZ * .M33)
                'fXAcc = fX : fZAcc = fZ
            End With

            'Dim vecTemp As Vector4 = Vector3.Transform(vecLoc, matWorld)
            'With vecTemp
            '    fXLoc = .X : fYLoc = .Y : fZLoc = .Z
            'End With
            'Dim vecSpeed As New Vector3(fXLoc, fYLoc, fZLoc)
            ''vecSpeed.TransformCoordinate(matWorld)
            'vecTemp = Vector3.Transform(vecSpeed, matWorld)
            'With vecTemp
            '    fXSpeed = .X : fYSpeed = .Y : fZSpeed = .Z
            'End With
            'Dim vecAcc As New Vector3(fXAcc, 0, fZAcc)
            ''vecAcc.TransformCoordinate(matWorld)
            'vecTemp = Vector3.Transform(vecAcc, matWorld)
            'With vecTemp
            '    fXAcc = .X : fZAcc = .Z
            'End With


            'Dim fRads As Single
            'Dim fCosR As Single
            'Dim fSinR As Single

            ''pitch
            'Dim fRads2 As Single = -((AttachedObject.mfTrueLocPitch / 10.0F)) * gdRadPerDegree
            'Dim fCosR2 As Single = CSng(Math.Cos(fRads2))
            'Dim fSinR2 As Single = CSng(Math.Sin(fRads2))
            'fDX = fZLoc
            'fDZ = fYLoc
            'fZLoc = +((fDX * fCosR2) + (fDZ * fSinR2))
            'fYLoc = -((fDX * fSinR2) + (fDZ * fCosR2))

            ''Yaw...
            'Dim fTempVal As Single = (AttachedObject.LocYaw - 1800)
            'If fTempVal < 0 Then fTempVal += 3600.0F
            ''fRads = -((AttachedObject.LocYaw / 10.0F)) * gdRadPerDegree
            'fRads = -((fTempVal / 10.0F)) * gdRadPerDegree
            'fCosR = CSng(Math.Cos(fRads))
            'fSinR = CSng(Math.Sin(fRads))
            'fDX = fXLoc
            'fDZ = fYLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Now set up for standard rotation...
            'fRads = ((AttachedObject.LocAngle / 10.0F) - 90) * gdRadPerDegree
            'fCosR = CSng(Math.Cos(fRads))
            'fSinR = CSng(Math.Sin(fRads))
            'fDX = fXLoc
            'fDZ = fZLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Speed
            'fDX = fXSpeed
            'fDZ = fZSpeed
            'fXSpeed = +((fDX * fCosR) + (fDZ * fSinR))
            'fZSpeed = -((fDX * fSinR) - (fDZ * fCosR))

            'fDX = fZSpeed
            'fDZ = fYSpeed
            'fZSpeed = +((fDX * fCosR2) + (fDZ * fSinR2))
            'fYSpeed = -((fDX * fSinR2) + (fDZ * fCosR2))

            ''Acc
            'fDX = fXAcc
            'fDZ = fZAcc
            'fXAcc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZAcc = -((fDX * fSinR) - (fDZ * fCosR))

        End Sub

    End Class

    'Public Class SmokeyFireFX
    '    Inherits FXEmitter

    '    Public Sub New(ByVal oVecLoc As Vector3, ByVal oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32)
    '        MyBase.New(oVecLoc, oParentEngine, lParticleCnt)
    '    End Sub

    '    Protected Overrides Sub ResetParticle(ByVal lIndex As Int32)
    '        Dim lX As Int32

    '        If mbEmitterStopping = True Then
    '            MyBase.moParticles(lIndex).mfA = 0
    '            MyBase.moParticles(lIndex).ParticleActive = False
    '            MyBase.moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

    '            MyBase.EmitterStopped = True
    '            For lX = 0 To MyBase.mlParticleUB
    '                If MyBase.moParticles(lX).ParticleActive = True Then
    '                    MyBase.EmitterStopped = False
    '                    Exit For
    '                End If
    '            Next lX
    '        Else
    '            Dim vecEmitter As Vector3 = MyBase.GetAbsolutePosition()

    '            'moParticles(lIndex).Reset(vecEmitter.X + ((GetNxtRnd() * 10) - 5), vecEmitter.Y, vecEmitter.Z + ((GetNxtRnd() * 10) - 5), 0, GetNxtRnd(), 0, 0, -0.001, 0, 255, 128, GetNxtRnd() * 50, (0.8F + (0.2 * GetNxtRnd())) * 255)
    '            moParticles(lIndex).Reset(vecEmitter.X + ((GetNxtRnd() * 10) - 5), vecEmitter.Y, vecEmitter.Z + ((GetNxtRnd() * 10) - 5), GetNxtRnd() * 0.5F - 0.25F, GetNxtRnd() + 1, GetNxtRnd() * 0.5F - 0.25F, 0, -0.001, 0, 255, 128, GetNxtRnd() * 50, (0.8F + (0.2F * GetNxtRnd())) * 255)
    '            moParticles(lIndex).fAChg = (0.01F - (GetNxtRnd() * 0.005F)) * -255
    '            moParticles(lIndex).fRChg = (0.01F + (GetNxtRnd() * 0.01F)) * -255
    '            moParticles(lIndex).fGChg = moParticles(lIndex).fRChg / 2
    '            moParticles(lIndex).fBChg = Math.Abs((64 - moParticles(lIndex).mfB) / ((moParticles(lIndex).mfR - 64) / moParticles(lIndex).fRChg))
    '        End If


    '    End Sub

    '    Public Overrides Sub Update()
    '        Dim fElapsed As Single
    '        Dim X As Int32
    '        Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
    '        'Dim matWorld As Matrix

    '        If MyBase.EmitterStopped = True Then Exit Sub

    '        If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
    '        mlPrevFrame = timeGetTime

    '        'Update the particles
    '        For X = 0 To mlParticleUB
    '            With moParticles(X)
    '                .Update(fElapsed)
    '                If .mfA <= 0 Then
    '                    ResetParticle(X)
    '                End If

    '                'check our limits
    '                If .mfR < 64 Then .mfR = 64
    '                If .mfG < 64 Then .mfG = 64
    '                If .mfB > 64 Then .mfB = 64

    '                MyBase.moPoints(X).Color = FXEmitter.moInvisColor.ToArgb
    '                If MyBase.AttachedObject Is Nothing = False Then
    '                    If MyBase.AttachedObject.yVisibility = eVisibilityType.Visible Then
    '                        MyBase.moPoints(X).Color = .ParticleColor.ToArgb
    '                    End If
    '                End If
    '                MyBase.moPoints(X).Position = .vecLoc
    '            End With
    '        Next X

    '    End Sub

    'End Class

    Public Class FireFX
        Inherits FXEmitter

        Private mfXLocMin As Single
        Private mfYLocMin As Single
        Private mfZLocMin As Single

        Private mfXLocMax As Single
        Private mfYLocMax As Single
        Private mfZLocMax As Single

        Private mfXSpeedMin As Single
        Private mfYSpeedMin As Single
        Private mfZSpeedMin As Single

        Private mfXSpeedMax As Single
        Private mfYSpeedMax As Single
        Private mfZSpeedMax As Single

        Private mfXAccMin As Single
        Private mfXAccMax As Single

        Private mfYAccMin As Single
        Private mfYAccMax As Single

        Private mfZAccMin As Single
        Private mfZAccMax As Single

        Private mfAO_VX As Single
        Private mfAO_VZ As Single

        Private mlEmitterMod As ParticleEngine.EmitterType
        Public Property lEmitterMod() As ParticleEngine.EmitterType
            Get
                Return mlEmitterMod
            End Get
            Set(ByVal value As ParticleEngine.EmitterType)
                mlEmitterMod = value

                Dim lRndRng As Int32 = Me.ParticleCount \ 3
                Dim lHlfRndRng As Int32 = lRndRng \ 2

                Select Case mlEmitterMod
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusX  'vertical (y-z) emitter going negative X
                        mfXLocMax = 0 : mfXLocMin = 0
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = -1.0F : mfXSpeedMin = 0.0F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = -0.03F : mfXAccMin = 0.0F
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusY  'horizontal (x-z) emitter going negative Y
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = 0 : mfYLocMin = 0
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = -1.0F : mfYSpeedMin = 0.0F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = -0.03F : mfYAccMin = 0.0F
                        mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusZ  'vertical (x-y) emitter going negative Z
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = 0 : mfZLocMin = 0
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = -1.0F : mfZSpeedMin = 0.0F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = -0.03F : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_PlusX   'vertical (y-z) emitter going positive X
                        mfXLocMax = 0 : mfXLocMin = 0
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 1.0F : mfXSpeedMin = 0.0F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = 0.03F : mfXAccMin = 0.0F
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_PlusZ   'vertical (x-y) emitter going positive Z
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = 0 : mfZLocMin = 0
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 1.0F : mfZSpeedMin = 0.0F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = 0 : mfYAccMin = 0
                        mfZAccMax = 0.03F : mfZAccMin = 0
                    Case Else                                               'horizontal (x-z) emitter going positive Y
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = 0 : mfYLocMin = 0
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        mfYSpeedMax = 1.0F : mfYSpeedMin = 0.0F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        mfXAccMax = 0 : mfXAccMin = 0
                        mfYAccMax = 0.03F : mfYAccMin = 0.0F
                        mfZAccMax = 0 : mfZAccMin = 0
                End Select
            End Set
        End Property

        Public Sub New(ByVal oVecLoc As Vector3, ByVal oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32, ByVal lActPCnt As Int32)
            MyBase.New(oVecLoc, oParentEngine, lParticleCnt, lActPCnt)
        End Sub

        Protected Overrides Sub ResetParticle(ByVal lIndex As Int32)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                MyBase.moParticles(lIndex).mfA = 0
                MyBase.moParticles(lIndex).ParticleActive = False
                MyBase.moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                MyBase.EmitterStopped = True
                For lX = 0 To MyBase.mlCurrentParticleUB
                    If MyBase.moParticles(lX).ParticleActive = True Then
                        MyBase.EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                Dim fX As Single = mvecEmitter.X + ((Rnd() * mfXLocMax) - mfXLocMin)
                Dim fY As Single = mvecEmitter.Y + ((Rnd() * mfYLocMax) - mfYLocMin)
                Dim fZ As Single = mvecEmitter.Z + ((Rnd() * mfZLocMax) - mfZLocMin)
                Dim fXS As Single = (Rnd() * mfXSpeedMax) - mfXSpeedMin
                Dim fYS As Single = (Rnd() * mfYSpeedMax) - mfYSpeedMin
                Dim fZS As Single = (Rnd() * mfZSpeedMax) - mfZSpeedMin
                Dim fXA As Single = (Rnd() * mfXAccMax) - mfXAccMin
                Dim fYA As Single = (Rnd() * mfYAccMax) - mfYAccMin
                Dim fZA As Single = (Rnd() * mfZAccMax) - mfZAccMin

                RotateVals(fX, fY, fZ, fXS, fZS, fXA, fZA)

                If MyBase.AttachedObject Is Nothing = False Then
                    'fX += AttachedObject.LocX
                    fX += AttachedObject.LocX
                    fY += AttachedObject.LocY
                    fZ += AttachedObject.LocZ

                    With CType(AttachedObject, BaseEntity)
                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                            fXS += mfAO_VX
                            fZS += mfAO_VZ
                        End If
                    End With
                End If

                If Rnd() * 100 < 40 Then
                    moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 1, 1, 1, (Rnd() * 64) + 190)
                Else
                    moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 255, 128, 50, (0.6F + (0.2F * Rnd())) * 255)
                End If

                'moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 255, 128, 50, (0.6F + (0.2F * GetNxtRnd())) * 255)
                moParticles(lIndex).fAChg = -((0.01F + Rnd() * 0.05F) * 255)
            End If

        End Sub

        Public Overrides Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)
            'Dim matWorld As Matrix

            If MyBase.EmitterStopped = True Then Exit Sub
            'mlParticleUB = 9
            Try

                Dim blCurrMS As Int64 = msw_Timer.ElapsedMilliseconds
                If mblPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (blCurrMS - mblPrevFrame) * 0.033F
                If fElapsed > 1.0F Then fElapsed = 1.0F
                mblPrevFrame = blCurrMS
                'If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
                'If fElapsed > 1.0F Then fElapsed = 1.0F
                'mlPrevFrame = timeGetTime

                Dim bSetInvisColor As Boolean = MyBase.AttachedObject Is Nothing OrElse MyBase.AttachedObject.yVisibility <> eVisibilityType.Visible
                'Dim lInvisColor As Int32 = FXEmitter.moInvisColor.ToArgb

                mfAO_VX = 0
                mfAO_VZ = 0
                If Me.AttachedObject Is Nothing = False Then
                    With CType(Me.AttachedObject, BaseEntity)
                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                            Dim fTotal As Single = Math.Abs(.VelX) + Math.Abs(.VelZ)
                            mfAO_VX = (.VelX / fTotal) * .TotalVelocity
                            mfAO_VZ = (.VelZ / fTotal) * .TotalVelocity
                        End If
                    End With
                End If

                'Update the particles
                For X = 0 To mlCurrentParticleUB
                    With moParticles(X)
                        .Update(fElapsed)
                        If .mfA <= 0 Then
                            ResetParticle(X)
                        End If

                        If bSetInvisColor = True Then
                            MyBase.moPoints(X).Color = mlInvisColor
                        Else
                            MyBase.moPoints(X).Color = .ParticleColor.ToArgb
                        End If

                        'MyBase.moPoints(X).Color = FXEmitter.moInvisColor.ToArgb
                        'If MyBase.AttachedObject Is Nothing = False Then
                        '    If MyBase.AttachedObject.yVisibility = eVisibilityType.Visible Then
                        '        MyBase.moPoints(X).Color = .ParticleColor.ToArgb
                        '    End If
                        'End If
                        MyBase.moPoints(X).Position = .vecLoc
                    End With
                Next X
            Catch
                'do nothing
            End Try
        End Sub

        Private Sub RotateVals(ByRef fXLoc As Single, ByRef fYLoc As Single, ByRef fZLoc As Single, ByRef fXSpeed As Single, ByRef fZSpeed As Single, ByRef fXAcc As Single, ByRef fZAcc As Single)
            'If True = True Then Return
            If MyBase.AttachedObject Is Nothing Then Return

            Dim matWorld As Matrix = MyBase.AttachedObject.GetWorldMatrix()

            With matWorld
                Dim fX As Single = fXLoc
                Dim fY As Single = fYLoc
                Dim fZ As Single = fZLoc
                fXLoc = (fX * .M11) + (fY * .M21) + (fZ * .M31) '+ .M41
                fYLoc = (fX * .M12) + (fY * .M22) + (fZ * .M32) '+ .M42
                fZLoc = (fX * .M13) + (fY * .M23) + (fZ * .M33) '+ .M43
                'fXLoc = fX : fYLoc = fY : fZLoc = fZ

                'fX = fXSpeed
                ''fY = fYSpeed
                'fZ = fZSpeed
                'fXSpeed = (fX * .M11) + (fY * .M21) + (fZ * .M31) '+ .M41
                ''fYSpeed = (fX * .M12) + (fY * .M22) + (fZ * .M32) '+ .M42
                'fZSpeed = (fX * .M13) + (fY * .M23) + (fZ * .M33) '+ .M43
                ''fX = (fXSpeed * .M11) + (fYSpeed * .M21) + (fZSpeed * .M31)
                ''fY = (fXSpeed * .M12) + (fYSpeed * .M22) + (fZSpeed * .M32)
                ''fZ = (fXSpeed * .M13) + (fYSpeed * .M23) + (fZSpeed * .M33)
                ''fXSpeed = fX : fYSpeed = fY : fZSpeed = fZ

                'fX = fXAcc
                'fZ = fZAcc
                'fXAcc = (fX * .M11) + (fZ * .M31)
                'fZAcc = (fX * .M13) + (fZ * .M33)
                ''fXAcc = fX : fZAcc = fZ
            End With
            Dim fRads As Single = -((AttachedObject.LocYaw / 10.0F)) * gdRadPerDegree
            Dim fCosR As Single = CSng(Math.Cos(fRads))
            Dim fSinR As Single = CSng(Math.Sin(fRads))

            ''Yaw...
            'fDX = fXLoc
            'fDZ = fYLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Now set up for standard rotation...
            'fRads = ((AttachedObject.LocAngle / 10.0F) - 90) * gdRadPerDegree
            'fCosR = CSng(Math.Cos(fRads))
            'fSinR = CSng(Math.Sin(fRads))

            ''Loc
            'fDX = fXLoc
            'fDZ = fZLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

            'Speed
            Dim fDX As Single
            Dim fDZ As Single

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

    End Class

    Public Class SmokeyFireFX
        Inherits FXEmitter

        Private mfXLocMin As Single
        Private mfYLocMin As Single
        Private mfZLocMin As Single

        Private mfXLocMax As Single
        Private mfYLocMax As Single
        Private mfZLocMax As Single

        Private mfXSpeedMin As Single
        'Private mfYSpeedMin As Single
        Private mfZSpeedMin As Single

        Private mfXSpeedMax As Single
        'Private mfYSpeedMax As Single
        Private mfZSpeedMax As Single

        'Private mfXAccMin As Single
        'Private mfXAccMax As Single

        ''Private mfYAccMin As Single
        ''Private mfYAccMax As Single

        'Private mfZAccMin As Single
        'Private mfZAccMax As Single

        Private mlEmitterMod As ParticleEngine.EmitterType
        Public Property lEmitterMod() As ParticleEngine.EmitterType
            Get
                Return mlEmitterMod
            End Get
            Set(ByVal value As ParticleEngine.EmitterType)
                mlEmitterMod = value

                Dim lRndRng As Int32 = Me.ParticleCount \ 3
                Dim lHlfRndRng As Int32 = lRndRng \ 2

                Select Case mlEmitterMod
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusX  'vertical (y-z) emitter going negative X
                        mfXLocMax = 0 : mfXLocMin = 0
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = -1.0F : mfXSpeedMin = 0.0F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        'mfXAccMax = -0.03F : mfXAccMin = 0.0F
                        'mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusY  'horizontal (x-z) emitter going negative Y
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = 0 : mfYLocMin = 0
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        'mfYSpeedMax = -1.0F : mfYSpeedMin = 0.0F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        'mfXAccMax = 0 : mfXAccMin = 0
                        'mfYAccMax = -0.03F : mfYAccMin = 0.0F
                        'mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_MinusZ  'vertical (x-y) emitter going negative Z
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = 0 : mfZLocMin = 0
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        'mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = -1.0F : mfZSpeedMin = 0.0F
                        'mfXAccMax = 0 : mfXAccMin = 0
                        'mfYAccMax = 0 : mfYAccMin = 0
                        'mfZAccMax = -0.03F : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_PlusX   'vertical (y-z) emitter going positive X
                        mfXLocMax = 0 : mfXLocMin = 0
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 1.0F : mfXSpeedMin = 0.0F
                        'mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        'mfXAccMax = 0.03F : mfXAccMin = 0.0F
                        'mfYAccMax = 0 : mfYAccMin = 0
                        'mfZAccMax = 0 : mfZAccMin = 0
                    Case ParticleEngine.EmitterType.eFireEmitterMod_PlusZ   'vertical (x-y) emitter going positive Z
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = lRndRng : mfYLocMin = lHlfRndRng
                        mfZLocMax = 0 : mfZLocMin = 0
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        ' mfYSpeedMax = 0.8F : mfYSpeedMin = 0.4F
                        mfZSpeedMax = 1.0F : mfZSpeedMin = 0.0F
                        'mfXAccMax = 0 : mfXAccMin = 0
                        'mfYAccMax = 0 : mfYAccMin = 0
                        'mfZAccMax = 0.03F : mfZAccMin = 0
                    Case Else                                               'horizontal (x-z) emitter going positive Y
                        mfXLocMax = lRndRng : mfXLocMin = lHlfRndRng
                        mfYLocMax = 0 : mfYLocMin = 0
                        mfZLocMax = lRndRng : mfZLocMin = lHlfRndRng
                        mfXSpeedMax = 0.8F : mfXSpeedMin = 0.4F
                        'mfYSpeedMax = 1.0F : mfYSpeedMin = 0.0F
                        mfZSpeedMax = 0.8F : mfZSpeedMin = 0.4F
                        'mfXAccMax = 0 : mfXAccMin = 0
                        ' mfYAccMax = 0.03F : mfYAccMin = 0.0F
                        'mfZAccMax = 0 : mfZAccMin = 0
                End Select
            End Set
        End Property

        Public Sub New(ByVal oVecLoc As Vector3, ByVal oParentEngine As ParticleEngine, ByVal lParticleCnt As Int32, ByVal lActPCnt As Int32)
            MyBase.New(oVecLoc, oParentEngine, lParticleCnt, lActPCnt)
        End Sub

        Protected Overrides Sub ResetParticle(ByVal lIndex As Int32)
            Dim lX As Int32

            If mbEmitterStopping = True Then
                MyBase.moParticles(lIndex).mfA = 0
                MyBase.moParticles(lIndex).ParticleActive = False
                MyBase.moPoints(lIndex).Color = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb

                MyBase.EmitterStopped = True
                For lX = 0 To MyBase.mlCurrentParticleUB
                    If MyBase.moParticles(lX).ParticleActive = True Then
                        MyBase.EmitterStopped = False
                        Exit For
                    End If
                Next lX
            Else
                Dim fX As Single = mvecEmitter.X + ((Rnd() * mfXLocMax) - mfXLocMin)
                Dim fY As Single = mvecEmitter.Y + ((Rnd() * mfYLocMax) - mfYLocMin)
                Dim fZ As Single = mvecEmitter.Z + ((Rnd() * mfZLocMax) - mfZLocMin)
                Dim fXS As Single = (Rnd() * mfXSpeedMax) - mfXSpeedMin
                Dim fZS As Single = (Rnd() * mfZSpeedMax) - mfZSpeedMin
                Dim fYS As Single = Rnd() * (Math.Max(Math.Abs(fXS), Math.Abs(fZS)) + 1.0F)
                'fYS *= 5

                'Dim fYS As Single = (GetNxtRnd() * mfYSpeedMax) - mfYSpeedMin
                'Dim fXA As Single = (GetNxtRnd() * mfXAccMax) - mfXAccMin
                'Dim fYA As Single = (GetNxtRnd() * mfYAccMax) - mfYAccMin
                'Dim fYA As Single = -0.001F
                'Dim fZA As Single = (GetNxtRnd() * mfZAccMax) - mfZAccMin

                RotateVals(fX, fY, fZ, fXS, fZS) ', fXA, fZA)

                If MyBase.AttachedObject Is Nothing = False Then
                    'fX += AttachedObject.LocX
                    fX += AttachedObject.LocX
                    fY += AttachedObject.LocY
                    fZ += AttachedObject.LocZ
                End If

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, 0, -0.001F, 0, 255, 128, Rnd() * 50, (0.8F + (0.2F * Rnd())) * 255)
                moParticles(lIndex).fAChg = ((0.01F - (Rnd() * 0.005F)) * -255) '* 10
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

            If MyBase.EmitterStopped = True Then Exit Sub

            Try

                Dim blCurrMS As Int64 = msw_Timer.ElapsedMilliseconds
                If mblPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (blCurrMS - mblPrevFrame) * 0.033F
                If fElapsed > 1.0F Then fElapsed = 1.0F
                mblPrevFrame = blCurrMS
                'If mlPrevFrame = 0 Then fElapsed = 1 Else fElapsed = (timeGetTime - mlPrevFrame) / 30.0F
                'If fElapsed > 1.0F Then fElapsed = 1.0F
                'mlPrevFrame = timeGetTime

                'Update the particles
                For X = 0 To mlCurrentParticleUB
                    With moParticles(X)
                        .Update(fElapsed)
                        If .mfA <= 0 Then
                            ResetParticle(X)
                        End If

                        'check our limits
                        If .mfR < 64 Then .mfR = 64
                        If .mfG < 64 Then .mfG = 64
                        If .mfB > 64 Then .mfB = 64

                        MyBase.moPoints(X).Color = mlInvisColor 'FXEmitter.moInvisColor.ToArgb
                        If MyBase.AttachedObject Is Nothing = False Then
                            If MyBase.AttachedObject.yVisibility = eVisibilityType.Visible Then
                                MyBase.moPoints(X).Color = .ParticleColor.ToArgb
                            End If
                        End If
                        MyBase.moPoints(X).Position = .vecLoc
                    End With
                Next X
            Catch
                'do nothing
            End Try
        End Sub

        Private Sub RotateVals(ByRef fXLoc As Single, ByRef fYLoc As Single, ByRef fZLoc As Single, ByRef fXSpeed As Single, ByRef fZSpeed As Single)
            'Dim fDX As Single
            'Dim fDZ As Single
            ''If True = True Then Return
            'If MyBase.AttachedObject Is Nothing Then Return
            'Dim fRads As Single = -((AttachedObject.LocYaw / 10.0F)) * gdRadPerDegree
            'Dim fCosR As Single = CSng(Math.Cos(fRads))
            'Dim fSinR As Single = CSng(Math.Sin(fRads))

            ''Yaw...
            'fDX = fXLoc
            'fDZ = fYLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Now set up for standard rotation...
            'fRads = ((AttachedObject.LocAngle / 10.0F) - 90) * gdRadPerDegree
            'fCosR = CSng(Math.Cos(fRads))
            'fSinR = CSng(Math.Sin(fRads))

            ''Loc
            'fDX = fXLoc
            'fDZ = fZLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Speed
            'fDX = fXSpeed
            'fDZ = fZSpeed
            'fXSpeed = +((fDX * fCosR) + (fDZ * fSinR))
            'fZSpeed = -((fDX * fSinR) - (fDZ * fCosR))

            ' ''Acc
            ''fDX = fXAcc
            ''fDZ = fZAcc
            ''fXAcc = +((fDX * fCosR) + (fDZ * fSinR))
            ''fZAcc = -((fDX * fSinR) - (fDZ * fCosR))
            If MyBase.AttachedObject Is Nothing Then Return

            Dim matWorld As Matrix = MyBase.AttachedObject.GetWorldMatrix()

            With matWorld
                Dim fX As Single = fXLoc
                Dim fY As Single = fYLoc
                Dim fZ As Single = fZLoc
                fXLoc = (fX * .M11) + (fY * .M21) + (fZ * .M31) '+ .M41
                fYLoc = (fX * .M12) + (fY * .M22) + (fZ * .M32) '+ .M42
                fZLoc = (fX * .M13) + (fY * .M23) + (fZ * .M33) '+ .M43
                'fXLoc = fX : fYLoc = fY : fZLoc = fZ

                'fX = fXSpeed
                ''fY = fYSpeed
                'fZ = fZSpeed
                'fXSpeed = (fX * .M11) + (fY * .M21) + (fZ * .M31) '+ .M41
                ''fYSpeed = (fX * .M12) + (fY * .M22) + (fZ * .M32) '+ .M42
                'fZSpeed = (fX * .M13) + (fY * .M23) + (fZ * .M33) '+ .M43
                ''fX = (fXSpeed * .M11) + (fYSpeed * .M21) + (fZSpeed * .M31)
                ''fY = (fXSpeed * .M12) + (fYSpeed * .M22) + (fZSpeed * .M32)
                ''fZ = (fXSpeed * .M13) + (fYSpeed * .M23) + (fZSpeed * .M33)
                ''fXSpeed = fX : fYSpeed = fY : fZSpeed = fZ

                'fX = fXAcc
                'fZ = fZAcc
                'fXAcc = (fX * .M11) + (fZ * .M31)
                'fZAcc = (fX * .M13) + (fZ * .M33)
                ''fXAcc = fX : fZAcc = fZ
            End With
            Dim fRads As Single = -((AttachedObject.LocYaw / 10.0F)) * gdRadPerDegree
            Dim fCosR As Single = CSng(Math.Cos(fRads))
            Dim fSinR As Single = CSng(Math.Sin(fRads))

            ''Yaw...
            'fDX = fXLoc
            'fDZ = fYLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fYLoc = -((fDX * fSinR) - (fDZ * fCosR))

            ''Now set up for standard rotation...
            'fRads = ((AttachedObject.LocAngle / 10.0F) - 90) * gdRadPerDegree
            'fCosR = CSng(Math.Cos(fRads))
            'fSinR = CSng(Math.Sin(fRads))

            ''Loc
            'fDX = fXLoc
            'fDZ = fZLoc
            'fXLoc = +((fDX * fCosR) + (fDZ * fSinR))
            'fZLoc = -((fDX * fSinR) - (fDZ * fCosR))

            'Speed
            Dim fDX As Single
            Dim fDZ As Single

            fDX = fXSpeed
            fDZ = fZSpeed
            fXSpeed = +((fDX * fCosR) + (fDZ * fSinR))
            fZSpeed = -((fDX * fSinR) - (fDZ * fCosR))

        End Sub

    End Class

End Namespace
