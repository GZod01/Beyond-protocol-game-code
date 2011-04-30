Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Namespace CityCars

    Public Class Particle
        Public vecLoc As Vector3
        Public vecSpeed As Vector3

        Public ParticleColor As System.Drawing.Color

        Public oFac1 As BaseEntity
        Public oFac2 As BaseEntity
        Public yTarget As Byte

        Public Sub Update(ByVal fElapsedTime As Single)
            vecLoc.Add(Vector3.Multiply(vecSpeed, fElapsedTime))


            If yTarget = 0 Then
                If Math.Abs(oFac1.LocX - vecLoc.X) < 150 AndAlso Math.Abs(oFac1.LocZ - vecLoc.Z) < 150 Then
                    yTarget = 1
                    vecSpeed.Multiply(-1)
                End If
            Else
                If Math.Abs(oFac2.LocX - vecLoc.X) < 150 AndAlso Math.Abs(oFac2.LocZ - vecLoc.Z) < 150 Then
                    yTarget = 0
                    vecSpeed.Multiply(-1)
                End If
            End If

            vecLoc.Y = CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(vecLoc.X, vecLoc.Z, True) + 125
        End Sub

    End Class

    Public Class CityCars
        Private moChildren() As Particle
        Private moPoints() As CustomVertex.PositionColored
        Public mlChildrenUB As Int32 = -1

        'The standard texture to use for emitters using this engine object
        Private moTex As Texture
        Private mfParticleSize As Single
        'Private msw As Stopwatch
        Public Sub New(ByVal fParticleSize As Single)
            mfParticleSize = fParticleSize
        End Sub

        Public Sub Render(ByVal bUpdateNoRender As Boolean)
            'Dim matWorld As Matrix
            Dim X As Int32

            On Error Resume Next

            If moTex Is Nothing OrElse moTex.Disposed = True Then
                moTex = goResMgr.GetTexture("Particle.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "textures.pak")
            End If

            If mlChildrenUB = -1 Then Return
            'If msw Is Nothing Then msw = Stopwatch.StartNew()
            Dim fElapsed As Single = 1 'msw.ElapsedMilliseconds / 30.0F
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
                    For X = 0 To mlChildrenUB
                        'Now, check for whether the emitter can be seen
                        moChildren(X).Update(fElapsed)
                        moPoints(X).Color = moChildren(X).ParticleColor.ToArgb
                        moPoints(X).Position = moChildren(X).vecLoc
                    Next X
                    .DrawUserPrimitives(PrimitiveType.PointList, mlChildrenUB + 1, moPoints)

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

        Public Sub AddParticle(ByVal oFac1 As BaseEntity, ByVal oFac2 As BaseEntity)
            mlChildrenUB += 1
            ReDim Preserve moChildren(mlChildrenUB)
            ReDim Preserve moPoints(mlChildrenUB)

            moChildren(mlChildrenUB) = New Particle
            With moChildren(mlChildrenUB)
                .vecLoc = New Vector3(oFac1.LocX, oFac1.LocY + 500, oFac1.LocZ)
                .vecSpeed = New Vector3(0, 0, 0)

                .vecSpeed = New Vector3(oFac2.LocX - oFac1.LocX, oFac2.LocY - oFac1.LocY, oFac2.LocZ - oFac1.LocZ)
                .vecSpeed.Normalize()
                .vecSpeed.Multiply((Rnd() * 7.0F) + 5) ') 12.0F)

                .ParticleColor = System.Drawing.Color.FromArgb(255, 128 + CInt((Rnd() * 64) - 32), 255 - CInt(Rnd() * 64), 255 - CInt(Rnd() * 64))
                .oFac1 = oFac1
                .oFac2 = oFac2
                .yTarget = 1
            End With
        End Sub

    End Class

End Namespace
