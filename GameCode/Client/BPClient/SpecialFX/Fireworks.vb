Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class FireworksMgr
    Private Const ml_MAX_SIMULTANEOUS_FIREWORKS As Int32 = 6
    Private Class Particle
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

        Public fMinR As Single
        Public fMinG As Single
        Public fMinB As Single

        Public Sub Update(ByVal fElapsedTime As Single)
            vecLoc.Add(Vector3.Multiply(vecSpeed, fElapsedTime))

            vecSpeed.Add(Vector3.Multiply(vecAcc, fElapsedTime))

            mfA += (fAChg * fElapsedTime)
            mfR += (fRChg * fElapsedTime)
            mfG += (fGChg * fElapsedTime)
            mfB += (fBChg * fElapsedTime)

            If mfA < 0 Then mfA = 0
            If mfA > 255 Then mfA = 255
            If mfR < fMinR Then mfR = fMinR
            If mfR > 255 Then mfR = 255
            If mfG < fMinG Then mfG = fMinG
            If mfG > 255 Then mfG = 255
            If mfB < fMinB Then mfB = fMinB
            If mfB > 255 Then mfB = 255

            ParticleColor = System.Drawing.Color.FromArgb(CInt(mfA), CInt(mfR), CInt(mfG), CInt(mfB))
        End Sub

        Public Sub Reset(ByVal fX As Single, ByVal fY As Single, ByVal fZ As Single, ByVal fXSpeed As Single, ByVal fYSpeed As Single, ByVal fZSpeed As Single, ByVal fXAcc As Single, ByVal fYAcc As Single, ByVal fZAcc As Single, ByVal fR As Single, ByVal fG As Single, ByVal fB As Single, ByVal fA As Single)
            vecLoc.X = fX : vecLoc.Y = fY : vecLoc.Z = fZ
            vecAcc.X = fXAcc : vecAcc.Y = fYAcc : vecAcc.Z = fZAcc
            vecSpeed.X = fXSpeed : vecSpeed.Y = fYSpeed : vecSpeed.Z = fZSpeed
            mfA = fA : mfR = fR : mfG = fG : mfB = fB

            fMinR = 0.0F : fMinG = 0.0F : fMinB = 0.0F

            ParticleActive = True
        End Sub

    End Class

    Private MustInherit Class Firework
        Public vecLoc As Vector3
        Protected vecSpeed As Vector3
        Protected vecAcc As Vector3

		Protected oSW As Stopwatch
        Public mlParticleUB As Int32 = -1
        Protected moParticles() As Particle
        Public moPoints() As CustomVertex.PositionColoredTextured

        Public FireworkStopped As Boolean = False

        Protected MustOverride Sub ResetParticle(ByVal lIndex As Int32)
        Protected MustOverride Sub FireworkExploded()
        Protected MustOverride Sub ExplodedUpdate(ByVal fElapsed As Single)

        Private mbExploded As Boolean = False
        Private mlExplosionHeight As Int32 = 16000

        Private mlLightIndex As Int32 = -1
        Private mfLightR As Single = 0
        Private mfLightG As Single = 0
        Private mfLightB As Single = 0

        Public Sub Launch(ByVal vecFromLoc As Vector3, ByVal vecLaunchSpeed As Vector3, ByVal lExplosionHeight As Int32)
            mlExplosionHeight = lExplosionHeight
            vecLoc = vecFromLoc
            vecSpeed = vecLaunchSpeed
            vecAcc.X = 0 : vecAcc.Y = -0.0098F : vecAcc.Z = 0

            For X As Int32 = 0 To mlParticleUB
                ResetLaunchParticle(X)
            Next X
        End Sub

        Public Sub Update()
            Dim fElapsed As Single
            Dim X As Int32
            Dim vecCenter As Vector3 = New Vector3(16, 16, 0)

            If FireworkStopped = True Then Return

            If oSW Is Nothing Then
                oSW = Stopwatch.StartNew()
                fElapsed = 1.0F
            Else
                fElapsed = oSW.ElapsedMilliseconds / 30.0F
                oSW.Stop()
                oSW.Reset()
                oSW.Start()
            End If

            vecSpeed.Add(Vector3.Scale(vecAcc, fElapsed))
            vecLoc.Add(vecSpeed)
            If vecLoc.Y > mlExplosionHeight AndAlso mbExploded = False Then
                mbExploded = True

                If goSound Is Nothing = False Then
                    Dim sName As String = "Explosions\HullHit2.wav"
                    Select Case CInt(Rnd() * 6)
                        Case 0
                            sName = "Explosions\FlakExplosion.wav"
                        Case 1
                            sName = "Explosions\HullHit2.wav"
                        Case 2
                            sName = "Explosions\HullHit3.wav"
                        Case 3
                            sName = "Explosions\SmallGroundDeath1.wav"
                        Case 4
                            sName = "Explosions\SmallSpaceDeath2.wav"
                        Case Else
                            sName = "Projectile Weapons\Artillery2.wav"
                    End Select

                    goSound.StartSound(sName, False, SoundMgr.SoundUsage.eFireworks, vecLoc, New Vector3(0, 0, 0))
                End If

                FireworkExploded()
            End If

            If mlLightIndex <> -1 AndAlso FireworksMgr.yLightUsed(mlLightIndex) <> 0 Then
                Dim fAdjust As Single = 3.0F * fElapsed
                mfLightR -= fAdjust
                mfLightG -= fAdjust
                mfLightB -= fAdjust
                If mfLightR < 0 Then mfLightR = 0
                If mfLightG < 0 Then mfLightG = 0
                If mfLightB < 0 Then mfLightB = 0
                If (mfLightR = 0 AndAlso mfLightG = 0 AndAlso mfLightB = 0) OrElse FireworkStopped = True Then
                    FireworksMgr.yLightUsed(mlLightIndex) = 0
                    mlLightIndex = -1
                Else
                    With FireworksMgr.Lights(mlLightIndex)
                        .Diffuse = System.Drawing.Color.FromArgb(255, CInt(mfLightR), CInt(mfLightG), CInt(mfLightB))
                        .Specular = .Diffuse
                        '.Update()
                    End With
                End If
            End If

            'Update the particles
            For X = 0 To mlParticleUB
                With moParticles(X)
                    If mbExploded = True Then
                        ExplodedUpdate(fElapsed)
                        If FireworkStopped = True Then
                            If mlLightIndex <> -1 Then
                                FireworksMgr.yLightUsed(mlLightIndex) = 0
                                mlLightIndex = -1
                            End If
                        End If
                    End If

                    .Update(fElapsed)
                    If .mfA <= 0 Then
                        If mbExploded = True Then ResetParticle(X) Else ResetLaunchParticle(X)
                    End If
                    moPoints(X).Color = .ParticleColor.ToArgb
                    moPoints(X).Position = .vecLoc
                End With
            Next X
        End Sub

        Private Sub ResetLaunchParticle(ByVal lIndex As Int32)
            If lIndex < mlParticleUB \ 3 Then
                Dim vecTemp As Vector3 = Vector3.Multiply(vecSpeed, -1.0F)
                Dim fXS As Single = Rnd() * vecTemp.X
                Dim fYS As Single = Rnd() * vecTemp.Y
                Dim fZS As Single = Rnd() * vecTemp.Z

                With moParticles(lIndex)
                    .Reset(vecLoc.X, vecLoc.Y, vecLoc.Z, fXS, fYS, fZS, 0, -0.0098F, 0, 255, 255, 255, 255)
                    .fRChg = -0.001F * Rnd()
                    .fGChg = -0.01F * Rnd()
                    .fBChg = -0.01F * Rnd()

                    .fMinR = 32
                    .fMinG = 32
                    .fMinB = 32
                End With
            End If
            moParticles(lIndex).fAChg = -((0.1F + (0.1F * Rnd())) * 255)
        End Sub

        Public Sub New()
            Dim lPCnt As Int32 = 300
            Select Case muSettings.BurnFXParticles
                Case 1      '25%
                    lPCnt = CInt(lPCnt * 0.25F)
                Case 2      '50%
                    lPCnt = CInt(lPCnt * 0.5F)
                Case 3      '75%
                    lPCnt = CInt(lPCnt * 0.75F)
            End Select

            mlParticleUB = lPCnt - 1
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB)

            For X As Int32 = 0 To mlParticleUB
                moParticles(X) = New Particle
            Next X
        End Sub

        Protected Shared Sub GetColorFromColorVal(ByVal lColorVal As Int32, ByRef lR As Int32, ByRef lG As Int32, ByRef lB As Int32)
            Select Case lColorVal
                Case 0
                    lR = CInt(Rnd() * 210 + 32)
                    lG = 32
                    lB = 32
                Case 1
                    lR = 32
                    lG = CInt(Rnd() * 210 + 32)
                    lB = 32
                Case 2
                    lR = 32
                    lG = 32
                    lB = CInt(Rnd() * 210 + 32)
                Case 3
                    lR = 32
                    lG = CInt(Rnd() * 210 + 32)
                    lB = CInt(Rnd() * 210 + 32)
                Case 4
                    lR = CInt(Rnd() * 210 + 32)
                    lG = 32
                    lB = CInt(Rnd() * 210 + 32)
                Case 5
                    lR = CInt(Rnd() * 210 + 32)
                    lG = CInt(Rnd() * 210 + 32)
                    lB = 32
                Case 6
                    lR = CInt(Rnd() * 210 + 32)
                    lG = CInt(Rnd() * 210 + 32)
                    lB = CInt(Rnd() * 210 + 32)
            End Select
        End Sub

        Protected Sub AddLightEffect(ByVal yColorVal As Byte)
            Dim lR As Int32 = 0
            Dim lG As Int32 = 0
            Dim lB As Int32 = 0
            GetColorFromColorVal(yColorVal, lR, lG, lB)

            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To FireworksMgr.lLightUB
                If FireworksMgr.yLightUsed(X) = 0 Then
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx = -1 Then
                lIdx = FireworksMgr.lLightUB + 1
                ReDim Preserve FireworksMgr.yLightUsed(lIdx)
                ReDim Preserve FireworksMgr.Lights(lIdx)
                FireworksMgr.lLightUB = lIdx
            End If

            FireworksMgr.Lights(lIdx) = New Light()
            With FireworksMgr.Lights(lIdx)
                .Ambient = Color.Black
                .Diffuse = System.Drawing.Color.FromArgb(255, lR, lG, lB)
                .Position = vecLoc
                .Specular = Color.White
                .Type = LightType.Point
                .Attenuation0 = 1
                .Attenuation1 = 0
                .Attenuation2 = 0
                .Range = 20000
                .Falloff = 0.3F
                '.Update()
            End With
            FireworksMgr.yLightUsed(lIdx) = 255

            mlLightIndex = lIdx
            mfLightR = lR
            mfLightG = lG
            mfLightB = lB
        End Sub
    End Class

    Private Class DiskBurst
        Inherits Firework

        Private Structure DiskPack
            Public vecLoc As Vector3
            Public vecSpeed As Vector3
        End Structure
        Private muPacks() As DiskPack
        Private myColor As Byte = 0

        Protected Overrides Sub ResetParticle(ByVal lIndex As Integer)
            If muPacks Is Nothing Then Return

            Dim lVal As Int32 = lIndex Mod muPacks.Length
            With muPacks(lVal)
                Dim lR As Int32 = 0
                Dim lG As Int32 = 0
                Dim lB As Int32 = 0
                Firework.GetColorFromColorVal(myColor, lR, lG, lB)
                moParticles(lIndex).Reset(.vecLoc.X, .vecLoc.Y, .vecLoc.Z, .vecSpeed.X / 10.0F, .vecSpeed.Y / 10.0F, .vecSpeed.Z / 10.0F, 0, -0.0098F, 0, lR, lG, lB, 255)
            End With

            With moParticles(lIndex)
                '.fAChg = -((0.1F + (0.1F * GetNxtRnd())) * 255)
                .fAChg = -((0.075F + (0.075F * Rnd())) * 255)
                .fRChg = -((0.1F * Rnd()) + 0.05F)
                .fGChg = -((0.1F * Rnd()) + 0.05F)
                .fBChg = -((0.1F * Rnd()) + 0.05F)
            End With
        End Sub

        Protected Overrides Sub FireworkExploded()
            ReDim muPacks(19)

            Dim fRotYAxis As Single = (Rnd() * 90) - 45
            myColor = CByte(Rnd() * 6)

            MyBase.AddLightEffect(myColor)

            For X As Int32 = 0 To muPacks.GetUpperBound(0)
                Dim fX As Single = 0.1F
                Dim fZ As Single = 0
                Dim fY As Single = 0
                RotatePoint(0, 0, fX, fZ, 18 * X + ((Rnd() * 3) - 1.5F))

                RotatePoint(0, 0, fX, fY, fRotYAxis)
                RotatePoint(0, 0, fY, fZ, fRotYAxis)

                muPacks(X).vecLoc = vecLoc
                With muPacks(X).vecSpeed
                    .X = fX
                    .Y = fY
                    .Z = fZ
                End With
            Next X
        End Sub

        Protected Overrides Sub ExplodedUpdate(ByVal fElapsed As Single)
            If muPacks Is Nothing = False Then
                Me.FireworkStopped = True
                For X As Int32 = 0 To muPacks.GetUpperBound(0)
                    muPacks(X).vecSpeed.Add(New Vector3(0, -0.00000298F, 0))
                    'If muPacks(X).vecSpeed.Y < -1.0F Then muPacks(X).vecSpeed.Y = -1.0F
                    muPacks(X).vecLoc.Add(Vector3.Scale(muPacks(X).vecSpeed, fElapsed))

                    If muPacks(X).vecLoc.Y > 0 Then Me.FireworkStopped = False
                Next X
            End If
        End Sub
    End Class

    Private Class NormalBurst
        Inherits Firework

        Private Structure NormalPack
            Public vecLoc As Vector3
            Public vecSpeed As Vector3
        End Structure
        Private muPacks() As NormalPack

        Private myColor As Byte = 0

        Protected Overrides Sub ResetParticle(ByVal lIndex As Integer)
            If muPacks Is Nothing Then Return

            Dim lVal As Int32 = lIndex Mod muPacks.Length
            With muPacks(lVal)
                Dim lR As Int32 = 0
                Dim lG As Int32 = 0
                Dim lB As Int32 = 0

                Firework.GetColorFromColorVal(myColor, lR, lG, lB)

                moParticles(lIndex).Reset(.vecLoc.X, .vecLoc.Y, .vecLoc.Z, .vecSpeed.X / 10.0F, .vecSpeed.Y / 10.0F, .vecSpeed.Z / 10.0F, 0, -0.0098F, 0, lR, lG, lB, 255)
            End With

            With moParticles(lIndex)
                '.fAChg = -((0.1F + (0.1F * GetNxtRnd())) * 255)
                .fAChg = -((0.075F + (0.075F * Rnd())) * 255)
                .fRChg = -((0.1F * Rnd()) + 0.05F)
                .fGChg = -((0.1F * Rnd()) + 0.05F)
                .fBChg = -((0.1F * Rnd()) + 0.05F)
            End With
        End Sub

        Protected Overrides Sub FireworkExploded()
            'Dim lPackCnt As Int32 = CInt(GetNxtRnd() * 20 + 30)
            Dim lPackCnt As Int32 = CInt(Rnd() * 20 + 5)

            ReDim muPacks(lPackCnt)

            myColor = CByte(Rnd() * 6)

            MyBase.AddLightEffect(myColor)

            For X As Int32 = 0 To muPacks.GetUpperBound(0)
                muPacks(X).vecLoc = vecLoc
                With muPacks(X).vecSpeed
                    .X = (Rnd() * 0.1F) - 0.05F
                    .Y = (Rnd() * 0.1F) - 0.05F
                    .Z = (Rnd() * 0.1F) - 0.05F
                End With
            Next X
        End Sub

        Protected Overrides Sub ExplodedUpdate(ByVal fElapsed As Single)
            If muPacks Is Nothing = False Then
                Me.FireworkStopped = True
                For X As Int32 = 0 To muPacks.GetUpperBound(0)
                    muPacks(X).vecSpeed.Add(New Vector3(0, -0.00000298F, 0))
                    'If muPacks(X).vecSpeed.Y < -1.0F Then muPacks(X).vecSpeed.Y = -1.0F
                    muPacks(X).vecLoc.Add(Vector3.Scale(muPacks(X).vecSpeed, fElapsed))

                    If muPacks(X).vecLoc.Y > 0 Then Me.FireworkStopped = False
                Next X
            End If
        End Sub
    End Class

    Private Class SparkleBurst
        Inherits Firework

        Private mbIgnited As Boolean = False
        Private sw_TimeIgnite As Stopwatch

        Private myColor As Byte = 0
        Protected Overrides Sub ResetParticle(ByVal lIndex As Integer)
            If mbIgnited = False Then
                If moParticles(lIndex).mfA <> 0 Then moParticles(lIndex).mfA = 0
            Else
                With moParticles(lIndex)
                    '.Reset(.vecLoc.X, .vecLoc.Y, .vecLoc.Z, 0, 0, 0, 0, -0.098F, 0, 255, 255, 255, 255)
                    .vecSpeed.Y = -9.8F * (Rnd() * 2.0F)
                    Dim lR As Int32 = 0
                    Dim lG As Int32 = 0
                    Dim lB As Int32 = 0
                    Firework.GetColorFromColorVal(myColor, lR, lG, lB)
                    .mfR = lR : .mfG = lG : .mfB = lB : .mfA = 255
                End With
            End If
        End Sub

        Protected Overrides Sub FireworkExploded()
            mbIgnited = False
            sw_TimeIgnite = Stopwatch.StartNew()

            myColor = CByte(Rnd() * 6)
            MyBase.AddLightEffect(myColor)
            For X As Int32 = 0 To mlParticleUB
                ResetParticle(X)

                Dim fX As Single = 1500
                Dim fY As Single = 0
                Dim fZ As Single = 0

                RotatePoint(0, 0, fX, fZ, Rnd() * 360.0F)
                RotatePoint(0, 0, fX, fY, Rnd() * 360.0F)
                RotatePoint(0, 0, fZ, fY, Rnd() * 360.0F)

                moParticles(X).vecLoc.X = vecLoc.X + fX
                moParticles(X).vecLoc.Y = vecLoc.Y + fY
                moParticles(X).vecLoc.Z = vecLoc.Z + fZ
                moParticles(X).vecAcc = New Vector3(0, -0.000098F, 0)
                moParticles(X).vecSpeed = New Vector3(0, 0, 0)
            Next X
        End Sub

        Protected Overrides Sub ExplodedUpdate(ByVal fElapsed As Single)
            If sw_TimeIgnite Is Nothing = False Then
                If sw_TimeIgnite.ElapsedMilliseconds > 300 Then
                    If mbIgnited = False AndAlso goSound Is Nothing = False Then
                        goSound.StartSound("Explosions\sparkle_fireworks.wav", False, SoundMgr.SoundUsage.eFireworks, vecLoc, New Vector3(0, 0, 0))
                    End If
                    mbIgnited = True
                End If
                If sw_TimeIgnite.ElapsedMilliseconds > 3300 Then
                    Me.FireworkStopped = True
                    sw_TimeIgnite.Stop()
                    sw_TimeIgnite = Nothing
                End If
            End If
        End Sub
    End Class

    Private moTex As Texture = Nothing
    Private mlChildrenUB As Int32 = -1
    Private myChildUsed() As Byte
    Private moChildren() As Firework

    Public Shared Lights() As Light
    Public Shared yLightUsed() As Byte
    Public Shared lLightUB As Int32 = -1

    Public Sub ClearTex()
        moTex = Nothing
    End Sub

    Public Sub Render(ByVal bUpdateNoRender As Boolean)
        'Dim matWorld As Matrix
        Dim X As Int32

        On Error Resume Next

        If moTex Is Nothing OrElse moTex.Disposed = True Then
            moTex = goResMgr.GetTexture("ExpParticle2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "Misc.pak")
        End If

        If mlChildrenUB = -1 Then Return
        Dim bFound As Boolean = False
        For X = 0 To mlChildrenUB
            If myChildUsed(X) <> 0 Then
                bFound = True
                Exit For
            End If
        Next X
        If bFound = False Then Return

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
                    If .DeviceCaps.MaxPointSize > 64 Then
                        .RenderState.PointSize = 64
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

                glFireworksRendered = 0
                For X = 0 To mlChildrenUB
                    If myChildUsed(X) > 0 Then
                        'Now, check for whether the emitter can be seen
                        moChildren(X).Update()
                        If moChildren(X).FireworkStopped = True Then
                            myChildUsed(X) = 0
                        Else
                            If goCamera.CullObject(CullBox.GetCullBox(moChildren(X).vecLoc.X, moChildren(X).vecLoc.Y, moChildren(X).vecLoc.Z, -5000, -5000, -5000, 5000, 5000, 5000)) = False Then
                                .DrawUserPrimitives(PrimitiveType.PointList, moChildren(X).mlParticleUB + 1, moChildren(X).moPoints)
                                glFireworksRendered += 1
                            End If
                        End If
                    End If
                Next X

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

    Public Sub AddEmitter(ByVal vecLaunch As Vector3, ByVal yType As Byte, ByVal lExplHeight As Int32)
        Dim X As Int32
        Dim lIdx As Int32 = -1
        Dim lCnt As Int32 = 0
        For X = 0 To mlChildrenUB
            If myChildUsed(X) = 0 Then
                If lIdx = -1 Then lIdx = X
            Else : lCnt += 1
            End If
        Next X

        If lCnt >= ml_MAX_SIMULTANEOUS_FIREWORKS Then Return

        If lIdx = -1 Then
            mlChildrenUB += 1
            ReDim Preserve moChildren(mlChildrenUB)
            ReDim Preserve myChildUsed(mlChildrenUB)
            lIdx = mlChildrenUB
        End If

        Select Case yType
            Case 0
                moChildren(lIdx) = New NormalBurst()
            Case 1
                moChildren(lIdx) = New DiskBurst()
            Case 2
                moChildren(lIdx) = New SparkleBurst()
        End Select

        moChildren(lIdx).Launch(vecLaunch, New Vector3((6 * Rnd()) - 3.0F, 30, (6 * Rnd()) - 3.0F), lExplHeight)

        myChildUsed(lIdx) = 255
    End Sub

End Class
