Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class LoginScreen
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

	Private Shared moPlanet As Mesh
	Private Shared moPlanetTex As Texture
	Private Shared moPlanetGlowTex As Texture
	Private Shared moCloud As Mesh
	Private Shared moCloudTex As Texture

    Private moExplosion() As LoginScreenExplosion

    Private Const ml_STARFIELD_COUNT As Int32 = 15000       'seems high
    Private Shared mbSkyboxReady As Boolean = False
    Private Shared moSkyboxVerts() As SkyboxVert 'CustomVertex.PositionColored
    Private Shared mvbSkybox As VertexBuffer
    Private Shared moSkyboxMat As Material
    Private Shared mlFarPlane As Int32 = 2000000

    Private moCosmoTex As Texture
    Private moCosmoSphere As Mesh

    Private mlLastTime As Int32 = Int32.MaxValue

    '-------- NEW STUFF -------------
    Private moFtrLow As Mesh 
    Private moFtrTex As Texture 

    Private Structure SkyboxVert
        Public Position As Vector3
        Public Size As Single
        Public Color As Int32
    End Structure
    Private SkyboxVertFmt As VertexFormats = VertexFormats.Position Or VertexFormats.PointSize Or VertexFormats.Diffuse

    Private Structure LoginUnit
        Public fTurnVel As Single
        Public fTurnAngle As Single
        Public fYLoc As Single
        Public fYVel As Single
        Public lLastWpnShot As Int32

        Public fLastX As Single
        Public fLastZ As Single
    End Structure
    Private muUnits(29) As LoginUnit

    Private mvecCntr() As Vector3
    Private mvecSpeed() As Vector3
    Private mlShotsFire As Int32 = 0
    Private bRenderCounter As Boolean = False

    Private moBSEntity As BaseEntity
    Private moStationEntity As BaseEntity
    Private moBSQuickDie As BaseEntity
    Private moBSEntity2 As BaseEntity
    Private moPlayerTex As Texture

    Private moCruiser1 As BaseEntity
    Private moCruiser2 As BaseEntity

    Private moFtrChase() As BaseEntity
    Private moFtrStrafe() As BaseEntity

    Public Shared BSDestroyed As Int32 = 0
    'END OF NEW STUFF

    Private mbFirstRender As Boolean = True
    Private mbQuickDieDying As Boolean = False

    Private mlRadioChatter1 As Int32 = -1
    Private mlRC_Val1 As Int32 = -1
    Private mlRadioChatter2 As Int32 = -1
    Private mlRC_Val2 As Int32 = -1

    'Private moMissileMgr As MissileMgr

    Private mlBattleAmbience As Int32

	Public Shared mfFadeoutAlpha As Single = 0.0F
	Public Shared mySequenceEnded As Byte = 0

	Public Shared UpdateButNoRender As Boolean = False
	Private mlPreviousStationShot As Int32 = 0

    Private moModelShader As ModelShader

    Private moSkyboxTexture As Texture

	Private Sub HandleBeamEnd(ByRef oAttacker As BaseEntity, ByRef oTarget As BaseEntity)
		Dim oIntLoc As IntersectInformation

		'shift target to 0,0,0
		Dim vecFrom As Vector3 = New Vector3(oAttacker.LocX - oTarget.LocX, oAttacker.LocY - oTarget.LocY, oAttacker.LocZ - oTarget.LocZ)

		Dim oDir As Vector3 = Vector3.Normalize(vecFrom)
		oDir.Multiply(-1)
		If oTarget.oMesh.oMesh.Intersect(vecFrom, oDir, oIntLoc) = False Then
			'Stop
		End If

		oDir.Scale(oIntLoc.Dist)
		Dim mvecHitLoc As Vector3 = Vector3.Add(New Vector3(oAttacker.LocX, oAttacker.LocY, oAttacker.LocZ), oDir)

		'moMissileMgr.AddMinorExplosion(mvecHitLoc, 3000)
        'goExplMgr.AddMinorExplosion(mvecHitLoc, 2000)
        goExplMgr.Add(mvecHitLoc, Rnd() * 360, Rnd(), CInt(Rnd() * 3), 100, 0, Color.White, Color.Red, 30, 200, True)
    End Sub

    Public Sub DoRender(ByRef moDevice As Device)
        Dim oMat As Material
        Dim matWorld As Matrix = Matrix.Identity
        Dim lCurrentTime As Int32 = timeGetTime

        moDevice.RenderState.Ambient = System.Drawing.Color.DarkGray

        If mfFadeoutAlpha = 255 Then
            If mySequenceEnded <> 255 Then mySequenceEnded = 1
            Return
        End If

        If mbFirstRender = True Then
            moCruiser1.LastUpdateCycle = lCurrentTime
            moCruiser2.LastUpdateCycle = lCurrentTime
            moBSQuickDie.LastUpdateCycle = lCurrentTime
            moBSEntity2.LastUpdateCycle = lCurrentTime - 50
            moStationEntity.LastUpdateCycle = lCurrentTime
            moStationEntity.ObjectID = 1
            moStationEntity.ObjTypeID = 8191
            mbFirstRender = False
            If goSound Is Nothing = False Then mlBattleAmbience = goSound.StartSound("Login.wav", True, SoundMgr.SoundUsage.eAmbience, Vector3.Empty, Vector3.Empty)
            For Y As Int32 = 0 To moFtrStrafe.GetUpperBound(0)
                moFtrStrafe(Y).LastUpdateCycle = lCurrentTime
            Next
        End If

        Dim fElapsed As Single
        If mlLastTime = Int32.MaxValue Then
            fElapsed = 1.0F
        Else : fElapsed = (lCurrentTime - mlLastTime) / 30.0F
        End If
        mlLastTime = lCurrentTime

        With moDevice.Lights(0)
            .Diffuse = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .Ambient = System.Drawing.Color.FromArgb(255, 64, 64, 64)
            .Type = LightType.Point
            .Range = 10000000
            .Specular = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .Attenuation0 = 1
            .Attenuation1 = 0
            .Attenuation2 = 0
            .Position = New Vector3(500000, 100000, 100000)
            .Enabled = True
            .Update()
        End With

        With oMat
            .Ambient = System.Drawing.Color.FromArgb(255, 16, 16, 16)
            .Diffuse = System.Drawing.Color.White
            .Specular = System.Drawing.Color.Gray
            .Emissive = System.Drawing.Color.Black
            .SpecularSharpness = 10
        End With
        moDevice.Material = oMat

        moDevice.Transform.World = Matrix.Identity

        'now, draw our sphere
        If UpdateButNoRender = False Then

            If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                If moModelShader Is Nothing Then moModelShader = New ModelShader()
                Dim vecTemp As Vector3 = New Vector3(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ)
                moModelShader.PrepareToRender(Vector3.Multiply(moDevice.Lights(0).Position, -1.0F), moDevice.Lights(0).Diffuse, moDevice.Lights(0).Ambient, moDevice.Lights(0).Specular)
            End If

            moDevice.RenderState.Wrap0 = WrapCoordinates.Zero
            moDevice.RenderState.CullMode = Cull.None
            moDevice.RenderState.Lighting = False
            moDevice.RenderState.ZBufferWriteEnable = False
            moDevice.SetTexture(0, moCosmoTex)

            matWorld.Multiply(Matrix.RotationY(3.0))
            matWorld.Multiply(Matrix.RotationX(0.3))
            moDevice.Transform.World = matWorld
            matWorld = Matrix.Identity

            moCosmoSphere.DrawSubset(0)
            moDevice.SetTexture(0, Nothing)
            moDevice.RenderState.ZBufferWriteEnable = True
            moDevice.RenderState.Lighting = True
            moDevice.RenderState.CullMode = Cull.CounterClockwise
            moDevice.RenderState.Wrap0 = 0 'WrapCoordinates.Two

            If mbSkyboxReady = False OrElse mvbSkybox Is Nothing OrElse mvbSkybox.Disposed = True Then CreateSkybox(moDevice)

            If moSkyboxTexture Is Nothing OrElse moSkyboxTexture.Disposed = True Then
                moSkyboxTexture = goResMgr.GetTexture("WhiteStar.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "") 
            End If

            With moDevice
                .SetTexture(0, moSkyboxTexture)
                .Material = moSkyboxMat
                'set up our renderstates
                .RenderState.Lighting = False
                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

                .RenderState.SourceBlend = Blend.SourceAlpha
                .RenderState.DestinationBlend = Blend.One
                .RenderState.ZBufferWriteEnable = False

                'Ok, if our device was created with mixed...
                If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                    'Set us up for software vertex processing as point sprites always work in software
                    moDevice.SoftwareVertexProcessing = True
                End If

                .RenderState.PointSpriteEnable = True
                .RenderState.PointSize = 1
                .RenderState.PointScaleEnable = False

                'render our points 
                Dim omatTemp As Matrix = Matrix.Identity
                .Transform.World = omatTemp
                omatTemp = Nothing

                .SetStreamSource(0, mvbSkybox, 0)
                .VertexFormat = SkyboxVertFmt ' CustomVertex.PositionColored.Format
                .DrawPrimitives(PrimitiveType.PointList, 0, ml_STARFIELD_COUNT)

                .RenderState.Lighting = True
                .RenderState.PointSpriteEnable = False
                .RenderState.PointScaleEnable = True
                .RenderState.DestinationBlend = Blend.InvSourceAlpha
                .RenderState.ZBufferWriteEnable = True

                'Ok, if our device was created with mixed...
                If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                    'Set us up for software vertex processing as point sprites always work in software
                    moDevice.SoftwareVertexProcessing = False
                End If
                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With

            moDevice.RenderState.Wrap0 = WrapCoordinates.Zero
        End If

        Dim matRot As Matrix = Matrix.RotationX(Math.PI / 2.0F)
        matWorld = Matrix.Identity

        matWorld.Multiply(matRot)
        matWorld.Multiply(Matrix.RotationZ(180.0F / gdDegreePerRad))

        matWorld.Multiply(Matrix.Translation(-11400, 0, 10300))

        If UpdateButNoRender = False Then
            moDevice.Transform.World = matWorld
            moDevice.Material = oMat
            moDevice.SetTexture(0, moPlanetTex)
            Try
                moPlanet.DrawSubset(0)
            Catch
                If moPlanet Is Nothing = False Then moPlanet.Dispose()
                moPlanet = Nothing
                moPlanet = goResMgr.CreateTexturedSphere(10000.0F, 72, 72, 0)
                If moPlanetTex Is Nothing = False Then moPlanetTex.Dispose()
                moPlanetTex = Nothing
                moPlanetTex = goResMgr.GetTexture("EarthMap.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "Login.pak")
                If moPlanetGlowTex Is Nothing = False Then moPlanetGlowTex.Dispose()
                moPlanetGlowTex = Nothing
                moPlanetGlowTex = goResMgr.GetTexture("PlanetGlow.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "Login.pak")
            End Try
            matWorld = Matrix.Identity
            moDevice.SetTexture(0, moCloudTex)
        End If

        Static fxAngle As Single = 0.0F
        fxAngle += (0.00049F * fElapsed)
        If fxAngle > Math.PI * 2.0F Then fxAngle = 0.0F

        If UpdateButNoRender = False Then
            matWorld.Multiply(matRot)
            matRot = Matrix.RotationY(fxAngle)
            matWorld.Multiply(matRot)

            matWorld.Multiply(Matrix.Translation(-11400, 0, 10300))

            moDevice.Transform.World = matWorld

            With oMat
                .Ambient = Color.Black
                .Diffuse = Color.White
                .Emissive = Color.Black
            End With
            moDevice.Material = oMat
            Try
                moCloud.DrawSubset(0)
            Catch
                If moCloud Is Nothing = False Then moCloud.Dispose()
                moCloud = Nothing
                moCloud = goResMgr.CreateTexturedSphere(10030.0F, 72, 72, 0)
            End Try
        End If
        moDevice.RenderState.Wrap0 = 0 'WrapCoordinates.Two

        With oMat
            .Ambient = Color.DarkGray
            .Diffuse = Color.White
            .Emissive = Color.Black
            .Specular = Color.White
            .SpecularSharpness = 10
        End With
        moDevice.Material = oMat

        If UpdateButNoRender = False AndAlso GFXEngine.gbDeviceLost = False AndAlso GFXEngine.gbPaused = False Then
            Using oSpr As New Sprite(moDevice)
                Const fScale As Single = 44.3892059F '(1000 / 0.88F) / 256.0F
                Dim matTemp As Matrix = Matrix.Scaling(fScale, fScale, fScale)

                matWorld = Matrix.Identity
                matWorld.Multiply(Matrix.Translation(-256.819F, 0, 232))
                matWorld.Multiply(matTemp)

                moDevice.RenderState.ZBufferWriteEnable = False

                moDevice.Transform.World = matWorld
                oSpr.SetWorldViewLH(moDevice.Transform.World, moDevice.Transform.View)
                oSpr.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)
                oSpr.Draw(moPlanetGlowTex, System.Drawing.Rectangle.Empty, New Vector3(256, 256, 0), New Vector3(0, 0, 0), Color.FromArgb(255, 128, 192, 255))
                oSpr.End()

                moDevice.RenderState.ZBufferWriteEnable = True
            End Using
        End If

        matWorld = Matrix.Identity
        With oMat
            .Ambient = Color.White
            .Diffuse = Color.White
            .Emissive = Color.Black
            .Specular = Color.Black
            .SpecularSharpness = 0
        End With
        moDevice.Material = oMat

        Dim bResetShader As Boolean = False
        If moStationEntity Is Nothing = False Then
            Static xfStationAngle As Single = 0.0F
            xfStationAngle += (0.0003F * fElapsed)
            moStationEntity.LocAngle = CShort(xfStationAngle * 10)
            matWorld.Multiply(Matrix.RotationY(xfStationAngle))

            matWorld.Multiply(Matrix.RotationZ(Math.PI / 16))

            matWorld.Multiply(Matrix.Translation(3000, -1500, 2300))

            moStationEntity.SetWorldMatrix(matWorld)
            moStationEntity.SetWorldMatrixCurrent()

            If UpdateButNoRender = False Then
                moDevice.Transform.World = matWorld

                If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                    moModelShader.RenderMesh(moStationEntity, GFXEngine.RenderModelType.eOldData, 3)
                Else
                    If GFXEngine.mbSupportsNewModelMethod = True Then
                        moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
                        moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                        moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.SelectArg1)
                        moDevice.SetTextureStageState(0, TextureStageStates.TextureCoordinateIndex, 0)

                        moDevice.SetTextureStageState(1, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
                        moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.Current)
                        moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.BlendTextureAlpha)
                        moDevice.SetTextureStageState(1, TextureStageStates.TextureCoordinateIndex, 0)
                        moDevice.SetSamplerState(1, SamplerStageStates.MinFilter, TextureFilter.Linear)
                        moDevice.SetSamplerState(1, SamplerStageStates.MagFilter, TextureFilter.Linear)
                        moDevice.SetSamplerState(1, SamplerStageStates.MipFilter, TextureFilter.Linear)

                        moDevice.SetTextureStageState(2, TextureStageStates.ColorOperation, TextureOperation.Modulate)
                        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument1, TextureArgument.Current)
                        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
                        moDevice.SetTextureStageState(2, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
                        moDevice.SetTextureStageState(2, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)

                        moDevice.SetTexture(0, moPlayerTex)
                        moDevice.SetTexture(1, moStationEntity.oMesh.Textures(0))
                    Else
                        moDevice.SetTexture(0, moStationEntity.oMesh.Textures(1))
                    End If

                    moDevice.Material = moStationEntity.oMesh.Materials(0)
                    moStationEntity.oMesh.oMesh.DrawSubset(0)

                    If GFXEngine.mbSupportsNewModelMethod = True Then
                        moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                        moDevice.SetTextureStageState(1, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                        moDevice.SetTextureStageState(2, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)

                        moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.Modulate)
                        moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Disable)
                        moDevice.SetTextureStageState(2, TextureStageStates.ColorOperation, TextureOperation.Disable)

                        moDevice.SetTextureStageState(0, TextureStageStates.TextureCoordinateIndex, 0)
                        moDevice.SetTextureStageState(1, TextureStageStates.TextureCoordinateIndex, 1)

                        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)

                        moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                        moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
                        moDevice.SetTextureStageState(2, TextureStageStates.AlphaOperation, TextureOperation.Disable)

                        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.Current)
                        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument2, TextureArgument.Current)

                        moDevice.SetSamplerState(1, SamplerStageStates.MinFilter, TextureFilter.None)
                        moDevice.SetSamplerState(1, SamplerStageStates.MagFilter, TextureFilter.None)
                        moDevice.SetSamplerState(1, SamplerStageStates.MipFilter, TextureFilter.None)
                    End If
                End If

            End If

            If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                If moStationEntity Is Nothing = False AndAlso moModelShader Is Nothing = False Then moModelShader.EndRender()
            End If

            moDevice.Transform.World = Matrix.Identity
            For X As Int32 = 0 To moExplosion.Length - 1 ' moexplosion(0)
                moExplosion(X).Render(UpdateButNoRender)
            Next X

            If muUnits Is Nothing = False Then
                If lCurrentTime - mlPreviousStationShot > 30 Then
                    mlPreviousStationShot = lCurrentTime
                    Dim lTmpIdxVal As Int32 = CInt(Rnd() * muUnits.GetUpperBound(0))
                    With muUnits(lTmpIdxVal)
                        Dim fTmpX As Single = .fLastX - 3000
                        fTmpX *= 2
                        fTmpX += 3000
                        Dim fTmpZ As Single = .fLastZ - 2300
                        fTmpZ *= 2
                        fTmpZ += 2300
                        goWpnMgr.GenerateSoundFX = False
                        goWpnMgr.AddNewEffect(moStationEntity, CInt(fTmpX), CInt(.fYLoc - 1500), CInt(fTmpZ), WeaponType.eShortRedPulse, False, 0)
                        goWpnMgr.GenerateSoundFX = True
                    End With
                End If
            End If

            If lCurrentTime - moStationEntity.LastUpdateCycle > 150000 AndAlso moStationEntity.ObjectID = 1 Then
                goEntityDeath.AddNewDeathSequence(moStationEntity, 1)
                If goSound Is Nothing = False Then goSound.StartSound(moStationEntity.oMesh.sDeathSeqWav(0), False, SoundMgr.SoundUsage.eWeaponsFireMissiles, New Vector3(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ), Nothing)
                moStationEntity.ObjectID = 2
            End If

            If BSDestroyed = 7 Then
                goCamera.ScreenShake(1000)

                bResetShader = True

                moStationEntity.ClearParticleFX()
                moStationEntity = Nothing
                For X As Int32 = 0 To muUnits.GetUpperBound(0)
                    muUnits(X).lLastWpnShot = Int32.MaxValue
                Next
                If goSound Is Nothing = False Then goSound.StopSound(mlBattleAmbience)
            End If
        Else
            mfFadeoutAlpha += fElapsed
            If mfFadeoutAlpha > 255 Then mfFadeoutAlpha = 255
        End If

        If UpdateButNoRender = False Then
            moDevice.SetTexture(0, moFtrTex)
            moDevice.RenderState.AlphaBlendEnable = False
        End If
        For X As Int32 = 0 To muUnits.GetUpperBound(0)
            muUnits(X).fTurnAngle += (muUnits(X).fTurnVel * fElapsed)
            If muUnits(X).fTurnAngle > 360.0F Then muUnits(X).fTurnAngle -= 360.0F
            If muUnits(X).fTurnAngle < 0.0F Then muUnits(X).fTurnAngle += 360.0F

            Dim fX As Single = 2200.0F
            Dim fZ As Single = 0.0F

            'This gets our X and Z location
            RotatePoint(0, 0, fX, fZ, muUnits(X).fTurnAngle)

            'Now, for our Y
            muUnits(X).fYLoc += muUnits(X).fYVel * fElapsed
            If muUnits(X).fYLoc > 1100 Then
                muUnits(X).fYLoc = 1100
                muUnits(X).fYVel = -muUnits(X).fYVel
            ElseIf muUnits(X).fYLoc < -2100 Then
                muUnits(X).fYLoc = -2100
                muUnits(X).fYVel = -muUnits(X).fYVel
            End If

            If muUnits(X).lLastWpnShot < lCurrentTime Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect_PointToEntity(CInt((3000 + fX)), CInt((muUnits(X).fYLoc - 1500)), CInt((2300 + fZ)), moStationEntity, WeaponType.eShortGreenPulse, True, False)
                goWpnMgr.GenerateSoundFX = True
                muUnits(X).lLastWpnShot = lCurrentTime + (500 + CInt(Rnd() * 100) - 50)
            End If
            muUnits(X).fLastX = 3000 + fX
            muUnits(X).fLastZ = 2300 + fZ

            If UpdateButNoRender = False Then
                matWorld = GetFighterMatrix(X)
                matWorld.Multiply(Matrix.Translation(3000 + fX, muUnits(X).fYLoc - 1500.0F, 2300 + fZ))
                moDevice.Transform.World = matWorld
                moFtrLow.DrawSubset(0)
            End If
        Next X

        If (bResetShader = True OrElse moStationEntity Is Nothing = False) AndAlso muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 AndAlso moModelShader Is Nothing = False Then
            Dim vecTemp As Vector3 = New Vector3(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ)
            moModelShader.PrepareToRender(Vector3.Multiply(moDevice.Lights(0).Position, -1.0F), moDevice.Lights(0).Diffuse, moDevice.Lights(0).Ambient, moDevice.Lights(0).Specular)
        End If

        'Battleship stuff...
        If moBSEntity Is Nothing = False Then
            moBSEntity.LocZ += (1.3F * fElapsed)
            moBSEntity.VelZ = 1.3F
            If moBSEntity.LocZ > -8600 AndAlso mlShotsFire = 0 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity.LocX, moBSEntity.LocY, moBSEntity.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity, moStationEntity)
                goWpnMgr.GenerateSoundFX = True

                mlShotsFire += 1
                moBSEntity.lChildUB = 0
            ElseIf moBSEntity.LocZ > -8500 AndAlso moBSEntity.lChildUB = 0 Then
                Dim fAngle As Single = LineAngleDegrees(CInt(moStationEntity.LocX), CInt(moStationEntity.LocZ), CInt(moBSEntity.LocX), CInt(moBSEntity.LocZ))

                Dim fMyAngle As Single = (moStationEntity.LocAngle / 10.0F) - 90.0F
                fAngle -= fMyAngle * fElapsed
                If fAngle > 360 Then fAngle -= 360
                If fAngle < 0 Then fAngle += 360

                Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
                Dim vecFrom As Vector3 = moStationEntity.GetFireFromLoc(ySide)
                vecFrom = Vector3.TransformCoordinate(vecFrom, moStationEntity.GetWorldMatrix)
                'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity, 50)
                goMissileMgr.AddMissile(vecFrom, moBSEntity, 50, 50, 0, Int32.MaxValue, True)
                moBSEntity.lChildUB = 1
            ElseIf moBSEntity.LocZ > -8400 AndAlso moBSEntity.lChildUB = 1 Then
                Dim fAngle As Single = LineAngleDegrees(CInt(moStationEntity.LocX), CInt(moStationEntity.LocZ), CInt(moBSEntity.LocX), CInt(moBSEntity.LocZ))

                Dim fMyAngle As Single = (moStationEntity.LocAngle / 10.0F) - 90.0F
                fAngle -= fMyAngle * fElapsed
                If fAngle > 360 Then fAngle -= 360
                If fAngle < 0 Then fAngle += 360

                Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
                Dim vecFrom As Vector3 = moStationEntity.GetFireFromLoc(ySide)
                vecFrom = Vector3.TransformCoordinate(vecFrom, moStationEntity.GetWorldMatrix)
                'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity, 50)
                goMissileMgr.AddMissile(vecFrom, moBSEntity, 50, 50, 0, Int32.MaxValue, True)
                moBSEntity.lChildUB += 1

            ElseIf moBSEntity.LocZ > -8300 AndAlso mlShotsFire = 1 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity.LocX, moBSEntity.LocY, moBSEntity.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                mlShotsFire += 1
            ElseIf moBSEntity.LocZ > -7500 AndAlso mlShotsFire = 2 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity.LocX, moBSEntity.LocY, moBSEntity.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                mlShotsFire += 1
            ElseIf moBSEntity.LocZ > -7200 AndAlso mlShotsFire = 3 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity.LocX, moBSEntity.LocY, moBSEntity.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                mlShotsFire += 1
            ElseIf moBSEntity.LocZ > -6500 AndAlso mlShotsFire = 4 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moBSEntity, moCruiser1, WeaponType.eFlickerPurpleBeam, True, False, 0)
                goWpnMgr.AddNewEffect(moBSEntity, moCruiser2, WeaponType.eFlickerPurpleBeam, True, False, 0)
                HandleBeamEnd(moCruiser1, moBSEntity)
                If goSound Is Nothing = False Then goSound.StartSound("Beam1a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity.LocX, moBSEntity.LocY, moBSEntity.LocZ), New Vector3(0, 0, 0))
                goWpnMgr.GenerateSoundFX = True
                mlShotsFire += 1
            ElseIf moBSEntity.LocZ > -6450 AndAlso mlShotsFire = 5 Then
                goEntityDeath.AddNewDeathSequence(moBSEntity, 1)
                mlShotsFire += 1
            End If

            If BSDestroyed = 3 Then
                moBSEntity.ClearParticleFX()
                moBSEntity = Nothing
                goCamera.ScreenShake(1000)
                moCruiser1.iCombatTactics += 1S
                moCruiser2.iCombatTactics += 1S
            Else
                If UpdateButNoRender = False Then
                    matWorld = Matrix.Identity
                    matWorld.Multiply(Matrix.RotationY(DegreeToRadian(185.0F)))
                    matWorld.Multiply(Matrix.Translation(moBSEntity.LocX, moBSEntity.LocY, moBSEntity.LocZ))
                    'moBSEntity.fMapWrapLocX = moBSEntity.LocX
                    moDevice.Transform.World = matWorld
                    moDevice.SetTexture(0, moBSEntity.oMesh.Textures(0))
                    moDevice.Material = moBSEntity.oMesh.Materials(0)

                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moBSEntity, GFXEngine.RenderModelType.eNoSpecial, 1)
                    Else
                        moBSEntity.oMesh.oMesh.DrawSubset(0)
                        moBSEntity.oMesh.oMesh.DrawSubset(1)
                    End If
                End If

                If lCurrentTime - moBSEntity.LastUpdateCycle > 1200 Then '650 Then
                    goWpnMgr.GenerateSoundFX = False
                    goWpnMgr.AddNewEffect(moStationEntity, moBSEntity, WeaponType.eMetallicProjectile_Lead, True, False, 0)
                    If goSound Is Nothing = False Then goSound.StartSound("LargeProj3.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity.LocX, moBSEntity.LocY, moBSEntity.LocZ), New Vector3(0, 0, 0))
                    goWpnMgr.GenerateSoundFX = True
                    moBSEntity.LastUpdateCycle = lCurrentTime
                End If
            End If
        End If

        If moBSEntity2 Is Nothing = False Then

            If moBSEntity2.iCombatTactics = 0 Then
                moBSEntity2.LocZ += (moBSEntity2.VelX * fElapsed)
            ElseIf moBSEntity2.iCombatTactics = 1 Then
                If moBSEntity2.LocAngle > 2150 Then
                    moBSEntity2.LocAngle -= CShort(1S * fElapsed)
                    moBSEntity2.LocYaw -= CShort(1S * fElapsed)
                    If moBSEntity2.LocYaw < -450 Then moBSEntity2.LocYaw = -450
                    If moBSEntity2.LocAngle < 2150 Then moBSEntity2.LocAngle = 2150
                ElseIf moBSEntity2.LocYaw <> 0 Then
                    moBSEntity2.LocYaw += CShort(1S * fElapsed)
                    If moBSEntity2.LocYaw > 0 Then moBSEntity2.LocYaw = 0
                End If

                Dim fSpeed As Single = 1.3F * fElapsed
                Dim fTemp As Single = (moBSEntity2.LocAngle - 1800.0F) / 900.0F
                moBSEntity2.VelX = -(fSpeed * (1.0F - fTemp))
                moBSEntity2.VelZ = fSpeed - Math.Abs(moBSEntity2.VelX)

                moBSEntity2.LocX += moBSEntity2.VelX
                moBSEntity2.LocZ += moBSEntity2.VelZ
            ElseIf moBSEntity2.iCombatTactics = 5 Then
                If moBSEntity2.LocYaw > 0 Then
                    moBSEntity2.LocYaw -= 1S
                ElseIf moBSEntity2.LocYaw < 0 Then
                    moBSEntity2.LocYaw += 1S
                End If
                moBSEntity2.LocZ += (1.3F * fElapsed)
            Else
                If moBSEntity2.LocAngle < 2750 Then
                    moBSEntity2.LocAngle += CShort(1S * fElapsed)
                    moBSEntity2.LocYaw += CShort(1S * fElapsed)
                    If moBSEntity2.LocYaw > 450 Then moBSEntity2.LocYaw = 450
                    If moBSEntity2.LocAngle > 2750 Then moBSEntity2.LocAngle = 2750
                ElseIf moBSEntity2.LocYaw <> 0 Then
                    If moCruiser1 Is Nothing = False AndAlso moCruiser1.iCombatTactics = 2 Then moCruiser1.iCombatTactics = 3
                    moBSEntity2.LocYaw -= CShort(1S * fElapsed)
                    If moBSEntity2.LocYaw < 0 Then moBSEntity2.LocYaw = 0
                End If

                Dim fSpeed As Single = 1.3F
                Dim fTemp As Single = 0

                If moBSEntity2.LocAngle > 2699 Then
                    moBSEntity2.VelX = 0
                    moBSEntity2.VelZ = fSpeed
                Else
                    fTemp = (2700 - moBSEntity2.LocAngle) / 900.0F
                    moBSEntity2.VelX = (fSpeed * (1.0F - fTemp)) - 0.65F ' -(fSpeed * (1.0F - fTemp))
                    moBSEntity2.VelZ = fSpeed - Math.Abs(moBSEntity2.VelX)
                End If

                moBSEntity2.LocX += (moBSEntity2.VelX * fElapsed)
                moBSEntity2.LocZ += (moBSEntity2.VelZ * fElapsed)
            End If

            If moBSEntity2.LocZ > -8600 AndAlso moBSEntity2.iTargetingTactics = 0 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity2, moStationEntity)
                goWpnMgr.GenerateSoundFX = True

                moBSEntity2.iTargetingTactics += 1S
                moBSEntity2.lChildUB = 0
            ElseIf moBSEntity2.LocZ > -8300 AndAlso moBSEntity2.iTargetingTactics = 1 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity2, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -7500 AndAlso moBSEntity2.iTargetingTactics = 2 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity2, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -6850 AndAlso moBSEntity2.iTargetingTactics = 3 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moCruiser1, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                'goWpnMgr.AddNewEffect(moBSEntity2, -11400, 0, 10300, WeaponType.eSolidRedBeam, True)
                'goWpnMgr.AddNewEffect(moBSEntity2, -11400, 0, 10300, WeaponType.eSolidRedBeam, True)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                'HandleBeamEnd(moBSEntity2, moCruiser1)
                goWpnMgr.GenerateSoundFX = True
                moBSEntity2.iTargetingTactics += 1S

                Dim fAngle As Single = LineAngleDegrees(CInt(moStationEntity.LocX), CInt(moStationEntity.LocZ), CInt(moBSEntity2.LocX), CInt(moBSEntity2.LocZ))

                Dim fMyAngle As Single = (moStationEntity.LocAngle / 10.0F) - 90.0F
                fAngle -= (fMyAngle * fElapsed)
                If fAngle > 360 Then fAngle -= 360
                If fAngle < 0 Then fAngle += 360

                Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
                Dim vecFrom As Vector3 = moStationEntity.GetFireFromLoc(ySide)
                vecFrom = Vector3.TransformCoordinate(vecFrom, moStationEntity.GetWorldMatrix)
                'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity2, 50)
                goMissileMgr.AddMissile(vecFrom, moBSEntity2, 50, 50, 0, Int32.MaxValue, True)
                vecFrom = moStationEntity.GetFireFromLoc(ySide)
                vecFrom = Vector3.TransformCoordinate(vecFrom, moStationEntity.GetWorldMatrix)
                vecFrom = moStationEntity.GetFireFromLoc(ySide)
                vecFrom = Vector3.TransformCoordinate(vecFrom, moStationEntity.GetWorldMatrix)
                'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity2, 49)
                goMissileMgr.AddMissile(vecFrom, moBSEntity2, 49, 50, 0, Int32.MaxValue, True)

                'moBSEntity2.iTargetingTactics += 1S
                moBSEntity2.iCombatTactics = 2

            ElseIf moBSEntity2.LocZ > -6750 AndAlso moBSEntity2.iTargetingTactics = 4 Then
                goEntityDeath.AddNewDeathSequence(moCruiser1, 1)
                moBSEntity2.iTargetingTactics += 1S

            ElseIf moBSEntity2.LocZ > -6500 AndAlso moBSEntity2.iTargetingTactics = 5 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity2, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -6450 AndAlso moBSEntity2.iTargetingTactics = 6 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moBSEntity2, moStationEntity, WeaponType.eFlickerTealBeam, True, False, 0)
                goWpnMgr.AddNewEffect(moBSEntity2, moStationEntity, WeaponType.eFlickerTealBeam, True, False, 0)
                goWpnMgr.AddNewEffect(moBSEntity2, moStationEntity, WeaponType.eFlickerTealBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam1a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moStationEntity, moBSEntity2)
                moBSEntity2.iTargetingTactics += 1S
                moBSEntity2.yStructureHP = 0
                goWpnMgr.GenerateSoundFX = True
                If moCruiser2 Is Nothing = False Then moCruiser2.LastUpdateCycle = lCurrentTime + 50000
            ElseIf moBSEntity2.LocZ > -6350 AndAlso moBSEntity2.iTargetingTactics = 7 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moCruiser2, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                'HandleBeamEnd(moBSEntity2, moCruiser2)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                goWpnMgr.GenerateSoundFX = True
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -6250 AndAlso moBSEntity2.iTargetingTactics = 8 Then
                goEntityDeath.AddNewDeathSequence(moCruiser2, 1)
                'moMissileMgr.AddTargetedMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100)
                goMissileMgr.AddMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100, 50, 0, Int32.MaxValue, True)
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -6100 AndAlso moBSEntity2.iTargetingTactics = 9 Then
                'moMissileMgr.AddTargetedMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100)
                goMissileMgr.AddMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100, 50, 0, Int32.MaxValue, True)
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -6050 AndAlso moBSEntity2.iTargetingTactics = 10 Then
                moBSEntity2.iCombatTactics = 5
                'goEntityDeath.AddNewDeathSequence(moBSEntity2, 1)
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -5700 AndAlso moBSEntity2.iTargetingTactics = 11 Then
                moBSEntity2.iTargetingTactics += 1S
                moFtrStrafe(0).LocX = -3000
                moFtrStrafe(1).LocX = -2800
                moFtrStrafe(2).LocX = -3000

                For ltmpX As Int32 = 0 To 2
                    moFtrStrafe(ltmpX).LocY = moBSEntity2.LocY + CInt(Rnd() * 20) + 350
                    moFtrStrafe(ltmpX).LocAngle = 3150
                    moFtrStrafe(ltmpX).LocYaw = -450S + CShort(Rnd() * 20)
                    moFtrStrafe(ltmpX).yStructureHP = CByte(Rnd() * 100)
                    moFtrStrafe(ltmpX).InitializeEngineFX()
                    moFtrStrafe(ltmpX).TestForBurnFX()
                Next ltmpX
            ElseIf moBSEntity2.LocZ > -5500 AndAlso moBSEntity2.iTargetingTactics = 12 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moBSEntity2, moStationEntity, WeaponType.eFlickerTealBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam1a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                'HandleBeamEnd(moStationEntity, moBSEntity2)
                moBSEntity2.iTargetingTactics += 1S
                moBSEntity2.yStructureHP = 0
                goWpnMgr.GenerateSoundFX = True
            ElseIf moBSEntity2.LocZ > -5480 AndAlso moBSEntity2.iTargetingTactics = 13 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity2, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -5250 AndAlso moBSEntity2.iTargetingTactics = 14 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moStationEntity, moBSEntity2, WeaponType.eSolidRedBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moBSEntity2, moStationEntity)
                goWpnMgr.GenerateSoundFX = True
                moBSEntity2.iTargetingTactics += 1S

                'moMissileMgr.AddTargetedMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100)
                goMissileMgr.AddMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100, 50, 0, Int32.MaxValue, True)
            ElseIf moBSEntity2.LocZ > -5200 AndAlso moBSEntity2.iTargetingTactics = 15 Then
                'moMissileMgr.AddTargetedMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100)
                goMissileMgr.AddMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSEntity2, 100, 50, 0, Int32.MaxValue, True)
                moBSEntity2.iTargetingTactics += 1S
            ElseIf moBSEntity2.LocZ > -5100 AndAlso moBSEntity2.iTargetingTactics = 16 Then
                goWpnMgr.GenerateSoundFX = False
                goWpnMgr.AddNewEffect(moBSEntity2, moStationEntity, WeaponType.eFlickerTealBeam, True, False, 0)
                If goSound Is Nothing = False Then goSound.StartSound("Beam1a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                HandleBeamEnd(moStationEntity, moBSEntity2)
                moBSEntity2.iTargetingTactics += 1S
                moBSEntity2.yStructureHP = 0
                goWpnMgr.GenerateSoundFX = True
            ElseIf moBSEntity2.LocZ > -5090 AndAlso moBSEntity2.iTargetingTactics = 17 Then
                goEntityDeath.AddNewDeathSequence(moBSEntity2, 1)
                moBSEntity2.iTargetingTactics += 1S
            End If

            If BSDestroyed = 6 Then
                moBSEntity2.ClearParticleFX()
                moBSEntity2 = Nothing
                goCamera.ScreenShake(100)
            Else

                If UpdateButNoRender = False Then
                    matWorld = moBSEntity2.GetWorldMatrix
                    'moBSEntity2.fMapWrapLocX = moBSEntity2.LocX
                    moDevice.Transform.World = matWorld
                    moDevice.SetTexture(0, moBSEntity2.oMesh.Textures(0))
                    moDevice.Material = moBSEntity2.oMesh.Materials(0)
                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moBSEntity2, GFXEngine.RenderModelType.eNoSpecial, 1)
                    Else
                        moBSEntity2.oMesh.oMesh.DrawSubset(0)
                        moBSEntity2.oMesh.oMesh.DrawSubset(1)
                    End If
                End If

                If lCurrentTime - moBSEntity2.LastUpdateCycle > 1200 Then '650 Then
                    goWpnMgr.GenerateSoundFX = False
                    goWpnMgr.AddNewEffect(moStationEntity, moBSEntity2, WeaponType.eMetallicProjectile_Lead, True, False, 0)
                    If goSound Is Nothing = False Then goSound.StartSound("LargeProj3.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSEntity2.LocX, moBSEntity2.LocY, moBSEntity2.LocZ), New Vector3(0, 0, 0))
                    goWpnMgr.GenerateSoundFX = True
                    If moBSEntity Is Nothing = False Then
                        moBSEntity2.LastUpdateCycle = moBSEntity.LastUpdateCycle + 600
                    Else : moBSEntity2.LastUpdateCycle = lCurrentTime
                    End If
                End If
            End If
        End If

        If moBSQuickDie Is Nothing = False Then
            If BSDestroyed > 0 Then
                goCamera.ScreenShake(100)
                moBSQuickDie.ClearParticleFX()
                moBSQuickDie = Nothing
            Else
                If moBSQuickDie.LocAngle <> 710 Then
                    moBSQuickDie.LocAngle -= 1S
                    If moBSQuickDie.LocAngle < 0 Then moBSQuickDie.LocAngle += 3600S
                    moBSQuickDie.LocYaw -= 1S
                    If moBSQuickDie.LocYaw < -450 Then moBSQuickDie.LocYaw = -450
                ElseIf moBSQuickDie.LocYaw <> 0 Then
                    moBSQuickDie.LocYaw += 1S
                End If

                If moBSQuickDie.ObjectID <> 22 Then
                    'moMissileMgr.AddTargetedMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSQuickDie, 100)
                    goMissileMgr.AddMissile(New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), moBSQuickDie, 100, 50, 0, Int32.MaxValue, True)
                    moBSQuickDie.ObjectID = 22
                End If

                If moBSQuickDie.LocAngle = 910 Then
                    goWpnMgr.GenerateSoundFX = False
                    goWpnMgr.AddNewEffect(moStationEntity, moBSQuickDie, WeaponType.eSolidRedBeam, True, False, 0)
                    If goSound Is Nothing = False Then goSound.StartSound("Beam2a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moBSQuickDie.LocX, moBSQuickDie.LocY, moBSQuickDie.LocZ), New Vector3(0, 0, 0))
                    HandleBeamEnd(moBSQuickDie, moStationEntity)
                    goWpnMgr.GenerateSoundFX = True
                End If

                moBSQuickDie.LocX += (0.5F * fElapsed)
                moBSQuickDie.LocZ -= (0.5F * fElapsed)

                'moBSQuickDie.fMapWrapLocX = moBSQuickDie.LocX

                If UpdateButNoRender = False Then
                    moDevice.Transform.World = moBSQuickDie.GetWorldMatrix()
                    moDevice.SetTexture(0, moBSQuickDie.oMesh.Textures(0))
                    moDevice.Material = moBSQuickDie.oMesh.Materials(0)
                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moBSQuickDie, GFXEngine.RenderModelType.eNoSpecial, 1)
                    Else
                        moBSQuickDie.oMesh.oMesh.DrawSubset(0)
                        moBSQuickDie.oMesh.oMesh.DrawSubset(1)
                    End If
                End If

                If lCurrentTime - moBSQuickDie.LastUpdateCycle > 15000 Then
                    moBSQuickDie.LastUpdateCycle = lCurrentTime + 150000
                    goEntityDeath.AddNewDeathSequence(moBSQuickDie, 1)
                    mbQuickDieDying = True
                End If
            End If
        End If

        If moCruiser1 Is Nothing = False Then

            If BSDestroyed >= 4 Then
                If moCruiser1.yStructureHP <> 0 Then moCruiser1.ClearParticleFX()
                moCruiser1.yStructureHP = 0
                moCruiser1.yShieldHP = 0
                moCruiser1.yArmorHP(0) = 0 : moCruiser1.yArmorHP(1) = 0 : moCruiser1.yArmorHP(2) = 0 : moCruiser1.yArmorHP(3) = 0
                'moCruiser1.fMapWrapLocX = moCruiser1.LocX
                moCruiser1.TestForBurnFX()

                moCruiser1.LocYaw += 2S
                If moCruiser1.LocYaw > 3600S Then moCruiser1.LocYaw -= 3600S
                moCruiser1.LocAngle -= 2S
                If moCruiser1.LocAngle < 0 Then moCruiser1.LocAngle += 3600S

                moCruiser1.LocZ -= (0.2F * fElapsed)


                moDevice.RenderState.CullMode = Cull.None
                matWorld = moCruiser1.GetWorldMatrix()
                'moCruiser1.fMapWrapLocX = moCruiser1.LocX

                If UpdateButNoRender = False Then
                    moDevice.Transform.World = matWorld
                    moDevice.SetTexture(0, moCruiser1.oMesh.Textures(0))
                    moDevice.Material = moCruiser1.oMesh.Materials(0)
                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moCruiser1, GFXEngine.RenderModelType.eNoSpecial, 3)
                    Else
                        moCruiser1.oMesh.oMesh.DrawSubset(0)
                    End If
                End If

                moDevice.RenderState.CullMode = Cull.CounterClockwise
            Else
                moDevice.RenderState.CullMode = Cull.None
                matWorld = moCruiser1.GetWorldMatrix()
                'moCruiser1.fMapWrapLocX = moCruiser1.LocX

                If UpdateButNoRender = False Then
                    moDevice.Transform.World = matWorld
                    moDevice.SetTexture(0, moCruiser1.oMesh.Textures(0))
                    moDevice.Material = moCruiser1.oMesh.Materials(0)
                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moCruiser1, GFXEngine.RenderModelType.eNoSpecial, 3)
                    Else
                        moCruiser1.oMesh.oMesh.DrawSubset(0)
                    End If
                End If

                moDevice.RenderState.CullMode = Cull.CounterClockwise

                If BSDestroyed = 0 Then
                    moCruiser1.LocZ -= (4.0F * fElapsed)
                    If lCurrentTime - moCruiser1.LastUpdateCycle > 10000 AndAlso mbQuickDieDying = False Then
                        goWpnMgr.GenerateSoundFX = False
                        goWpnMgr.AddNewEffect(moBSQuickDie, moCruiser1, WeaponType.eFlickerPurpleBeam, True, False, 0)
                        If goSound Is Nothing = False Then goSound.StartSound("Beam1a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moCruiser1.LocX, moCruiser1.LocY, moCruiser1.LocZ), New Vector3(0, 0, 0))
                        HandleBeamEnd(moCruiser1, moBSQuickDie)
                        goWpnMgr.GenerateSoundFX = True
                        moCruiser1.LastUpdateCycle += 2500
                    End If
                ElseIf moCruiser1.iCombatTactics < 2 Then
                    If moCruiser1.LocAngle <> 1350 Then
                        moCruiser1.LocAngle += CShort(3S * fElapsed)
                        If moCruiser1.LocAngle > 1350 Then moCruiser1.LocAngle = 1350
                        moCruiser1.LocYaw += 3S
                        If moCruiser1.LocYaw > 450 Then moCruiser1.LocYaw = 450
                    ElseIf moCruiser1.LocYaw <> 0 Then
                        moCruiser1.LastUpdateCycle = lCurrentTime - 12000
                        moCruiser1.LocYaw -= 3S
                        If moCruiser1.LocYaw < 0 Then moCruiser1.LocYaw = 0
                    End If
                    moCruiser1.LocX -= (2.0F * fElapsed)
                    moCruiser1.LocZ -= (2.0F * fElapsed)

                    If moCruiser1.LocYaw = 0 Then
                        If lCurrentTime - moCruiser1.LastUpdateCycle > 15000 Then
                            Dim fAngle As Single = LineAngleDegrees(CInt(moCruiser1.LocX), CInt(moCruiser1.LocZ), CInt(moBSEntity.LocX), CInt(moBSEntity.LocZ))

                            Dim fMyAngle As Single = (moCruiser1.LocAngle / 10.0F) - 90.0F
                            fAngle -= (fMyAngle * fElapsed)
                            If fAngle > 360 Then fAngle -= 360
                            If fAngle < 0 Then fAngle += 360

                            Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
                            Dim vecFrom As Vector3 = moCruiser1.GetFireFromLoc(ySide)
                            vecFrom = Vector3.TransformCoordinate(vecFrom, moCruiser1.GetWorldMatrix)
                            'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity, 100)
                            goMissileMgr.AddMissile(vecFrom, moBSEntity, 100, 50, 1, Int32.MaxValue, True)

                            moCruiser1.LastUpdateCycle = lCurrentTime
                            If moCruiser1.iTargetingTactics = 0 Then
                                moCruiser1.iTargetingTactics = 1
                                moCruiser1.LastUpdateCycle -= 14000
                            Else
                                moCruiser1.iTargetingTactics = 0
                                moCruiser1.iCombatTactics += 1S
                            End If
                        End If
                    End If
                ElseIf moCruiser1.iCombatTactics < 3 Then
                    If moBSEntity2.iCombatTactics = 0 Then moBSEntity2.iCombatTactics = 1

                    If moFtrChase(1) Is Nothing = False Then
                        If moFtrChase(1).MaxColonists <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(moFtrChase(1).MaxColonists)
                        moFtrChase(1) = Nothing
                    End If

                    If moCruiser1.LocAngle <> 700 Then
                        moCruiser1.LocAngle -= CShort(3S * fElapsed)
                        If moCruiser1.LocAngle < 700 Then moCruiser1.LocAngle = 700
                        moCruiser1.LocYaw -= 3S
                        If moCruiser1.LocYaw < -450 Then moCruiser1.LocYaw = -450
                    ElseIf moCruiser1.LocYaw <> 0 Then
                        moCruiser1.LocYaw += 3S
                        If moCruiser1.LocYaw > 0 Then moCruiser1.LocYaw = 0
                    End If
                    moCruiser1.LocX += (2.0F * fElapsed)
                    moCruiser1.LocZ -= (2.0F * fElapsed)
                Else
                    If moCruiser1.LocAngle <> 1350 Then
                        moCruiser1.LocAngle += 3S
                        If moCruiser1.LocAngle > 1350 Then moCruiser1.LocAngle = 1350
                        moCruiser1.LocYaw += 3S
                        If moCruiser1.LocYaw > 450 Then moCruiser1.LocYaw = 450
                    ElseIf moCruiser1.LocYaw <> 0 Then
                        moCruiser1.LocYaw -= 3S
                        If moCruiser1.LocYaw < 0 Then moCruiser1.LocYaw = 0
                        moCruiser1.LastUpdateCycle = lCurrentTime - 15000
                        moCruiser2.LastUpdateCycle = lCurrentTime - 13000
                    Else
                        If lCurrentTime - moCruiser1.LastUpdateCycle > 15000 Then
                            Dim fAngle As Single = LineAngleDegrees(CInt(moCruiser1.LocX), CInt(moCruiser1.LocZ), CInt(moBSEntity2.LocX), CInt(moBSEntity2.LocZ))

                            Dim fMyAngle As Single = (moCruiser1.LocAngle / 10.0F) - 90.0F
                            fAngle -= (fMyAngle * fElapsed)
                            If fAngle > 360 Then fAngle -= 360
                            If fAngle < 0 Then fAngle += 360

                            Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
                            Dim vecFrom As Vector3 = moCruiser1.GetFireFromLoc(ySide)
                            vecFrom = Vector3.TransformCoordinate(vecFrom, moCruiser1.GetWorldMatrix)
                            'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity2, 100)
                            goMissileMgr.AddMissile(vecFrom, moBSEntity2, 100, 50, 1, Int32.MaxValue, True)

                            moCruiser1.LastUpdateCycle = lCurrentTime
                            If moCruiser1.iTargetingTactics = 0 Then
                                moCruiser1.iTargetingTactics = 1
                                moCruiser1.LastUpdateCycle -= 14000
                            Else
                                moCruiser1.iTargetingTactics = 0
                            End If
                        End If
                    End If

                    Dim fSpeed As Single = (4.0F * fElapsed)
                    Dim fTemp As Single = (1350 - moCruiser1.LocAngle) / 650.0F
                    moCruiser1.VelX = (fSpeed * fTemp) - 2.0F
                    moCruiser1.VelZ = -(fSpeed - Math.Abs(moCruiser1.VelX))

                    moCruiser1.LocX += (moCruiser1.VelX * fElapsed)
                    moCruiser1.LocZ -= (moCruiser1.VelZ * fElapsed)
                End If
            End If

        End If

        If moCruiser2 Is Nothing = False Then
            moDevice.RenderState.CullMode = Cull.None
            matWorld = moCruiser2.GetWorldMatrix()
            'moCruiser2.fMapWrapLocX = moCruiser2.LocX
            moDevice.Transform.World = matWorld

            If UpdateButNoRender = False Then
                moDevice.SetTexture(0, moCruiser2.oMesh.Textures(0))
                moDevice.Material = moCruiser2.oMesh.Materials(0)
                If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                    moModelShader.RenderMesh(moCruiser2, GFXEngine.RenderModelType.eNoSpecial, 3)
                Else
                    moCruiser2.oMesh.oMesh.DrawSubset(0)
                End If
            End If

            moDevice.RenderState.CullMode = Cull.CounterClockwise

            If BSDestroyed = 0 Then
                moCruiser2.LocZ -= (4.0F * fElapsed)
                If lCurrentTime - moCruiser2.LastUpdateCycle > 10000 AndAlso mbQuickDieDying = False Then
                    goWpnMgr.GenerateSoundFX = False
                    goWpnMgr.AddNewEffect(moBSQuickDie, moCruiser2, WeaponType.eFlickerPurpleBeam, True, False, 0)
                    If goSound Is Nothing = False Then goSound.StartSound("Beam1a.wav", False, SoundMgr.SoundUsage.eWeaponsFireBombs, New Vector3(moCruiser2.LocX, moCruiser2.LocY, moCruiser2.LocZ), New Vector3(0, 0, 0))
                    goWpnMgr.GenerateSoundFX = True
                    moCruiser2.LastUpdateCycle += 2500
                End If
            ElseIf moCruiser2.iCombatTactics < 2 Then
                If moCruiser2.LocAngle <> 1350 Then
                    moCruiser2.LocAngle += CShort(3S * fElapsed)
                    If moCruiser2.LocAngle > 1350 Then moCruiser2.LocAngle = 1350
                    moCruiser2.LocYaw += 3S
                    If moCruiser2.LocYaw > 450 Then moCruiser2.LocYaw = 450
                ElseIf moCruiser2.LocYaw <> 0 Then
                    moCruiser2.LastUpdateCycle = lCurrentTime - 16000
                    moCruiser2.LocYaw -= 3S
                    If moCruiser2.LocYaw < 0 Then moCruiser2.LocYaw = 0
                End If
                moCruiser2.LocX -= (2.0F * fElapsed)
                moCruiser2.LocZ -= (2.0F * fElapsed)

                If moCruiser2.LocYaw = 0 Then
                    If lCurrentTime - moCruiser2.LastUpdateCycle > 15000 Then
                        If moBSEntity Is Nothing = False Then

                            Dim fAngle As Single = LineAngleDegrees(CInt(moCruiser2.LocX), CInt(moCruiser2.LocZ), CInt(moBSEntity.LocX), CInt(moBSEntity.LocZ))

                            Dim fMyAngle As Single = (moCruiser2.LocAngle / 10.0F) - 90.0F
                            fAngle -= (fMyAngle * fElapsed)
                            If fAngle > 360 Then fAngle -= 360
                            If fAngle < 0 Then fAngle += 360

                            Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
                            Dim vecFrom As Vector3 = moCruiser2.GetFireFromLoc(ySide)
                            vecFrom = Vector3.TransformCoordinate(vecFrom, moCruiser2.GetWorldMatrix)
                            'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity, 100)
                            goMissileMgr.AddMissile(vecFrom, moBSEntity, 100, 50, 1, Int32.MaxValue, True)

                            moCruiser2.LastUpdateCycle = lCurrentTime
                            If moCruiser2.iTargetingTactics = 0 Then
                                moCruiser2.iTargetingTactics = 1
                                moCruiser2.LastUpdateCycle -= 14000
                            Else
                                moCruiser2.iTargetingTactics = 0
                                moCruiser2.iCombatTactics += 1S
                            End If
                        End If
                    End If
                End If
            ElseIf moCruiser2.iCombatTactics < 3 Then
                moCruiser2.LocX -= (2.0F * fElapsed)
                moCruiser2.LocZ -= (2.0F * fElapsed)
            Else
                If moCruiser2.LocAngle <> 700 Then
                    moCruiser2.LocAngle -= CShort(3S * fElapsed)
                    If moCruiser2.LocAngle < 700 Then moCruiser2.LocAngle = 700
                    moCruiser2.LocYaw -= 3S
                    If moCruiser2.LocYaw < -450 Then moCruiser2.LocYaw = -450
                ElseIf moCruiser2.LocYaw <> 0 Then
                    moCruiser2.LocYaw += 3S
                    If moCruiser2.LocYaw > 0 Then moCruiser2.LocYaw = 0
                End If

                'at 1350, I am moving -2 x, -2 z
                'at 700, I am moving 2x, -2z


                Dim fSpeed As Single = (4.0F * fElapsed)
                Dim fTemp As Single = (1350 - moCruiser2.LocAngle) / 650.0F
                moCruiser2.VelX = (fSpeed * fTemp) - 2.0F
                moCruiser2.VelZ = -(fSpeed - Math.Abs(moCruiser2.VelX))

                moCruiser2.LocX += moCruiser2.VelX
                moCruiser2.LocZ += moCruiser2.VelZ

                If BSDestroyed = 5 Then
                    moCruiser2.ClearParticleFX()
                    moCruiser2 = Nothing
                Else
                    If lCurrentTime - moCruiser2.LastUpdateCycle > 15000 Then
                        Dim fAngle As Single = LineAngleDegrees(CInt(moCruiser2.LocX), CInt(moCruiser2.LocZ), CInt(moBSEntity2.LocX), CInt(moBSEntity2.LocZ))

                        Dim fMyAngle As Single = (moCruiser2.LocAngle / 10.0F) - 90.0F
                        fAngle -= fMyAngle
                        If fAngle > 360 Then fAngle -= 360
                        If fAngle < 0 Then fAngle += 360

                        Dim ySide As Byte = AngleToQuadrant(CInt(fAngle))
                        Dim vecFrom As Vector3 = moCruiser2.GetFireFromLoc(ySide)
                        vecFrom = Vector3.TransformCoordinate(vecFrom, moCruiser2.GetWorldMatrix)
                        'moMissileMgr.AddTargetedMissile(vecFrom, moBSEntity2, 100)
                        goMissileMgr.AddMissile(vecFrom, moBSEntity2, 100, 50, 1, Int32.MaxValue, True)

                        moCruiser2.LastUpdateCycle = lCurrentTime
                        If moCruiser2.iTargetingTactics = 0 Then
                            moCruiser2.iTargetingTactics = 1
                            moCruiser2.LastUpdateCycle -= 14000
                        Else
                            moCruiser2.iTargetingTactics = 0
                            'moCruiser2.iCombatTactics += 1S
                        End If
                    End If
                End If
            End If
        End If

        If moFtrChase(1) Is Nothing = False Then
            If BSDestroyed = 1 Then
                moDevice.RenderState.CullMode = Cull.None
                For X As Int32 = 0 To moFtrChase.GetUpperBound(0)

                    'we'll control it with the angle of the ship...
                    moFtrChase(X).ObjectID += moFtrChase(X).ObjTypeID
                    If moFtrChase(X).ObjectID < -440 Then
                        'If moFtrChase(X).ObjTypeID < 0 Then moFtrChase(X).ObjTypeID *= -1S
                        moFtrChase(X).ObjTypeID = 6S
                    ElseIf moFtrChase(X).ObjectID > 10 Then
                        moFtrChase(X).ObjTypeID = 0
                    End If
                    Dim fTmpY As Single = (12.0F * fElapsed)
                    fTmpY *= (moFtrChase(X).ObjectID / 450.0F)
                    moFtrChase(X).LocY += CInt(fTmpY)

                    moFtrChase(X).LocX += (Rnd() - 0.5F) '(15.0F * fElapsed)
                    moFtrChase(X).LocZ += (30.0F * fElapsed)
                    'moFtrChase(X).fMapWrapLocX = moFtrChase(X).LocX

                    If moFtrChase(X).LocYaw > 0 Then
                        moFtrChase(X).LocYaw -= 1S
                    ElseIf moFtrChase(X).LocYaw < 0 Then
                        moFtrChase(X).LocYaw += 1S
                    End If

                    'matWorld = GetFighterEntityMatrix(moFtrChase(X))
                    If UpdateButNoRender = False Then
                        matWorld = moFtrChase(X).GetWorldMatrix()
                        moDevice.Transform.World = matWorld
                        moDevice.SetTexture(0, moFtrChase(1).oMesh.Textures(0))
                        moDevice.Material = moFtrChase(1).oMesh.Materials(0)

                        If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                            Dim lTexMod As Int32 = 3
                            If X = 1 Then
                                lTexMod = 1
                            End If
                            moModelShader.RenderMesh(moFtrChase(X), GFXEngine.RenderModelType.eNoSpecial, lTexMod)
                        Else
                            moFtrChase(X).oMesh.oMesh.DrawSubset(0)
                        End If
                    End If

                Next X
                moDevice.RenderState.CullMode = Cull.CounterClockwise

                'moFtrChase(1).MaxColonists = goSound.HandleMovingSound(moFtrChase(1).MaxColonists, "Unit Sounds\SmallShipRoar1.wav", EpicaSound.SoundUsage.eUnitSounds, New Vector3(moFtrChase(1).LocX, moFtrChase(1).LocY, moFtrChase(1).LocZ), New Vector3(moFtrChase(1).VelX, 0, moFtrChase(1).VelZ))

                If moFtrChase(0).LocZ > -9800 AndAlso moFtrChase(0).iCombatTactics <> 1 Then
                    moFtrChase(0).iCombatTactics = 1
                    If goSound Is Nothing = False Then goSound.StartSound("Unit Sounds\LoginFlyby.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                    goCamera.ScreenShake(50)
                End If

                If moFtrChase(1).LocZ > -9800 AndAlso moFtrChase(1).iCombatTactics = 0 Then
                    moFtrChase(1).iCombatTactics = 1
                    moFtrChase(1).LastUpdateCycle = lCurrentTime - 130
                    goCamera.ScreenShake(50)
                End If

                If moFtrChase(1).iCombatTactics = 1 Then
                    If lCurrentTime - moFtrChase(1).LastUpdateCycle > 130 Then
                        goWpnMgr.GenerateSoundFX = False
                        goWpnMgr.AddNewEffect(moFtrChase(0), moFtrChase(1), WeaponType.eMetallicProjectile_Gold, True, False, 0)
                        If goSound Is Nothing = False Then moFtrChase(1).lChildUB = goSound.HandleMovingSound(moFtrChase(1).lChildUB, "Projectile Weapons\FastRepeatSmall.wav", SoundMgr.SoundUsage.eWeaponsFireProjectile, New Vector3(moFtrChase(1).LocX, moFtrChase(1).LocY, moFtrChase(1).LocZ), Nothing)
                        goWpnMgr.GenerateSoundFX = True
                        moFtrChase(1).lProducingID += 1
                        If moFtrChase(1).lProducingID = 50 Then
                            moFtrChase(1).iCombatTactics = 2
                            goEntityDeath.AddNewDeathSequence(moFtrChase(0), 1)
                            If moFtrChase(1).lChildUB <> -1 AndAlso goSound Is Nothing = False Then goSound.StopSound(moFtrChase(1).lChildUB)
                        End If
                    End If
                End If
            ElseIf BSDestroyed > 0 Then 'If BSDestroyed = 2 Then
                If moFtrChase(0) Is Nothing = False Then
                    moFtrChase(0).ClearParticleFX()
                    moFtrChase(0) = Nothing
                End If

                If moFtrChase(1).LocAngle <> 0 Then
                    moFtrChase(1).LocAngle += 20S
                    If moFtrChase(1).LocAngle > 3600 Then moFtrChase(1).LocAngle = 0
                    moFtrChase(1).LocYaw += 5S
                    If moFtrChase(1).LocYaw > 450 Then moFtrChase(1).LocYaw = 450
                ElseIf moFtrChase(1).LocYaw <> 0 Then
                    moFtrChase(1).LocYaw -= 5S
                End If

                moFtrChase(1).ObjectID += moFtrChase(1).ObjTypeID
                If moFtrChase(1).ObjectID < -440 Then
                    'If moFtrChase(X).ObjTypeID < 0 Then moFtrChase(X).ObjTypeID *= -1S
                    moFtrChase(1).ObjTypeID = 6S
                ElseIf moFtrChase(1).ObjectID > 10 Then
                    moFtrChase(1).ObjTypeID = 0
                End If
                Dim fTmpY As Single = (12.0F * fElapsed)
                fTmpY *= (moFtrChase(1).ObjectID / 450.0F)
                moFtrChase(1).LocY += CInt(fTmpY)

                'at 2250, we are going -15, 15
                'at 0, we are going 30, 0
                Dim fTmpAngle As Single = moFtrChase(1).LocAngle
                If fTmpAngle = 0 Then fTmpAngle = 3600

                Dim fSpeed As Single = 30.0F
                Dim fTemp As Single = (fTmpAngle - 1800) / 1800.0F

                moFtrChase(1).VelX = -30 + (60 * fTemp)
                moFtrChase(1).VelZ = 30 - Math.Abs(moFtrChase(1).VelX)

                moFtrChase(1).LocX += moFtrChase(1).VelX
                moFtrChase(1).LocZ += moFtrChase(1).VelZ

                'moFtrChase(1).fMapWrapLocX = moFtrChase(1).LocX

                If UpdateButNoRender = False Then
                    matWorld = moFtrChase(1).GetWorldMatrix()
                    moDevice.Transform.World = matWorld
                    moDevice.SetTexture(0, moFtrChase(1).oMesh.Textures(0))
                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moFtrChase(1), GFXEngine.RenderModelType.eNoSpecial, 1)
                    Else
                        moFtrChase(1).oMesh.oMesh.DrawSubset(0)
                    End If
                End If

                If moFtrChase(1).LocX > 10000 Then moFtrChase(1) = Nothing
            End If
        End If

        If moBSEntity2 Is Nothing = False AndAlso moBSEntity2.iTargetingTactics > 11 Then
            For X As Int32 = 0 To moFtrStrafe.GetUpperBound(0)

                moFtrStrafe(X).LocYaw += 10S

                If moFtrStrafe(X).LocYaw > 0 Then

                    If moFtrStrafe(X).LocAngle <> 3550 Then
                        moFtrStrafe(X).LocAngle += 10S
                        If moFtrStrafe(X).LocAngle > 3550 Then moFtrStrafe(X).LocAngle = 3550
                    End If
                    Dim fTemp As Single = (3600.0F - moFtrStrafe(X).LocAngle) / 900.0F
                    moFtrStrafe(X).VelZ = fTemp * 30.0F
                    moFtrStrafe(X).VelX = 30 - moFtrStrafe(X).VelZ

                    moFtrStrafe(X).LocX += moFtrStrafe(X).VelX
                    moFtrStrafe(X).LocZ += moFtrStrafe(X).VelZ

                    If moFtrStrafe(X).lProducingID <> 500 Then
                        'moMissileMgr.AddTargetedMissile(New Vector3(moFtrStrafe(X).LocX, moFtrStrafe(X).LocY, moFtrStrafe(X).LocZ), moBSEntity2, 100)
                        goMissileMgr.AddMissile(New Vector3(moFtrStrafe(X).LocX, moFtrStrafe(X).LocY, moFtrStrafe(X).LocZ), moBSEntity2, 100, 50, 2, Int32.MaxValue, True)
                        moFtrStrafe(X).lProducingID = 500
                    End If
                Else
                    moFtrStrafe(X).LocX += (15.0F * fElapsed)
                    moFtrStrafe(X).LocZ += (15.0F * fElapsed)
                End If
                'moFtrStrafe(X).fMapWrapLocX = moFtrStrafe(X).LocX
                If UpdateButNoRender = False Then
                    matWorld = moFtrStrafe(X).GetWorldMatrix()
                    moDevice.Transform.World = matWorld
                    moDevice.SetTexture(0, moFtrStrafe(X).oMesh.Textures(0))
                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moFtrStrafe(X), GFXEngine.RenderModelType.eNoSpecial, 3)
                    Else
                        moFtrStrafe(X).oMesh.oMesh.DrawSubset(0)
                    End If
                End If
            Next X
        ElseIf lCurrentTime - moFtrStrafe(0).LastUpdateCycle > 60000 Then

            If moFtrStrafe(0).ObjectID <> 22 Then
                If goSound Is Nothing = False Then goSound.StartSound("Unit Sounds\LoginFlyby2.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                moFtrStrafe(0).ObjectID = 22
            End If

            For X As Int32 = 0 To moFtrStrafe.GetUpperBound(0)

                moFtrStrafe(X).LocX -= (20.0F * fElapsed)
                moFtrStrafe(X).LocZ += (10.0F * fElapsed)

                If moFtrStrafe(X).ObjTypeID <> 22 Then
                    If moFtrStrafe(X).LocYaw < 0 Then
                        moFtrStrafe(X).LocYaw += 10S
                    ElseIf moFtrStrafe(X).LocYaw > 0 Then
                        moFtrStrafe(X).LocYaw -= 10S
                    End If
                Else
                    moFtrStrafe(X).LocYaw -= 5S
                End If

                'moFtrStrafe(X).fMapWrapLocX = moFtrStrafe(X).LocX

                If UpdateButNoRender = False Then
                    matWorld = moFtrStrafe(X).GetWorldMatrix()
                    moDevice.Transform.World = matWorld
                    moDevice.SetTexture(0, moFtrStrafe(X).oMesh.Textures(0))
                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                        moModelShader.RenderMesh(moFtrStrafe(X), GFXEngine.RenderModelType.eNoSpecial, 3)
                    Else
                        moFtrStrafe(X).oMesh.oMesh.DrawSubset(0)
                    End If
                End If

                If moFtrStrafe(X).LocX < 300 AndAlso moFtrStrafe(X).iCombatTactics = 0 Then
                    moFtrStrafe(X).iCombatTactics = 1
                    'moMissileMgr.AddTargetedMissile(New Vector3(moFtrStrafe(X).LocX, moFtrStrafe(X).LocY, moFtrStrafe(X).LocZ), moBSEntity, 100)
                    goMissileMgr.AddMissile(New Vector3(moFtrStrafe(X).LocX, moFtrStrafe(X).LocY, moFtrStrafe(X).LocZ), moBSEntity, 100, 50, 2, Int32.MaxValue, True)
                ElseIf moFtrStrafe(X).LocX < 0 AndAlso moFtrStrafe(X).iCombatTactics = 1 Then
                    moFtrStrafe(X).iCombatTactics = 2
                    'moMissileMgr.AddTargetedMissile(New Vector3(moFtrStrafe(X).LocX, moFtrStrafe(X).LocY, moFtrStrafe(X).LocZ), moBSEntity, 100)
                    goMissileMgr.AddMissile(New Vector3(moFtrStrafe(X).LocX, moFtrStrafe(X).LocY, moFtrStrafe(X).LocZ), moBSEntity, 100, 50, 2, Int32.MaxValue, True)
                Else
                    moFtrStrafe(X).ObjTypeID = 22
                End If
            Next X

            If lCurrentTime - moFtrStrafe(0).LastUpdateCycle > 67300 Then
                moFtrStrafe(0).LastUpdateCycle = lCurrentTime + 500000
                For X As Int32 = 0 To moFtrStrafe.GetUpperBound(0)
                    moFtrStrafe(X).ClearParticleFX()
                Next X
            End If
        End If

        If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
            If moModelShader Is Nothing = False Then moModelShader.EndRender()
        End If

        If UpdateButNoRender = False Then moDevice.RenderState.AlphaBlendEnable = True
        moDevice.Transform.World = Matrix.Identity

        goBurnMarkMgr.BeginFrame()
        'goBurnMarkMgr.RenderFromObj(moStationEntity)
        goBurnMarkMgr.RenderFromObj(moBSEntity)
        goBurnMarkMgr.RenderFromObj(moBSEntity2)
        goBurnMarkMgr.RenderFromObj(moCruiser1)
        goBurnMarkMgr.RenderFromObj(moCruiser2)
        goBurnMarkMgr.EndFrame()


        'Placed all of these in try...catch blocks because of an intermittent race crash occurring at login
        Try
            goWpnMgr.RenderFX(UpdateButNoRender)
            goEntityDeath.RenderSequences(UpdateButNoRender)
            goPFXEngine32.Render(UpdateButNoRender)
            goShldMgr.RenderFX(UpdateButNoRender)
            goMissileMgr.Render(UpdateButNoRender)
            goExplMgr.Render(UpdateButNoRender)
        Catch
            'Do nothing...
        End Try

        moDevice.Transform.World = Matrix.Identity

        If moStationEntity Is Nothing = False Then
            If Rnd() * 100 < 10 Then
                Dim lIdx As Int32 = CInt(Rnd() * 8) + 1
                If lIdx <> mlRC_Val2 AndAlso goSound Is Nothing = False Then mlRadioChatter1 = goSound.StartSound("RC" & lIdx & ".wav", False, SoundMgr.SoundUsage.eRadioChatter, New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), New Vector3(0, 0, 0))
            End If

            If Rnd() * 100 < 10 Then
                Dim lIdx As Int32 = CInt(Rnd() * 8) + 1
                If lIdx <> mlRC_Val1 AndAlso goSound Is Nothing = False Then mlRadioChatter2 = goSound.StartSound("RC" & lIdx & ".wav", False, SoundMgr.SoundUsage.eRadioChatter, New Vector3(moStationEntity.LocX, moStationEntity.LocY, moStationEntity.LocZ), New Vector3(0, 0, 0))
            End If
        Else
            moDevice.RenderState.ZBufferEnable = False
            Dim rcSrc As Rectangle

            Dim rcDest As Rectangle = System.Drawing.Rectangle.FromLTRB(0, 0, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight)

            rcSrc.Location = New Point(192, 0)
            rcSrc.Width = 62
            rcSrc.Height = 64

            'Now, draw it...
            If GFXEngine.gbDeviceLost = False AndAlso GFXEngine.gbPaused = False Then
                With goUILib.Pen
                    goUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
                    '.Begin(SpriteFlags.AlphaBlend)
                    .Draw2D(goUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(0, 0), Color.FromArgb(CInt(mfFadeoutAlpha), 0, 0, 0))
                    '.End()
                    goUILib.EndPenSprite()
                End With
            End If

            moDevice.RenderState.ZBufferEnable = True
        End If
    End Sub

    Private Sub ExplosionComplete(ByRef oExplosion As LoginScreenExplosion)
        Dim lIdx As Int32 = CInt(Rnd() * muUnits.GetUpperBound(0))
        oExplosion.ResetExplosion(New Vector3(muUnits(lIdx).fLastX, (muUnits(lIdx).fYLoc - 1500), muUnits(lIdx).fLastZ))
    End Sub

    Private Sub CreateSkybox(ByRef moDevice As Device)
        Dim lIdx As Int32
        Dim fX As Single
        Dim fY As Single
        Dim fZ As Single
        Dim fD As Single

        Dim lColor As System.Drawing.Color
        'Dim lCVal As Int32

        'now, generate our starfield...
        ReDim moSkyboxVerts(ml_STARFIELD_COUNT - 1)
        For lIdx = 0 To ml_STARFIELD_COUNT - 1
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

            'Dim lCValR As Int32 = CInt(64 * (0.9 + (Rnd() * 0.2F)))
            'Dim lCValG As Int32 = CInt(128 * (0.9 + (Rnd() * 0.2F)))
            'If Rnd() < 0.2F Then lCValR = lCValG
            'Dim lCValB As Int32 = CInt(230 * (0.9 + (Rnd() * 0.2F)))
            'lColor = System.Drawing.Color.FromArgb(255, lCValR, lCValG, lCValB)

            Dim lRandRoll As Int32 = CInt(Rnd() * 100)
            If lRandRoll < 50 Then
                Dim lCValR As Int32 = 192 'CInt(64 * (0.9 + (Rnd() * 0.2F)))
                Dim lCValG As Int32 = 220 'CInt(128 * (0.9 + (Rnd() * 0.2F)))
                Dim lCValB As Int32 = 255 'CInt(230 * (0.9 + (Rnd() * 0.2F)))
                lColor = System.Drawing.Color.FromArgb(255, lCValR, lCValG, lCValB)
            ElseIf lRandRoll < 75 Then
                Dim lCValR As Int32 = 255 'CInt(230 * (0.9 + (Rnd() * 0.2F)))
                Dim lCValG As Int32 = 255 'CInt(230 * (0.9 + (Rnd() * 0.2F)))
                Dim lCValB As Int32 = 192 'CInt(64 * (0.9 + (Rnd() * 0.2F)))
                lColor = System.Drawing.Color.FromArgb(255, lCValR, lCValG, lCValB)
            Else
                Dim lCValR As Int32 = 255 'CInt(230 * (0.9 + (Rnd() * 0.2F)))
                Dim lCValG As Int32 = 192 'CInt(64 * (0.9 + (Rnd() * 0.2F)))
                Dim lCValB As Int32 = 192 'CInt(64 * (0.9 + (Rnd() * 0.2F)))
                lColor = System.Drawing.Color.FromArgb(255, lCValR, lCValG, lCValB)
            End If


            moSkyboxVerts(lIdx).Position = New Vector3(fX, fY, fZ)
            Dim fSize As Single = (CSng(Rnd() ^ 3) * 15.0F)
            If fSize > 10.0F AndAlso Rnd() > 0.25F Then fSize /= 2
            moSkyboxVerts(lIdx).Size = fSize + 3.5F '(Rnd() * 10.0F) + 1.0F 
            moSkyboxVerts(lIdx).Color = lColor.ToArgb

            'lCVal = CInt(Int(Rnd() * 192) + 63)
            'lColor = System.Drawing.Color.FromArgb(255, lCVal, lCVal, lCVal)

            ''moSkyboxVerts(lIdx) = New CustomVertex.PositionColored(fX, fY, fZ, lColor.ToArgb)
            'moSkyboxVerts(lIdx).Position = New Vector3(fX, fY, fZ)
            ''moSkyboxVerts(lIdx).Size = CShort(20 * Rnd())
            'moSkyboxVerts(lIdx).Size = (Rnd() * 10.0F) + 3.5F '(Rnd() * 10.0F) + 1.0F
            'moSkyboxVerts(lIdx).Color = lColor.ToArgb
        Next lIdx

        'Now, create our buffer
        mvbSkybox = New VertexBuffer(New SkyboxVert().GetType, ml_STARFIELD_COUNT, moDevice, Usage.Points, SkyboxVertFmt, Pool.Managed)
        mvbSkybox.SetData(moSkyboxVerts, 0, LockFlags.None)

        With moSkyboxMat
            .Ambient = System.Drawing.Color.White
            .Diffuse = .Ambient
            .Emissive = .Ambient
            .Specular = System.Drawing.Color.Black
        End With

        mbSkyboxReady = True
    End Sub

    Private Class LoginScreenExplosion
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

        Public FXIndex As Int32 'for parent usage

        Private moTex As Texture
        Public mvecEmitter As Vector3
        Private moParticles() As Particle
        Private mlPrevFrame As Int32
        Private mbEmitterStopping As Boolean = False

        Public moPoints() As CustomVertex.PositionColoredTextured
        Public mlParticleUB As Int32
        Public EmitterStopped As Boolean = False
        'Public AttachedObject As RenderObject

        Private mlResetCnt As Int32

        Private moFlashParticle As Particle
        Private moFlashPoint As CustomVertex.PositionColoredTextured

        Public Event ExplosionComplete(ByRef oExplosion As LoginScreenExplosion)

        Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

        Public Sub ResetExplosion(ByVal vecLoc As Vector3)

            If goSound Is Nothing = False Then goSound.StartSound("Explosions\FlakExplosion.wav", False, SoundMgr.SoundUsage.eNonDeathExplosions, vecLoc, New Vector3(0, 0, 0))

            EmitterStopped = False
            mbEmitterStopping = False
            mlResetCnt = 0
            mvecEmitter = vecLoc

            For X As Int32 = 0 To mlParticleUB
                'moParticles(X) = New Particle()
                ResetParticle(X)
            Next X

            moFlashParticle.Reset(vecLoc.X, vecLoc.Y, vecLoc.Z, 0, 0, 0, 0, 0, 0, 255, 255, 255, 255)
            moFlashParticle.fAChg = -15


        End Sub

        Public Sub New(ByVal vecLoc As Vector3)
            Dim X As Int32

            mvecEmitter = vecLoc
            mlParticleUB = 400
            ReDim moParticles(mlParticleUB)
            ReDim moPoints(mlParticleUB + 1)

            For X = 0 To mlParticleUB
                moParticles(X) = New Particle()
                ResetParticle(X)
            Next X

            moFlashParticle = New Particle()

            moFlashParticle.Reset(vecLoc.X, vecLoc.Y, vecLoc.Z, 0, 0, 0, 0, 0, 0, 255, 255, 255, 255)
            moFlashParticle.fAChg = -25 '-15
        End Sub

        Public Sub Render(ByVal bUpdateNoRender As Boolean)

            If moTex Is Nothing OrElse moTex.Disposed = True Then
                moTex = goResMgr.GetTexture("Flare2.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "xpl.pak")
            End If

            Update()

            If EmitterStopped = True Then Exit Sub

            Dim moDevice As Device = GFXEngine.moDevice

            If bUpdateNoRender = False Then
                'And now render them... first set up our device for renders...
                With moDevice
                    '.Transform.World = Matrix.Identity

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
                        If .DeviceCaps.MaxPointSize > 8 Then
                            .RenderState.PointSize = 8  '64
                        Else : .RenderState.PointSize = .DeviceCaps.MaxPointSize
                        End If
                        .RenderState.PointScaleA = 0
                        .RenderState.PointScaleB = 0
                        .RenderState.PointScaleC = 0.3F
                    End If

                    .RenderState.SourceBlend = Blend.SourceAlpha
                    .RenderState.DestinationBlend = Blend.One
                    .RenderState.AlphaBlendEnable = True

                    .RenderState.ZBufferWriteEnable = False

                    .RenderState.Lighting = False

                    .VertexFormat = CustomVertex.PositionColoredTextured.Format
                    '.VertexFormat = moPoints(0).Format
                    .SetTexture(0, moTex)

                    .DrawUserPrimitives(PrimitiveType.PointList, mlParticleUB + 1, moPoints)

                    .RenderState.PointSize = 128
                    .DrawUserPrimitives(PrimitiveType.PointList, 1, moFlashPoint)

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

                If EmitterStopped = True Then
                    RaiseEvent ExplosionComplete(Me)
                End If
            Else
                If lIndex = 0 Then mlResetCnt += 1
                If mlResetCnt > 1 Then mbEmitterStopping = True

                'is it part of the initial explosion?
                'If lIndex < 200 AndAlso mlResetCnt < 2 Then
                ''Ok, if we haven't hit this yet
                fX = mvecEmitter.X + (Rnd() * 100 - 50)  '(GetNxtRnd() - 0.5)
                fY = mvecEmitter.Y + (Rnd() * 100 - 50) '(GetNxtRnd() - 0.5)
                fZ = mvecEmitter.Z + (Rnd() * 100 - 50) ' (GetNxtRnd() - 0.5)
                fXS = (Rnd() * 20) - 10
                fYS = (Rnd() * 20) - 10
                fZS = (Rnd() * 20) - 10

                fXA = Rnd() * 2.0F : fYA = Rnd() * 2.0F : fZA = Rnd() * 2.0F
                If fXS > 0 Then fXA = -fXA
                If fYS > 0 Then fYA = -fYA
                If fZS > 0 Then fZA = -fZA

                moParticles(lIndex).Reset(fX, fY, fZ, fXS, fYS, fZS, fXA, fYA, fZA, 200 + (Rnd() * 50), 90 + (Rnd() * 30), 32 + (Rnd() * 32), 255)
                moParticles(lIndex).fAChg = -((Rnd() * 32) + 24) '16)
            End If


        End Sub

        Private Sub Update()
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

                        moPoints(X).Color = .ParticleColor.ToArgb
                        moPoints(X).Position = .vecLoc
                    End If
                End With
            Next X

            moFlashParticle.Update(fElapsed)
            moFlashPoint.Color = moFlashParticle.ParticleColor.ToArgb
            moFlashPoint.Position = moFlashParticle.vecLoc
        End Sub
    End Class

    Protected Overrides Sub Finalize()

        If glCurrentEnvirView <> CurrentView.eStartupDSELogo AndAlso glCurrentEnvirView <> CurrentView.eStartupLogin Then
            If goResMgr Is Nothing = False Then
                goResMgr.DeleteTexture("SpaceBox1.bmp")
                goResMgr.DeleteTexture("EarthMap.bmp")
                goResMgr.DeleteTexture("PlanetGlow.dds")
                goResMgr.DeleteTexture("PlanetCloud1.dds")
            End If
            If moPlanet Is Nothing = False AndAlso moPlanet.Disposed = False Then
                moPlanet.Dispose()
            End If
            moPlanet = Nothing

            If moPlanetTex Is Nothing = False Then moPlanetTex.Dispose()
            moPlanetTex = Nothing
            moPlanetGlowTex = Nothing

            If moCloud Is Nothing = False AndAlso moCloud.Disposed = False Then
                moCloud.Dispose()
            End If
            moCloud = Nothing

            If moCosmoSphere Is Nothing = False AndAlso moCosmoSphere.Disposed = False Then
                moCosmoSphere.Dispose()
            End If
            moCosmoSphere = Nothing
            moCosmoTex = Nothing
        End If

        moStationEntity = Nothing

        If moExplosion Is Nothing = False Then
            For X As Int32 = 0 To moExplosion.Length - 1
                moExplosion(X) = Nothing
            Next X
        End If
        Erase moExplosion

        If moFtrLow Is Nothing = False Then moFtrLow.Dispose()
        moFtrLow = Nothing
        moFtrTex = Nothing

        If moBSEntity Is Nothing = False Then
            moBSEntity.ClearParticleFX()
            moBSEntity = Nothing
        End If
        If moStationEntity Is Nothing = False Then
            moStationEntity.ClearParticleFX()
            moStationEntity = Nothing
        End If
        If moBSQuickDie Is Nothing = False Then
            moBSQuickDie.ClearParticleFX()
            moBSQuickDie = Nothing
        End If
        If moBSEntity2 Is Nothing = False Then
            moBSEntity2.ClearParticleFX()
            moBSEntity2 = Nothing
        End If
        If moCruiser1 Is Nothing = False Then
            moCruiser1.ClearParticleFX()
            moCruiser1 = Nothing
        End If
        If moCruiser2 Is Nothing = False Then
            moCruiser2.ClearParticleFX()
            moCruiser2 = Nothing
        End If
        If moFtrChase Is Nothing = False Then
            For X As Int32 = 0 To moFtrChase.GetUpperBound(0)
                If moFtrChase(X) Is Nothing = False Then
                    moFtrChase(X).ClearParticleFX()
                    moFtrChase(X) = Nothing
                End If
            Next X
        End If
        If moFtrStrafe Is Nothing = False Then
            For X As Int32 = 0 To moFtrStrafe.GetUpperBound(0)
                If moFtrStrafe(X) Is Nothing = False Then
                    moFtrStrafe(X).ClearParticleFX()
                    moFtrStrafe(X) = Nothing
                End If
            Next
        End If

        moPlayerTex = Nothing

        If glCurrentEnvirView <> CurrentView.eStartupDSELogo AndAlso glCurrentEnvirView <> CurrentView.eStartupLogin Then
            If goEntityDeath Is Nothing = False Then goEntityDeath.ClearAll()
            If goPFXEngine32 Is Nothing = False Then goPFXEngine32.ClearAllEmitters()
            If goWpnMgr Is Nothing = False Then goWpnMgr.CleanAll()
            If goSound Is Nothing = False Then goSound.KillAllSounds()
            If goMissileMgr Is Nothing = False Then goMissileMgr.KillAll()

            If mvbSkybox Is Nothing = False Then mvbSkybox.Dispose()
            mvbSkybox = Nothing
            Erase moSkyboxVerts

            SoundMgr.lMinDistance = 1000
            SoundMgr.lMaxDistance = 15000
        End If

        Erase muUnits

        'Force a garbage collection
        System.GC.Collect()

        MyBase.Finalize()
    End Sub

    Public Sub New(ByRef moDevice As Device)

        'Dim oINI As InitFile = New InitFile()
        'If CInt(Val(oINI.GetString("SETTINGS", "ShowIntroScreen", "1"))) = 0 Then
        '    mfFadeoutAlpha = 255.0F
        '    Return
        'End If
        If muSettings.ShowIntro = False Then
            mfFadeoutAlpha = 255.0F
            mySequenceEnded = 255
            Return
        End If

        mbFirstRender = True

        If moPlanet Is Nothing OrElse moPlanet.Disposed = True Then moPlanet = goResMgr.CreateTexturedSphere(10000.0F, 72, 72, 0)
        If moPlanetTex Is Nothing OrElse moPlanetTex.Disposed = True Then moPlanetTex = goResMgr.GetTexture("EarthMap.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "Login.pak")
        If moPlanetGlowTex Is Nothing OrElse moPlanetGlowTex.Disposed = True Then moPlanetGlowTex = goResMgr.GetTexture("PlanetGlow.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "Login.pak")
        If moCloud Is Nothing OrElse moCloud.Disposed = True Then moCloud = goResMgr.CreateTexturedSphere(10030.0F, 72, 72, 0)
        If moCloudTex Is Nothing OrElse moCloudTex.Disposed = True Then moCloudTex = goResMgr.GetTexture("PlanetCloud1.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "Login.pak")
        If moCosmoTex Is Nothing OrElse moCosmoTex.Disposed = True Then moCosmoTex = goResMgr.GetTexture("SpaceBox1.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "Login.pak")
        If moCosmoSphere Is Nothing OrElse moCosmoSphere.Disposed = True Then moCosmoSphere = goResMgr.CreateTexturedSphere(250000, 32, 32, 0)

        If GFXEngine.mbSupportsNewModelMethod = True Then moPlayerTex = goResMgr.GetTexture("Player.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "textures2.pak")

        Randomize()

        SoundMgr.lMinDistance = 3000
        SoundMgr.lMaxDistance = 15000

        Dim lBigShipModelID As Int32 = 5
        Dim lChasedModelID As Int32 = 2
        Dim lChaserModelID As Int32 = 3
        Dim lStrafeModelID As Int32 = 2
        Dim lCorvetteModelID As Int32 = 4

        Dim lTmpModelRoll As Int32 = CInt(Rnd() * 100)
        If lTmpModelRoll < 25 Then
            lChaserModelID = 66
        ElseIf lTmpModelRoll < 50 Then
            lChaserModelID = 67
        ElseIf lTmpModelRoll < 75 Then
            lChaserModelID = 3
        Else
            lChaserModelID = 73
        End If

        lTmpModelRoll = CInt(Rnd() * 100)
        If lTmpModelRoll < 25 Then
            lChasedModelID = 46
        ElseIf lTmpModelRoll < 50 Then
            lChasedModelID = 68
        ElseIf lTmpModelRoll < 75 Then
            lChasedModelID = 96
        Else
            lChasedModelID = 2
        End If

        lTmpModelRoll = CInt(Rnd() * 100)
        If lTmpModelRoll < 25 Then
            lStrafeModelID = 45
        ElseIf lTmpModelRoll < 50 Then
            lStrafeModelID = 47
        ElseIf lTmpModelRoll < 75 Then
            lStrafeModelID = 72
        Else
            lStrafeModelID = 2
        End If

        lTmpModelRoll = CInt(Rnd() * 100)
        If lTmpModelRoll < 15 Then
            lBigShipModelID = 5
        ElseIf lTmpModelRoll < 30 Then
            lBigShipModelID = 44
        ElseIf lTmpModelRoll < 45 Then
            lBigShipModelID = 61
        ElseIf lTmpModelRoll < 60 Then
            lBigShipModelID = 75
        ElseIf lTmpModelRoll < 75 Then
            lBigShipModelID = 81
        Else : lBigShipModelID = 87
        End If

        lTmpModelRoll = CInt(Rnd() * 100)
        If lTmpModelRoll < 10 Then
            lCorvetteModelID = 4
        ElseIf lTmpModelRoll < 20 Then
            lCorvetteModelID = 36
        ElseIf lTmpModelRoll < 30 Then
            lCorvetteModelID = 53
        ElseIf lTmpModelRoll < 40 Then
            lCorvetteModelID = 55
        ElseIf lTmpModelRoll < 50 Then
            lCorvetteModelID = 48
        ElseIf lTmpModelRoll < 60 Then
            lCorvetteModelID = 56
        ElseIf lTmpModelRoll < 70 Then
            lCorvetteModelID = 49
        ElseIf lTmpModelRoll < 80 Then
            lCorvetteModelID = 52
        ElseIf lTmpModelRoll < 90 Then
            lCorvetteModelID = 80
        Else : lCorvetteModelID = 58
        End If

        moBSEntity = New BaseEntity
        With moBSEntity
            .oMesh = goResMgr.GetMesh(lBigShipModelID)
            .LocX = -900
            .LocY = -400
            .LocZ = -9500
            .LocAngle = 2750
            '.fMapWrapLocX = .LocX
            .bCulled = False
            .yArmorHP(0) = 0
            .yArmorHP(1) = 100
            .yArmorHP(2) = 30
            .yArmorHP(3) = 20
            .yShieldHP = 0
            .yStructureHP = 90
            .TestForBurnFX()
            .yVisibility = eVisibilityType.Visible
            .DestX = 0 : .DestY = 0 : .DestZ = 0
            .CurrentStatus = elUnitStatus.eUnitMoving
            .InitializeEngineFX()
        End With
        moBSEntity2 = New BaseEntity
        With moBSEntity2
            .oMesh = goResMgr.GetMesh(lBigShipModelID)
            .LocX = 900
            .LocY = -700
            .LocZ = -9800
            .LocAngle = 2700
            '.fMapWrapLocX = .LocX
            .bCulled = False
            .yArmorHP(0) = 100
            .yArmorHP(1) = 100
            .yArmorHP(2) = 100
            .yArmorHP(3) = 100
            .yShieldHP = 50
            .yStructureHP = 100
            .TestForBurnFX()
            .yVisibility = eVisibilityType.Visible
            .LocX = 900
            .LocY = -700
            .LocZ = -9800
            .LocAngle = 2700
            '.fMapWrapLocX = .LocX
            .DestX = 0 : .DestY = 0 : .DestZ = 0
            .CurrentStatus = elUnitStatus.eUnitMoving
            .InitializeEngineFX()
            .iCombatTactics = 0
            .VelX = 1.3F
        End With
        moStationEntity = New BaseEntity
        With moStationEntity
            .oMesh = goResMgr.GetMesh(26)
            .LocX = 3000
            .LocY = -1500
            .LocZ = 2300
            .LocAngle = 0
            '.fMapWrapLocX = .LocX
            .bCulled = False
            .yArmorHP(0) = 30
            .yArmorHP(1) = 30
            .yArmorHP(2) = 30
            .yArmorHP(3) = 30
            .TestForBurnFX()
            .yStructureHP = 100
            .yShieldHP = 0
            .yVisibility = eVisibilityType.Visible
        End With
        moBSQuickDie = New BaseEntity
        With moBSQuickDie
            .oMesh = goResMgr.GetMesh(lBigShipModelID)
            .LocX = 4500
            .LocY = 0
            .LocZ = 4000
            .LocAngle = 1000 '1540
            '.fMapWrapLocX = .LocX
            .bCulled = False
            .yArmorHP(0) = 0
            .yArmorHP(1) = 0
            .yArmorHP(2) = 0
            .yArmorHP(3) = 0
            .yStructureHP = 1
            .yShieldHP = 0
            .TestForBurnFX()
            .yVisibility = eVisibilityType.Visible
            .CurrentStatus = elUnitStatus.eUnitMoving
            .InitializeEngineFX()
            .LastUpdateCycle = Int32.MinValue ' timeGetTime
        End With
        moCruiser1 = New BaseEntity
        With moCruiser1
            .oMesh = goResMgr.GetMesh(lCorvetteModelID)
            .LocX = 4500
            .LocY = 500
            .LocZ = 5000
            .LocAngle = 900
            '.fMapWrapLocX = .LocX
            .bCulled = False
            .yArmorHP(0) = 90
            .yArmorHP(1) = 100
            .yArmorHP(2) = 100
            .yArmorHP(3) = 90
            .yStructureHP = 100
            .yShieldHP = 0
            .TestForBurnFX()
            .yVisibility = eVisibilityType.Visible
            .CurrentStatus = elUnitStatus.eUnitMoving
            .InitializeEngineFX()
            .LastUpdateCycle = Int32.MinValue 'timeGetTime
            .iCombatTactics = 0
        End With
        moCruiser2 = New BaseEntity
        With moCruiser2
            .oMesh = goResMgr.GetMesh(lCorvetteModelID)
            .LocX = 4000
            .LocY = 700
            .LocZ = 5000
            .LocAngle = 900
            '.fMapWrapLocX = .LocX
            .bCulled = False
            .yArmorHP(0) = 100
            .yArmorHP(1) = 100
            .yArmorHP(2) = 100
            .yArmorHP(3) = 100
            .yStructureHP = 100
            .yShieldHP = 0
            .TestForBurnFX()
            .yVisibility = eVisibilityType.Visible
            .CurrentStatus = elUnitStatus.eUnitMoving
            .InitializeEngineFX()
            .LastUpdateCycle = Int32.MinValue 'timeGetTime
            .iCombatTactics = 0
        End With

        ReDim moFtrChase(1)
        For X As Int32 = 0 To moFtrChase.GetUpperBound(0)
            moFtrChase(X) = New BaseEntity
            With moFtrChase(X)
                If X = 1 Then .oMesh = goResMgr.GetMesh(lChaserModelID) Else .oMesh = goResMgr.GetMesh(lChasedModelID)
                .LocX = 200 - (100 * X)
                '.LocY = 120 '+ (20 * X)
                If X = 0 Then .LocY = 120 Else .LocY = 200 '-120
                .LocZ = -(10001 + (1650 * X))
                .LocAngle = 2700
                If X = 0 Then .LocYaw = 450 Else .LocYaw = 50

                '.fMapWrapLocX = .LocX
                .bCulled = False
                .yArmorHP(0) = 100
                .yArmorHP(1) = 100
                .yArmorHP(2) = 100
                .yArmorHP(3) = 100
                .yStructureHP = 100
                .yShieldHP = 0
                .TestForBurnFX()
                .yVisibility = eVisibilityType.Visible
                .CurrentStatus = elUnitStatus.eUnitMoving
                .InitializeEngineFX()
                .LastUpdateCycle = Int32.MinValue 'timeGetTime
                .iCombatTactics = 0
                .lChildUB = -1
                .MaxColonists = -1
                .ObjTypeID = -6
            End With
        Next X

        ReDim moFtrStrafe(2)
        For X As Int32 = 0 To moFtrStrafe.GetUpperBound(0)
            moFtrStrafe(X) = New BaseEntity
            With moFtrStrafe(X)
                .ObjTypeID = ObjectType.eUnit
                .oMesh = goResMgr.GetMesh(lStrafeModelID)
                If X = 1 Then .LocX = 2400 Else .LocX = 2200
                .LocY = (20 * X)
                If X = 2 Then .LocZ = -9401 Else .LocZ = -9001
                .LocAngle = 2000
                .LocYaw = 450S - CShort(Rnd() * 20)
                '.fMapWrapLocX = .LocX
                .bCulled = False
                For Y As Int32 = 0 To 3
                    .yArmorHP(Y) = 100
                Next
                .yStructureHP = 100
                .yShieldHP = 100
                .TestForBurnFX()
                .yVisibility = eVisibilityType.Visible
                .CurrentStatus = elUnitStatus.eUnitMoving
                .InitializeEngineFX()
                .LastUpdateCycle = Int32.MinValue
                .iCombatTactics = 0
            End With
        Next X

        If moFtrLow Is Nothing Then
            moFtrLow = goResMgr.LoadScratchMeshNoTextures("SmallFtr2_Low.x", "Login.pak")
            moFtrTex = goResMgr.GetTexture("ff06.bmp", GFXResourceManager.eGetTextureType.ModelTexture)

            For X As Int32 = 0 To muUnits.GetUpperBound(0)
                With muUnits(X)
                    .fTurnAngle = Rnd() * 360.0F
                    .fTurnVel = (Rnd() - 0.5F) * 2
                    .fYLoc = Rnd() * 220
                    .fYVel = (Rnd() * 1.0F - 0.5F) * 2
                End With
            Next X
        End If


        ReDim moExplosion(14)

        For X As Int32 = 0 To moExplosion.GetUpperBound(0)
            moExplosion(X) = New LoginScreenExplosion(New Vector3((Rnd() * 2500) + 2500, (Rnd() * -2500) - 1000, (Rnd() * 2500) + 2500))
            moExplosion(X).EmitterStopped = X < moExplosion.GetUpperBound(0) \ 2

            AddHandler moExplosion(X).ExplosionComplete, AddressOf ExplosionComplete
        Next X
    End Sub

    Private Function GetFighterMatrix(ByVal lIndex As Int32) As Matrix
        Dim fLocAngle As Single
        Dim fLocYaw As Single
        Dim fLocPitch As Single

        With muUnits(lIndex)
            If .fTurnVel > 0.0F Then
                fLocAngle = .fTurnAngle '+ 90
                fLocYaw = 45
            Else
                fLocAngle = .fTurnAngle - 180 '0
                fLocYaw = -45
            End If

            If .fYVel > 0.0F Then
                fLocPitch = 20
            Else : fLocPitch = -20
            End If
        End With

        fLocAngle = DegreeToRadian(fLocAngle)
        fLocYaw = DegreeToRadian(fLocYaw)
        fLocPitch = DegreeToRadian(fLocPitch)

        'roll = yaw
        'yaw = angle
        'pitch = pitch

        Dim matWorld As Matrix = Matrix.Identity
        'matWorld.Multiply(Matrix.Scaling(0.1, 0.1, 0.1))
        matWorld.Multiply(Matrix.RotationYawPitchRoll(fLocAngle, fLocPitch, fLocYaw))

        Return matWorld
        'Dim lAdjAngle As Int32 = LocAngle
        'Dim fX1 As Single = 1.0F
        'Dim fZ1 As Single = 0.0F
        'Call RotatePoint(0, 0, fX1, fZ1, lAdjAngle / 10.0F)

        ''Now, get our angle...
        'Dim vN As Vector3 = CType(goCurrentEnvir.oGeoObject, Planet).GetTriangleNormal(LocX, LocZ)
        'Dim vU As Vector3 = New Vector3(0, 1, 0)
        'Dim vF As Vector3 = New Vector3(fX1, 0, fZ1)
        'Dim vcpNU As Vector3 = Vector3.Cross(vU, vN)
        'Dim fdpNU As Single = Vector3.Dot(vU, vN)
        'fYaw = RadianToDegree(CSng(-Math.Acos(Vector3.Dot(vF, vU))))

        'If fYaw = 90 Then fYaw = 0

        'If lAdjAngle < 0 Then lAdjAngle += 3600
        'fYaw += (lAdjAngle / 10.0F)
        'If fYaw < 0 Then fYaw += 360
        'If fYaw > 360 Then fYaw -= 360

        ''Now, assign our values...
        'If fYaw > 315 Or fYaw < 45 Then
        '    fPitch = vcpNU.X
        '    fRoll = vcpNU.Z
        'ElseIf fYaw < 135 Then
        '    fPitch = -vcpNU.Z
        '    fRoll = vcpNU.X
        'ElseIf fYaw < 225 Then
        '    fPitch = -vcpNU.X
        '    fRoll = -vcpNU.Z
        'Else
        '    fPitch = vcpNU.Z
        '    fRoll = -vcpNU.X
        'End If
        'fYaw = DegreeToRadian(fYaw)
    End Function

    Private Function GetFighterEntityMatrix(ByRef oEntity As BaseEntity) As Matrix
        Dim fLocAngle As Single
        Dim fLocYaw As Single
        Dim fLocPitch As Single

        With oEntity
            fLocAngle = .LocAngle + 900.0F
            If fLocAngle < 0 Then fLocAngle += 3600
            fLocAngle = DegreeToRadian(fLocAngle / 10.0F)
            fLocYaw = DegreeToRadian(.LocYaw / 10.0F)
            fLocPitch = DegreeToRadian(.ObjectID / 10.0F)
        End With

        fLocAngle = DegreeToRadian(fLocAngle)
        fLocYaw = DegreeToRadian(fLocYaw)
        fLocPitch = DegreeToRadian(fLocPitch)

        'roll = yaw
        'yaw = angle
        'pitch = pitch

        Dim matWorld As Matrix = Matrix.Identity
        'matWorld.Multiply(Matrix.Scaling(0.1, 0.1, 0.1))
        matWorld.Multiply(Matrix.RotationYawPitchRoll(fLocAngle, fLocPitch, fLocYaw))

        Dim matTemp As Matrix = Matrix.Identity
        matTemp.Translate(oEntity.LocX, oEntity.LocY, oEntity.LocZ)
        matWorld.Multiply(matTemp)

        Return matWorld
    End Function
 
    Public Sub CleanUpResources()
        If moModelShader Is Nothing = False Then moModelShader.DisposeMe()
        moModelShader = Nothing

    End Sub
End Class