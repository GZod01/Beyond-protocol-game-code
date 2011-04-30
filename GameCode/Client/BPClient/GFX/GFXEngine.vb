Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class GFXEngine

    Private Const ml_WP_IND_SIZE As Int32 = 16

    Public Shared gbPaused As Boolean = False
    Public Shared gbDeviceLost As Boolean = False
    Public Shared gbDeviceLostHard As Boolean = False 'Special boolean used in testing a forced crash -> Recreate Everything

    Public Shared moDevice As Device        'I am the true device, the one and only... all others point to me

    Private mbInitialized As Boolean = False

    Private moHPStat As Sprite

    Private moRadarBlip As Texture
    'Private moRadarMesh As Mesh
    Private matRadar As Material
    Private mlRadarAlpha As Int32 = 0
	Private mlRadarAlphaMod As Int32 = 1
    Private matIronCurtain As Material
    Private moBoxMeshTargets As Mesh
    Private moBoxMeshSelStar As Mesh

    Private moWPInd As Texture

    Private mbSysMiniMapInit As Boolean = False
    Private moSysMiniMapMats() As Material

    Private moPMapRadarTex As Texture
	Private mfFlashVal As Single = 1.0F

	Private moBombMesh As Mesh
	Private moBombTex As Texture
	Private mfBombRot As Single

	Private moLogo As BPLogo

    Public swStartup As Stopwatch
    Public mbBreakout As Boolean = False
    Private mlStartupDlgAlpha As Int32 = 255

    Private moLoginScreen As LoginScreen

    Private msw_Utility As Stopwatch

    Private mlLastSystemMinimapUpdate As Int32 = Int32.MinValue
    Private moSystemMinimapTexture As Texture

    Private moPost As PostShader

    Private Shared mbToggleFullScreen As Boolean = False
    Public Shared bRenderInProgress As Boolean = False
    Public Shared mbCaptureScreenshot As Boolean = False      'for capturing screenshots

    Private mfrmMain As Form = Nothing

    Public Event CriticalFailure()

#Region "  Directional Moving Rendering  "
	Public bInCtrlMove As Boolean = False
	Public lMouseDownX As Int32
	Public lMouseDownZ As Int32
	Public lFormationID As Int32
	Public lMouseDirX As Int32
	Public lMouseDirZ As Int32

	Private Sub RenderDirectionalMoving()

		Dim oFormation As FormationDef = Nothing
		For X As Int32 = 0 To goCurrentPlayer.lFormationUB
			If goCurrentPlayer.lFormationIdx(X) = lFormationID Then
				oFormation = goCurrentPlayer.oFormations(X)
				Exit For
			End If
		Next X
		If oFormation Is Nothing Then Return

		If moBombMesh Is Nothing OrElse moBombMesh.Disposed = True Then
			moBombMesh = goResMgr.LoadScratchMeshNoTextures("reticle.x", "misc.pak")
		End If
		If moBombTex Is Nothing OrElse moBombTex.Disposed = True Then
			moBombTex = goResMgr.GetTexture("reticle.dds", GFXResourceManager.eGetTextureType.ModelTexture, "ob.pak")
		End If

        Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
        Dim lPrevLighting As Boolean = moDevice.RenderState.Lighting
        Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
        Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
        Dim lPrevAlplaBlnd As Boolean = moDevice.RenderState.AlphaBlendEnable

        With moDevice.RenderState
            .SourceBlend = Blend.SourceAlpha
            .DestinationBlend = Blend.InvSourceAlpha
            .AlphaBlendEnable = True
            .ZBufferEnable = False
            .Lighting = False
        End With

		'ok, now, place it...
		Dim matWorld As Matrix = Matrix.Identity

		moDevice.SetTexture(0, moBombTex)
		Dim oMat As Material
		With oMat
			.Ambient = Color.White
			.Diffuse = Color.White
			.Specular = Color.Black
		End With
		moDevice.Material = oMat

		Dim oPlanet As Planet = Nothing
		If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView = CurrentView.ePlanetView Then
			Try
				oPlanet = CType(goCurrentEnvir.oGeoObject, Planet)
			Catch
			End Try
		End If

        Dim fAngle As Single = LineAngleDegrees(lMouseDownX, lMouseDownZ, lMouseDirX, lMouseDirZ)
        fAngle -= 180.0F
        If fAngle < 0 Then fAngle += 360.0F
		Dim lHalfVal As Int32 = UIFormation.ml_GRID_SIZE_WH \ 2

		'Now, let's place our stuff...
		For X As Int32 = 0 To oFormation.lLocUB

			Dim fPtX As Single = ((oFormation.ptLocs(X).X - lHalfVal) * oFormation.lCellSize * gl_FINAL_GRID_SQUARE_SIZE)
			Dim fPtZ As Single = ((oFormation.ptLocs(X).Y - lHalfVal) * oFormation.lCellSize * gl_FINAL_GRID_SQUARE_SIZE)

			RotatePoint(0, 0, fPtX, fPtZ, fAngle)
			fPtX += lMouseDownX
			fPtZ += lMouseDownZ

			Dim vecLoc As Vector3 = New Vector3(fPtX, 0, fPtZ)
			If oPlanet Is Nothing = False Then
				vecLoc.Y = oPlanet.GetHeightAtPoint(fPtX, fPtZ, True) + 1000
			End If

			matWorld = Matrix.Identity
			matWorld.Multiply(Matrix.Scaling(2, 2, 2))
			matWorld.Multiply(Matrix.Translation(vecLoc))
			moDevice.Transform.World = matWorld
			moBombMesh.DrawSubset(0)
		Next X

        With moDevice.RenderState
            .ZBufferEnable = lPrevZBuffer
            .Lighting = lPrevLighting
            .SourceBlend = lPrevSrcBlnd
            .DestinationBlend = lPrevDestBlnd
            .AlphaBlendEnable = lPrevAlplaBlnd
        End With
    End Sub
#End Region

    Public Function RecreateEverything(ByRef ofrm As Form, ByRef oMsgSys As MsgSystem, ByVal bRecreate As Boolean) As Boolean
        'ok, go thru everything...
        Try
            ' SolarSystem.shared resources
            SolarSystem.ResetSkyBox()
            SolarSystem.ResetCosmoSphere()

            ' Planet.everything (some shared, some on a per-planet basis)
            Planet.ReleaseSprites()
            If goGalaxy Is Nothing = False Then
                For X As Int32 = 0 To goGalaxy.mlSystemUB
                    Dim oSys As SolarSystem = goGalaxy.moSystems(X)
                    If oSys Is Nothing = False Then
                        If oSys.moPlanets Is Nothing = False Then
                            For Y As Int32 = 0 To oSys.PlanetUB
                                If oSys.moPlanets(Y) Is Nothing = False Then
                                    oSys.moPlanets(Y).ReleaseRenderTargets()
                                End If
                            Next Y
                        End If
                    End If
                Next X
            End If
            TerrainClass.ReleaseVisibleTexture()
            TerrainClass.ResetShaders()

            ' StarType.Shared resources and (per startype basis) moTexture and moStarMesh
            StarType.ReleaseSprite()
            For X As Int32 = 0 To glStarTypeUB
                If goStarTypes(X) Is Nothing = False Then
                    goStarTypes(X).ClearResources()
                End If
            Next X

            ' every entity's HPTexture, moProducingMesh
            If goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            If oEntity.oMesh Is Nothing = False Then
                                oEntity.oMesh.CleanMe()
                            End If
                        End If
                    End If
                Next X
            End If

            ' BPFont stuff
            BPFont.ClearAllFonts()

            ' BPLine stuff
            BPLine.ClearLine()

            ' me.moHPStat, moRadarBlip, moRadarMesh, moWPInd, moPMapRadarTex, moBombMesh, moBombTex, moSystemMinimapTexture, moPost, moBoxMesh
            If moHPStat Is Nothing = False Then moHPStat.Dispose()
            moHPStat = Nothing
            If moRadarBlip Is Nothing = False Then moRadarBlip.Dispose()
            moRadarBlip = Nothing

            If moWPInd Is Nothing = False Then moWPInd.Dispose()
            moWPInd = Nothing
            If moPMapRadarTex Is Nothing = False Then moPMapRadarTex.Dispose()
            moPMapRadarTex = Nothing
            If moBombMesh Is Nothing = False Then moBombMesh.Dispose()
            moBombMesh = Nothing
            If moSystemMinimapTexture Is Nothing = False Then moSystemMinimapTexture.Dispose()
            moSystemMinimapTexture = Nothing
            If moPost Is Nothing = False Then moPost.DisposeMe()
            moPost = Nothing
            If moModelShader Is Nothing = False Then moModelShader.DisposeMe()
            moModelShader = Nothing

            If moBoxMeshTargets Is Nothing = False Then moBoxMeshTargets.Dispose()
            moBoxMeshTargets = Nothing

            If moBoxMeshSelStar Is Nothing = False Then moBoxMeshSelStar.Dispose()
            moBoxMeshSelStar = Nothing

            ' MineralCache.CacheMesh(), texCache, DebrisMesh, DebrisTex()
            MineralCache.Clear3DResources()

            ' all of gfxresourcemanager
            If goResMgr Is Nothing = False Then
                goResMgr.ClearAllResources()
            End If
            goResMgr = Nothing

            ' BPLogo.moBeyondTex, moProtocolTex, moBPTex, moSprite
            If moLogo Is Nothing = False Then
                moLogo.ReleaseSprite()
            End If
            moLogo = Nothing

            ' BurnFX
            If goPFXEngine32 Is Nothing = False Then goPFXEngine32.ClearAllEmitters()
            goPFXEngine32 = Nothing

            If goEntityDeath Is Nothing = False Then goEntityDeath.ClearAll()
            DeathSequence.ClearShared()
            goEntityDeath = Nothing

            'If goRewards Is Nothing = False Then goRewards.ClearAll()
            'goRewards = Nothing

            ' EntityBurnMarkManager.moBurnTex (shared)
            If EntityBurnMarkManager.moBurnTex Is Nothing = False Then EntityBurnMarkManager.moBurnTex.Dispose()
            EntityBurnMarkManager.moBurnTex = Nothing

            ' ExplosionManager.moExplTex()
            goExplMgr.RemoveAll()
            goExplMgr.ClearResources()
            goExplMgr = Nothing

            ' FireworksMgr.moTex
            If goFireworks Is Nothing = False Then goFireworks.ClearTex()
            goFireworks = Nothing

            ' MissileMgr.moTex
            If goMissileMgr Is Nothing = False Then
                goMissileMgr.KillAll()
                goMissileMgr.ClearTex()
            End If
            goMissileMgr = Nothing

            ' PlanetPFX.moTex (shared)
            If PlanetFX.ParticleEngine.moTex Is Nothing = False Then PlanetFX.ParticleEngine.moTex.Dispose()
            PlanetFX.ParticleEngine.moTex = Nothing

            ' ShieldFXManager.moTexture()
            If goShldMgr Is Nothing = False Then goShldMgr.TurnOffMgr()
            ShieldFXManager.ClearTextures()
            goShldMgr = Nothing

            ' WormholeManager.moParticle, moWormholeCloudTex
            If goWormholeMgr Is Nothing = False Then
                goWormholeMgr.ClearAllWormholes()
                goWormholeMgr.ClearTextures()
            End If
            goWormholeMgr = Nothing

            ' WpnFXManager.moTexture (shared)
            If goWpnMgr Is Nothing = False Then goWpnMgr.CleanAll()
            WpnFXManager.moTexture = Nothing
            goWpnMgr = Nothing

            ' ctlCalendar.moSprite (shared), moTex
            ctlCalendar.ReleaseDefaultPool()

            If goUILib Is Nothing = False Then
                ' UIControl.oCanvas of any control with a canvas
                ' Close All Windows
                ' UIListBox.oIconTexture
                goUILib.RemoveAllWindows()
                ' UILib.Pen, oDevice, moInterfaceTexture, BuildGhost, 
                '     Cleanup FocusedControl, moToolTip, moMsgBox, moSelection, moSingleSel, moMultiSel, moAdvDisplay, oScrollingBar, oTechProp, oScrollingRel, oButtonDown, CurrentComboBoxSelected
                goUILib.ReleaseInterfaceTextures()
            End If
            goUILib = Nothing

            ' UIHullSlots.moHullTex (Shared)
            If UIHullSlots.moHullTex Is Nothing = False Then UIHullSlots.moHullTex.Dispose()
            UIHullSlots.moHullTex = Nothing
            Try
                If moDevice Is Nothing = False Then moDevice.Dispose()
            Catch
            End Try
            moDevice = Nothing

            If bRecreate = False Then Return True

            'Ok, initialize Direct3D
            If InitD3D(ofrm) = False Then Return False
            goResMgr = New GFXResourceManager()
            goUILib = New UILib(moDevice)
            goUILib.SetMsgSystem(oMsgSys)

            If glCurrentEnvirView = CurrentView.eStartupLogin Then
                InitializeLoginScreen()
            Else
                'Show our Chat window...
                Dim oTmpCht As frmChat
                oTmpCht = New frmChat(goUILib)
                goUILib.AddWindow(oTmpCht)
                oTmpCht = Nothing

                'Create a new QuickBar
                Dim ofrmQB As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                If ofrmQB Is Nothing Then ofrmQB = New frmQuickBar(goUILib)
                ofrmQB = Nothing

                'Create a new Envirdisplay
                Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
                If ofrmED Is Nothing Then ofrmED = New frmEnvirDisplay(goUILib)
                ofrmED.Visible = True
                ofrmED = Nothing

                If muSettings.ShowConnectionStatus = True Then
                    Dim ofrmCS As frmConnectionStatus = CType(goUILib.GetWindow("frmConnectionStatus"), frmConnectionStatus)
                    If ofrmCS Is Nothing Then ofrmCS = New frmConnectionStatus(goUILib)
                    ofrmCS.Visible = True
                    ofrmCS = Nothing
                End If

            End If

            'go through our current envir and recreate
            If goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            oEntity.oMesh = goResMgr.GetMesh(oEntity.oUnitDef.ModelID)
                        End If
                    End If
                Next X
            End If

            Return True
        Catch ex As Exception
            LogCrashEvent(ex, True, False, "RecreateEverything", True, oMsgSys)
            Return False
        End Try
    End Function

	'Private Sub SpriteDispose(ByVal sender As Object, ByVal e As EventArgs)
	'    moHPStat = Nothing
	'End Sub

	'Private Sub SpriteLost(ByVal sender As Object, ByVal e As EventArgs)
	'    If moHPStat Is Nothing = False Then moHPStat.Dispose()
	'    moHPStat = Nothing
	'End Sub

    'Public Function GetDevice() As Device
    '    Return moDevice
    'End Function

    Private mbSaveRSOnce As Boolean = False
    Private Sub SaveRS()
        If mbSaveRSOnce = True Then
            mbSaveRSOnce = False
            With moDevice.RenderState
                Debug.WriteLine("AlphaBlendEnable=" & .AlphaBlendEnable)
                Debug.WriteLine("AlphaBlendOperation=" & .AlphaBlendOperation)
                Debug.WriteLine("AlphaDestinationBlend=" & .AlphaDestinationBlend)
                Debug.WriteLine("AlphaSourceBlend=" & .AlphaSourceBlend)
                Debug.WriteLine("BlendOperation=" & .BlendOperation)
                Debug.WriteLine("DestinationBlend=" & .DestinationBlend)
                Debug.WriteLine("SourceBlend=" & .SourceBlend)
            End With
        End If
    End Sub

#Region "  Model Rendering  "

    Private Sub SetModelTextureStates()
        If moRelTex Is Nothing Then
            ReDim moRelTex(4)
            moRelTex(0) = goResMgr.GetTexture("Neutral.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "textures2.pak")
            moRelTex(1) = goResMgr.GetTexture("Player.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "textures2.pak")
            moRelTex(2) = goResMgr.GetTexture("Ally.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "textures2.pak")
            moRelTex(3) = goResMgr.GetTexture("Enemy.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "textures2.pak")
            moRelTex(4) = goResMgr.GetTexture("Guild.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "gi.pak")
        End If
        moDevice.RenderState.AlphaBlendEnable = True

        'SaveRS()
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
    End Sub
    Private Sub ResetModelTextureStates()
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
    End Sub

    Public Shared mbSupportsNewModelMethod As Boolean = False
	Private moRelTex() As Texture = Nothing
	Public Enum RenderModelType As Byte
		eNoSpecial = 0
		eOldData = 1
        eIronCurtain = 2
        eSelected = 4
    End Enum
    Private mclrSelectedEmissive As System.Drawing.Color
    Private mlLastSelectedEmissiveSet As Int32 = 0
    Private mlSelectedEmissiveChng As Int32 = 1
    Private mlSelectedEmissive As Int32 = 0
	Private Sub RenderModel(ByVal oObj As RenderObject, ByVal lTexMod As Int32, ByVal ySpecialType As RenderModelType)
		Dim X As Int32
        Try
            'TODO: Possibly handle this better
            If oObj.oMesh Is Nothing Then Return
            If oObj.oMesh.oMesh Is Nothing Then Return

            'TODO: This can probably be optimized by removing the Material assignments and Texture sets...
            '   This assumes that our models will use a single texture/material

            'If bAsOldData = True Then moDevice.Transform.World = oObj.CurrentWorldMatrix Else moDevice.Transform.World = oObj.GetWorldMatrix
            If (ySpecialType And RenderModelType.eOldData) <> 0 Then
                moDevice.Transform.World = oObj.CurrentWorldMatrix
            Else : moDevice.Transform.World = oObj.GetWorldMatrix
            End If

            Dim bGetMesh As Boolean = False
            With oObj.oMesh
                If .oMesh Is Nothing Then Return
                If .oMesh.Device Is moDevice = False Then
                    bGetMesh = True
                End If
            End With
            If bGetMesh = True Then oObj.oMesh = goResMgr.GetMesh(oObj.oMesh.lModelID)

            If mbSupportsNewModelMethod = True Then
                With oObj.oMesh

                    Dim oTmpMat As Material
                    If (ySpecialType And RenderModelType.eIronCurtain) <> 0 Then
                        oTmpMat = matIronCurtain
                    Else : oTmpMat = .Materials(0)
                    End If

                    If muSettings.FlashSelections = True AndAlso (ySpecialType And RenderModelType.eSelected) <> 0 Then
                        If mlLastSelectedEmissiveSet <> glCurrentCycle Then
                            mlLastSelectedEmissiveSet = glCurrentCycle
                            mlSelectedEmissive += mlSelectedEmissiveChng
                            If mlSelectedEmissive < Math.Max(muSettings.AmbientLevel, 0) Then
                                mlSelectedEmissive = Math.Max(muSettings.AmbientLevel, 0)
                                mlSelectedEmissiveChng = muSettings.FlashRate
                            ElseIf mlSelectedEmissive > 75 Then
                                mlSelectedEmissive = 75
                                mlSelectedEmissiveChng = -muSettings.FlashRate
                            End If
                            mclrSelectedEmissive = System.Drawing.Color.FromArgb(255, mlSelectedEmissive, mlSelectedEmissive, mlSelectedEmissive)
                        End If
                        oTmpMat.Emissive = mclrSelectedEmissive
                    End If
                    moDevice.Material = oTmpMat

                    moDevice.SetTexture(0, moRelTex(lTexMod))
                    For X = 0 To .NumOfMaterials - 1
                        moDevice.SetTexture(1, .Textures(X))
                        .oMesh.DrawSubset(X)
                    Next X

                    'Now, check for a turret...
                    If .bTurretMesh = True Then
                        'If bAsOldData = True Then moDevice.Transform.World = oObj.CurrentTurretMatrix Else moDevice.Transform.World = oObj.GetTurretMatrix
                        If (ySpecialType And RenderModelType.eOldData) <> 0 Then
                            moDevice.Transform.World = oObj.CurrentTurretMatrix
                        Else : moDevice.Transform.World = oObj.GetTurretMatrix
                        End If

                        'NOTE: Only allowed one material for the turret...
                        moDevice.SetTexture(1, .Textures(0))
                        .oTurretMesh.DrawSubset(0)
                    End If
                End With
            Else
                With oObj.oMesh
                    Dim oTmpMat As Material
                    If (ySpecialType And RenderModelType.eIronCurtain) <> 0 Then
                        oTmpMat = matIronCurtain
                    Else : oTmpMat = .Materials(0)
                    End If

                    If muSettings.FlashSelections = True AndAlso (ySpecialType And RenderModelType.eSelected) <> 0 Then
                        If mlLastSelectedEmissiveSet <> glCurrentCycle Then
                            mlLastSelectedEmissiveSet = glCurrentCycle
                            mlSelectedEmissive += mlSelectedEmissiveChng
                            If mlSelectedEmissive < Math.Max(muSettings.AmbientLevel, 0) Then
                                mlSelectedEmissive = Math.Max(muSettings.AmbientLevel, 0)
                                mlSelectedEmissiveChng = muSettings.FlashRate
                            ElseIf mlSelectedEmissive > 75 Then
                                mlSelectedEmissive = 75
                                mlSelectedEmissiveChng = -muSettings.FlashRate
                            End If
                            mclrSelectedEmissive = System.Drawing.Color.FromArgb(255, mlSelectedEmissive, mlSelectedEmissive, mlSelectedEmissive)
                        End If
                        oTmpMat.Emissive = mclrSelectedEmissive
                    End If
                    moDevice.Material = oTmpMat

                    For X = 0 To .NumOfMaterials - 1
                        moDevice.SetTexture(0, .Textures((X * 4) + lTexMod))
                        .oMesh.DrawSubset(X)
                    Next X

                    'Now, check for a turret...
                    If .bTurretMesh = True Then
                        'If bAsOldData = True Then moDevice.Transform.World = oObj.CurrentTurretMatrix Else moDevice.Transform.World = oObj.GetTurretMatrix
                        If (ySpecialType And RenderModelType.eOldData) <> 0 Then
                            moDevice.Transform.World = oObj.CurrentTurretMatrix
                        Else : moDevice.Transform.World = oObj.GetTurretMatrix
                        End If

                        'NOTE: Only allowed one material for the turret...
                        moDevice.SetTexture(0, .Textures(lTexMod))
                        .oTurretMesh.DrawSubset(0)
                    End If
                End With
            End If
        Catch
            If oObj Is Nothing = False AndAlso oObj.oMesh Is Nothing = False Then
                oObj.oMesh = goResMgr.GetMesh(oObj.oMesh.lModelID)
            Else
                oObj.oMesh = Nothing
            End If
        End Try
    End Sub

	Private moModelShader As ModelShader
#End Region

    'Public Sub Render(ByVal oObj As RenderObject, ByVal lTexMod As Int32)
    '    Dim X As Int32

    '    'TODO: Possibly handle this better
    '    If oObj.oMesh Is Nothing Then Exit Sub
    '    If oObj.oMesh.oMesh Is Nothing Then Exit Sub

    '    'TODO: This can probably be optimized by removing the Material assignments and Texture sets...
    '    '   This assumes that our models will use a single texture/material

    '    moDevice.Transform.World = oObj.GetWorldMatrix()
    '    With oObj.oMesh
    '        For X = 0 To .NumOfMaterials - 1
    '            moDevice.Material = .Materials(X)
    '            moDevice.SetTexture(0, .Textures((X * 4) + lTexMod))
    '            .oMesh.DrawSubset(X)
    '        Next X

    '        'Now, check for a turret...
    '        If .bTurretMesh = True Then
    '            moDevice.Transform.World = oObj.GetTurretMatrix()
    '            'NOTE: Only allowed one material for the turret...
    '            moDevice.Material = .Materials(0)
    '            moDevice.SetTexture(0, .Textures(lTexMod))
    '            .oTurretMesh.DrawSubset(0)
    '        End If
    '    End With

    'End Sub

    'Public Sub RenderOldData(ByVal oObj As RenderObject, ByVal lTexMod As Int32)
    '    Dim X As Int32

    '    'TODO: Possibly handle this better
    '    If oObj.oMesh Is Nothing Then Exit Sub
    '    If oObj.oMesh.oMesh Is Nothing Then Exit Sub

    '    'TODO: This can probably be optimized by removing the Material assignments and Texture sets...
    '    '   This assumes that our models will use a single texture/material

    '    moDevice.Transform.World = oObj.CurrentWorldMatrix
    '    With oObj.oMesh
    '        For X = 0 To .NumOfMaterials - 1
    '            moDevice.Material = .Materials(X)
    '            moDevice.SetTexture(0, .Textures((X * 4) + lTexMod))
    '            .oMesh.DrawSubset(X)
    '        Next X

    '        'Now, check for a turret...
    '        If .bTurretMesh = True Then
    '            moDevice.Transform.World = oObj.CurrentTurretMatrix
    '            'NOTE: Only allowed one material for the turret...
    '            moDevice.Material = .Materials(0)
    '            moDevice.SetTexture(0, .Textures(lTexMod))
    '            .oTurretMesh.DrawSubset(0)
    '        End If
    '    End With

    'End Sub

    Public Shared bNVidiaCard As Boolean = False

    Public Function InitD3D(ByRef frmForm As Form) As Boolean
        Dim uParms As PresentParameters
        Dim bRes As Boolean

        mfrmMain = frmForm

        muSettings.LoadSettings()

        'Debug.WriteLine("InitD3D in thread: " & Threading.Thread.CurrentThread.ManagedThreadId.ToString)

        Try
            Dim lCreateFlags As CreateFlags
            Dim bReduced As Boolean = False
            uParms = CreatePresentationParams(mfrmMain)
            If uParms Is Nothing Then Return False

            'lCreateFlags = CType(Val(oINI.GetString("GRAPHICS", "VertexProcessing", CInt(CreateFlags.HardwareVertexProcessing).ToString)), CreateFlags)
            lCreateFlags = CType(muSettings.VertexProcessing, CreateFlags)

            If lCreateFlags <> CreateFlags.HardwareVertexProcessing AndAlso lCreateFlags <> CreateFlags.MixedVertexProcessing AndAlso lCreateFlags <> CreateFlags.SoftwareVertexProcessing Then
                lCreateFlags = CreateFlags.HardwareVertexProcessing
            End If

            Dim uDevCaps As Caps = Manager.GetDeviceCaps(Manager.Adapters.Default.Adapter, DeviceType.Hardware)
            If uDevCaps.DeviceCaps.SupportsHardwareTransformAndLight = False Then
                If lCreateFlags = CreateFlags.HardwareVertexProcessing Then
                    lCreateFlags = CreateFlags.SoftwareVertexProcessing
                    bReduced = True
                End If
            End If

            Try
                moDevice = New Device(0, DeviceType.Hardware, frmForm.Handle, lCreateFlags, uParms)
            Catch
                bReduced = True
                If (lCreateFlags And CreateFlags.HardwareVertexProcessing) = CreateFlags.HardwareVertexProcessing Then
                    lCreateFlags = (lCreateFlags Xor CreateFlags.HardwareVertexProcessing) Or CreateFlags.MixedVertexProcessing
                    Try
                        moDevice = New Device(0, DeviceType.Hardware, frmForm.Handle, lCreateFlags, uParms)
                    Catch
                        lCreateFlags = (lCreateFlags Xor CreateFlags.MixedVertexProcessing) Or CreateFlags.SoftwareVertexProcessing
                        Try
                            moDevice = New Device(0, DeviceType.Hardware, frmForm.Handle, lCreateFlags, uParms)
                        Catch ex As Exception
                            MsgBox("Beyond Protocol will not run on this computer because the video hardware is not sufficient.", MsgBoxStyle.OkOnly, "Error")
                            End
                        End Try
                    End Try
                ElseIf (lCreateFlags And CreateFlags.MixedVertexProcessing) = CreateFlags.MixedVertexProcessing Then
                    lCreateFlags = (lCreateFlags Xor CreateFlags.MixedVertexProcessing) Or CreateFlags.SoftwareVertexProcessing
                    Try
                        moDevice = New Device(0, DeviceType.Hardware, frmForm.Handle, lCreateFlags, uParms)
                    Catch ex As Exception
                        MsgBox("Beyond Protocol will not run on this computer because the video hardware is not sufficient.", MsgBoxStyle.OkOnly, "Error")
                        End
                    End Try
                End If
            End Try

            If bReduced = True Then MsgBox("Beyond Protocol was required to alter settings that will affect performance in order to run the game.", MsgBoxStyle.OkOnly, "Error")
            muSettings.VertexProcessing = CInt(lCreateFlags)

            bRes = Not moDevice Is Nothing

            If muSettings.MiniMapLocX > uParms.BackBufferWidth Then muSettings.MiniMapLocX = uParms.BackBufferWidth - muSettings.MiniMapWidthHeight
            If muSettings.MiniMapLocY > uParms.BackBufferHeight Then muSettings.MiniMapLocY = uParms.BackBufferHeight - muSettings.MiniMapWidthHeight
        Catch
            MsgBox(Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Error in initialize Direct3D")
            bRes = False
        Finally
            uParms = Nothing
        End Try

        If bRes = False Then
            MsgBox("Unable to initialize Direct3D 9.0c. Your video card may not fully support" & vbCrLf & _
                   "DirectX 9.0c. Download the latest drivers for your video card and try" & vbCrLf & _
                   "again. If you feel you should be able to run the game, check the" & vbCrLf & _
                   "'Can You Run It' link on the main page of the website to verify. If you" & vbCrLf & _
                   "pass and are still unable to run the game, contact Technical Support.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Unable to Initialize")
            Return False
        End If

        With moDevice.DeviceCaps.VertexShaderVersion
            If .Major = 1 Then
                'If Not (.Major > 1 OrElse (.Major = 1 AndAlso .Minor = 1)) Then
                muSettings.PostGlowAmt = 0.0F
                muSettings.LightQuality = EngineSettings.LightQualitySetting.VSPS1
            End If
        End With

        mbSupportsNewModelMethod = moDevice.DeviceCaps.TextureOperationCaps.SupportsBlendTextureAlpha = True AndAlso _
            moDevice.DeviceCaps.MaxTextureBlendStages > 2 AndAlso moDevice.DeviceCaps.MaxSimultaneousTextures > 1
        'mbSupportsNewModelMethod = False

        'Instantiate our FX Managers
        If goWpnMgr Is Nothing Then goWpnMgr = New WpnFXManager()
        If goShldMgr Is Nothing Then goShldMgr = New ShieldFXManager()
        If goExplMgr Is Nothing Then goExplMgr = New ExplosionManager() '  New ExplosionFXManager(moDevice)
        If goPFXEngine32 Is Nothing Then goPFXEngine32 = New BurnFX.ParticleEngine(32) '32 for size of points
        If goEntityDeath Is Nothing Then goEntityDeath = New DeathSequenceMgr()
        'If goRewards Is Nothing Then goRewards = New WarpointRewards()
        If goMissileMgr Is Nothing Then goMissileMgr = New MissileMgr()
        If goBurnMarkMgr Is Nothing Then goBurnMarkMgr = New EntityBurnMarkManager()

        InitializeFXColors()

        With matIronCurtain
            .Emissive = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .Diffuse = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .Ambient = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .Specular = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .SpecularSharpness = 10
        End With


        AddHandler moDevice.DeviceLost, AddressOf moDevice_DeviceLost
        AddHandler moDevice.DeviceReset, AddressOf moDevice_DeviceReset
        AddHandler moDevice.Disposing, AddressOf moDevice_Disposing

        mbInitialized = bRes

        Return bRes
    End Function

    Public Function CreatePresentationParams(ByVal frmForm As Form) As PresentParameters
        Dim uParms As PresentParameters = Nothing
        Dim uDispMode As DisplayMode
        Dim bRes As Boolean

        ' muSettings.Windowed = True

        Try
            If moDevice Is Nothing = False Then moDevice.TestCooperativeLevel()
        Catch exDevNotReset As DeviceNotResetException
            'ok, we're ok
        Catch
            Return Nothing
        End Try

        Try
            If muSettings.Windowed = False Then
                mfrmMain.FormBorderStyle = FormBorderStyle.None
            Else : mfrmMain.FormBorderStyle = FormBorderStyle.FixedSingle
            End If
        Catch
        End Try

        Try
            uDispMode = Manager.Adapters.Default.CurrentDisplayMode

            bNVidiaCard = False
            Dim sDevName As String = Manager.Adapters.Default.Information.Description.ToUpper
            If sDevName.Contains("NVIDIA") = True OrElse sDevName.Contains("GEFORCE") = True OrElse sDevName.Contains("NFORCE") = True Then
                bNVidiaCard = True
            End If

            uParms = New PresentParameters()
            With uParms
                .Windowed = muSettings.Windowed
                .SwapEffect = SwapEffect.Discard

                If muSettings.TripleBuffer = True Then .BackBufferCount = 2 Else .BackBufferCount = 1

                '.DeviceWindow = mfrmMain

                If muSettings.Windowed = True Then
                    .BackBufferFormat = uDispMode.Format
                    .BackBufferHeight = frmForm.ClientSize.Height
                    .BackBufferWidth = frmForm.ClientSize.Width
                Else
                    .BackBufferFormat = uDispMode.Format ' Format.A8R8G8B8
                    If muSettings.FullScreenResX < 1 Then
                        muSettings.FullScreenResX = uDispMode.Width
                    End If
                    If muSettings.FullScreenResY < 1 Then
                        muSettings.FullScreenResY = uDispMode.Height
                    End If
                    If muSettings.FullScreenRefreshRate < 1 Then
                        muSettings.FullScreenRefreshRate = uDispMode.RefreshRate
                    End If

                    Dim bValid As Boolean = False
                    For Each oTmpMode As DisplayMode In Manager.Adapters.Default.SupportedDisplayModes(.BackBufferFormat)
                        If oTmpMode.Width = muSettings.FullScreenResX AndAlso oTmpMode.Height = muSettings.FullScreenResY AndAlso oTmpMode.RefreshRate = muSettings.FullScreenRefreshRate Then
                            bValid = True
                            Exit For
                        End If
                    Next
                    If bValid = False Then
                        muSettings.FullScreenResX = uDispMode.Width
                        muSettings.FullScreenResY = uDispMode.Height
                        muSettings.FullScreenRefreshRate = uDispMode.RefreshRate
                    End If
                    'frmMain.IgnoreResizeEvents = True
                    'mfrmMain.Width = muSettings.FullScreenResX
                    'mfrmMain.Height = muSettings.FullScreenResY
                    'frmMain.IgnoreResizeEvents = False

                    .BackBufferWidth = muSettings.FullScreenResX
                    .BackBufferHeight = muSettings.FullScreenResY
                    .FullScreenRefreshRateInHz = muSettings.FullScreenRefreshRate
                End If

                Dim uDevCaps As Caps = Manager.GetDeviceCaps(Manager.Adapters.Default.Adapter, DeviceType.Hardware)

                Dim lPrsntInt As Int32 = PresentInterval.Default
                If muSettings.Windowed = True Then
                    If (uDevCaps.PresentationIntervals And PresentInterval.Immediate) = PresentInterval.Immediate Then
                        lPrsntInt = PresentInterval.Immediate
                    End If
                Else
                    If (uDevCaps.PresentationIntervals And PresentInterval.Immediate) = PresentInterval.Immediate Then
                        lPrsntInt = PresentInterval.Immediate
                    ElseIf (uDevCaps.PresentationIntervals And PresentInterval.One) = PresentInterval.One Then
                        lPrsntInt = PresentInterval.One
                    End If

                    If muSettings.VSync = True Then
                        lPrsntInt = PresentInterval.Default
                    End If
                End If
                'lPrsntInt = CInt(Val(oINI.GetString("GRAPHICS", "PresentInterval", lPrsntInt.ToString)))

                .PresentationInterval = CType(lPrsntInt, PresentInterval)
                '.PresentFlag = PresentFlag.DiscardDepthStencil

                Dim bD32 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D32, 0)
                Dim bD24X8 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24X8, 0)
                Dim bD24S8 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24S8, 0)
                Dim bD24X4S4 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24X4S4, 0)
                Dim bD16 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D16, 0)
                Dim bD15S1 As Boolean = Manager.CheckDepthStencilMatch(Manager.Adapters.Default.Adapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D15S1, 0)

                Dim lDpthFmt As Int32
                If muSettings.DepthBuffer = 0 Then
                    If bD32 = True Then
                        lDpthFmt = DepthFormat.D32
                    ElseIf bD24X8 = True Then
                        lDpthFmt = DepthFormat.D24X8
                    ElseIf bD24S8 = True Then
                        lDpthFmt = DepthFormat.D24S8
                    ElseIf bD16 = True Then
                        lDpthFmt = DepthFormat.D16
                    Else
                        MsgBox("Unable to determine Depth/Stencil Buffer format." & vbCrLf & "Please contact support for assistance!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Error")
                        End
                    End If
                    muSettings.DepthBuffer = lDpthFmt
                End If
                'lDpthFmt = CInt(Val(oINI.GetString("GRAPHICS", "DepthBuffer", lDpthFmt.ToString)))
                lDpthFmt = muSettings.DepthBuffer

                .AutoDepthStencilFormat = CType(lDpthFmt, DepthFormat)
                .EnableAutoDepthStencil = True
                .MultiSample = MultiSampleType.None
                .MultiSampleQuality = 0

            End With

        Catch
            MsgBox(Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Error in initialize Direct3D")
            bRes = False
        Finally
            uDispMode = Nothing
        End Try

        Return uParms
    End Function

#Region " Game Startup And Login Rendering "
    Private moVid As Microsoft.DirectX.AudioVideoPlayback.Video = Nothing
    Private mswVid As Stopwatch = Nothing
    Public mswDSELogoTimer As Stopwatch = Nothing
    Private mlStartupState As Int32 = 0
    Private Sub DrawDSELogoNoVid()

        'Regardless of anything, we always clear our buffer
        moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)

        If swStartup Is Nothing OrElse swStartup.IsRunning = False Then
            swStartup = Nothing
            swStartup = Stopwatch.StartNew()
            mlStartupDlgAlpha = 128
        End If

        If swStartup.ElapsedMilliseconds > 700 AndAlso mlStartupState > 1 Then
            If goSound Is Nothing = False Then
                If goSound.MusicOn = True AndAlso goSound.MusicStarted = False Then
                    'Start our music
                    goSound.StartMusic()
                End If
            End If
        End If


        If swStartup.ElapsedMilliseconds > 3000 Then ' OrElse mbBreakout = True Then   '6000
            'Skip all the intro hoohaa so the DEV can quick enter/exit the game.
            'If Debugger.IsAttached Then
            '    mlStartupState = 3
            '    mlStartupDlgAlpha = 0
            'End If

            '6 seconds *should* be enough for our thunder WAV to play...
            If mlStartupState = 0 Then      'decline of the DSE Logo
                mlStartupDlgAlpha -= 10
                If mlStartupDlgAlpha < 10 Then
                    mlStartupState += 1
                End If
            ElseIf mlStartupState = 1 Then  'incline of the Alienware logo
                mlStartupDlgAlpha += 10
                If mlStartupDlgAlpha > 255 Then
                    mlStartupDlgAlpha = 255
                    mlStartupState += 1
                    'Reset our timer so that the timer goes another 3 seconds
                    swStartup.Reset()
                    swStartup.Start()
                End If
            ElseIf mlStartupState = 2 Then  'decline of the alienware logo
                mlStartupDlgAlpha -= 10
                If mlStartupDlgAlpha < 10 Then
                    mlStartupState = 3
                    mlStartupDlgAlpha = 0
                End If
            ElseIf mlStartupState = 3 Then  '
                'Start up our logo, if it isn't
                If moLogo Is Nothing Then moLogo = New BPLogo(moDevice, 19)

                If mlStartupDlgAlpha <= 0 Then
                    mlStartupDlgAlpha = 0


                    'ok, tell resource manager to kill this texture
                    goResMgr.DeleteTexture("DSELogo.bmp")

                    'Ok, we also go ahead and cache our next image
                    muSettings.bRanBefore = True

                    'Now, set our new render view
                    glCurrentEnvirView = CurrentView.eStartupLogin


                    swStartup.Reset()

                    Exit Sub
                End If
            End If




        Else
            mlStartupDlgAlpha += 10
            If mlStartupDlgAlpha > 255 Then mlStartupDlgAlpha = 255

            If mlStartupDlgAlpha = 198 Then
                If goSound Is Nothing = False Then
                    goSound.StartSound("DSEThunder.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Vector3.Empty, Vector3.Empty)
                End If
            End If

        End If

        'ok, create a quick sprite
        Device.IsUsingEventHandlers = False
        Dim oTmpSprite As Sprite = New Sprite(moDevice)
        Device.IsUsingEventHandlers = True

        Dim lTempW As Int32 = moDevice.PresentationParameters.BackBufferWidth
        Dim lTempH As Int32 = moDevice.PresentationParameters.BackBufferHeight

        'ok, now, get the scaling factor
        Dim oTex1 As Texture = goResMgr.GetTexture("DSELogo.bmp", GFXResourceManager.eGetTextureType.StartupTexture, "Misc.pak")
        Dim oTex2 As Texture = goResMgr.GetTexture("AlienwarePBO.bmp", GFXResourceManager.eGetTextureType.StartupTexture, "apbo.pak")

        If mlStartupDlgAlpha < 0 Then mlStartupDlgAlpha = 0
        'Now, do it...
        oTmpSprite.Begin(SpriteFlags.None)
        If mlStartupState = 0 AndAlso mlStartupDlgAlpha <> 0 Then
            Dim lDestW As Int32 = Math.Min(lTempW, 1024)
            Dim lDestH As Int32 = 1024
            If lDestW <> 1024 Then lDestH = CInt((lDestW / 1024.0F) * 1024)
            Dim lNewH As Int32 = Math.Min(lDestH, lTempH)
            If lNewH <> lDestH Then lDestW = CInt((lNewH / lDestH) * lDestW)
            lDestH = lNewH

            'Now, position the image in the center...
            Dim fPtX As Single = (lTempW / 2.0F) - (lDestW / 2.0F)
            Dim fPtY As Single = (lTempH / 2.0F) - (lDestH / 2.0F)

            'Dim ptPos As Point = New Point(CInt((lTempW - 819) / 2), CInt((lTempH - 1024) / 2))
            Dim ptPos As Point = New Point(CInt(fPtX), CInt(fPtY))
            Dim fMultX As Single = 1024.0F / lDestW '819.0F / CSng(lTempW - ptPos.X)
            Dim fMultY As Single = 1024.0F / lDestH '1024.0F / CSng(lTempH - ptPos.Y)

            'ok, the hope is that whatever the screen resolution might be... we put our logo in the middle...
            'ptPos.X = CInt((lTempW) / 2)
            'ptPos.Y = CInt((lTempH) / 2)
            If oTex1 Is Nothing Then Return

            oTmpSprite.Draw2D(oTex1, _
                      New Rectangle(0, 0, 1024, 1024), System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lDestW, ptPos.Y + lDestH), _
                      System.Drawing.Point.Empty, 0, New Point(CInt(ptPos.X * fMultX), CInt(ptPos.Y * fMultY)), _
                      System.Drawing.Color.FromArgb(mlStartupDlgAlpha, 255, 255, 255))
        End If
        If mlStartupState > 0 AndAlso oTex2 Is Nothing = False Then
            Dim lDestW As Int32 = 1024
            Dim lDestH As Int32 = 512

            'Now, position the image in the center...
            Dim fPtX As Single = (lTempW / 2.0F) - (lDestW / 2.0F)
            Dim fPtY As Single = (lTempH / 2.0F) - (lDestH / 2.0F)

            'Dim ptPos As Point = New Point(CInt((lTempW - 819) / 2), CInt((lTempH - 1024) / 2))
            Dim ptPos As Point = New Point(CInt(fPtX), CInt(fPtY))
            Dim fMultX As Single = 1024.0F / lDestW '819.0F / CSng(lTempW - ptPos.X)
            Dim fMultY As Single = 512.0F / lDestH '1024.0F / CSng(lTempH - ptPos.Y)

            oTmpSprite.Draw2D(oTex2, _
                      New Rectangle(0, 0, 1024, 512), System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lDestW, ptPos.Y + lDestH), _
                      System.Drawing.Point.Empty, 0, New Point(CInt(ptPos.X * fMultX), CInt(ptPos.Y * fMultY)), _
                      System.Drawing.Color.FromArgb(mlStartupDlgAlpha, 255, 255, 255))
        End If
        oTmpSprite.End()
        oTmpSprite.Dispose()
        oTmpSprite = Nothing


    End Sub

    Private Sub DrawDSELogo()

        'Regardless of anything, we always clear our buffer
        moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)

        If swStartup Is Nothing OrElse swStartup.IsRunning = False Then
            swStartup = Nothing
            swStartup = Stopwatch.StartNew()
            mlStartupDlgAlpha = 128
        End If

        'If mswDSELogoTimer Is Nothing Then mswDSELogoTimer = Stopwatch.StartNew()
        'If mswDSELogoTimer.ElapsedMilliseconds > 11000 OrElse (mbBreakout = True AndAlso mswDSELogoTimer.ElapsedMilliseconds > 6000) Then
        '    mlStartupDlgAlpha = 0

        '    If moVid Is Nothing = False Then
        '        moVid.Stop()
        '        moVid.Dispose()
        '    End If
        '    moVid = Nothing

        '    'ok, tell resource manager to kill this texture
        '    goResMgr.DeleteTexture("DSELogo.bmp")

        '    'Ok, we also go ahead and cache our next image
        '    muSettings.bRanBefore = True

        '    'Now, set our new render view
        '    glCurrentEnvirView = CurrentView.eStartupLogin

        '    'goUILib.SetToolTip("Press a key to login", -1, -1)
        '    'goUILib.RetainTooltip = True
        '    mswDSELogoTimer.Reset()
        '    mswDSELogoTimer = Nothing

        '    swStartup.Reset()
        '    Return
        'End If

        If swStartup.ElapsedMilliseconds > 3000 Then ' OrElse mbBreakout = True Then   '6000
            mlStartupDlgAlpha -= 10
            If moVid Is Nothing = False AndAlso swStartup.ElapsedMilliseconds < 8000 Then

                If moVid.SeekingCaps.CanGetCurrentPosition = False OrElse moVid.SeekingCaps.CanGetDuration = False Then
                    If mswVid Is Nothing AndAlso swStartup.ElapsedMilliseconds < 5000 Then
                        mswVid = Stopwatch.StartNew
                        moVid.Play()
                        Return
                    ElseIf mswVid Is Nothing = False AndAlso mswVid.ElapsedMilliseconds < 4000 Then
                        'Start the music if it isn't
                        If swStartup.ElapsedMilliseconds > 4800 Then        '3500
                            If goSound Is Nothing = False Then
                                If goSound.MusicOn = True AndAlso goSound.MusicStarted = False Then
                                    'Start our music
                                    goSound.StartMusic()
                                End If
                            End If
                        End If

                        Return
                    End If
                    If mswVid Is Nothing = False Then mswVid.Stop()
                    mswVid = Nothing
                Else
                    If moVid.CurrentPosition = 0 Then
                        moVid.Play()
                    ElseIf moVid.Playing = True AndAlso moVid.Paused = False AndAlso moVid.CurrentPosition + 0.0001F < moVid.Duration Then
                        'Start the music if it isn't
                        If swStartup.ElapsedMilliseconds > 4800 Then        '3500
                            If goSound Is Nothing = False Then
                                If goSound.MusicOn = True AndAlso goSound.MusicStarted = False Then
                                    'Start our music
                                    goSound.StartMusic()
                                End If
                            End If
                        End If

                        Return
                    End If
                End If
            Else

                'Start the music if it isn't
                If goSound Is Nothing = False Then
                    If goSound.MusicOn = True AndAlso goSound.MusicStarted = False Then
                        'Start our music
                        goSound.StartMusic()
                    End If
                End If

                mlStartupDlgAlpha = 0

                If moVid Is Nothing = False Then
                    moVid.Stop()
                    moVid.Dispose()
                End If
                moVid = Nothing

                'ok, tell resource manager to kill this texture
                goResMgr.DeleteTexture("DSELogo.bmp")

                'Ok, we also go ahead and cache our next image
                muSettings.bRanBefore = True

                'Now, set our new render view
                glCurrentEnvirView = CurrentView.eStartupLogin

                'goUILib.SetToolTip("Press a key to login", -1, -1)
                'goUILib.RetainTooltip = True
                If mswDSELogoTimer Is Nothing = False Then mswDSELogoTimer.Reset()
                mswDSELogoTimer = Nothing

                swStartup.Reset()
                Return
            End If


        Else
            mlStartupDlgAlpha += 10
            If mlStartupDlgAlpha > 255 Then mlStartupDlgAlpha = 255

            If mlStartupDlgAlpha = 198 Then
                If goSound Is Nothing = False Then
                    goSound.StartSound("DSEThunder.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Vector3.Empty, Vector3.Empty)
                End If
            End If

            'This doesn't get hit normally, if the delay above for the parent if/else is < 3500
            If moVid Is Nothing AndAlso swStartup.ElapsedMilliseconds > 500 Then        '3500
                If mlStartupDlgAlpha = 255 Then
                    'Start the music if it isn't
                    If goSound Is Nothing = False Then
                        If goSound.MusicOn = True AndAlso goSound.MusicStarted = False Then
                            'Start our music
                            goSound.StartMusic()
                        End If
                    End If
                End If
            End If
        End If

        'ok, create a quick sprite
        Device.IsUsingEventHandlers = False
        Dim oTmpSprite As Sprite = New Sprite(moDevice)
        Device.IsUsingEventHandlers = True

        Dim lTempW As Int32 = moDevice.PresentationParameters.BackBufferWidth
        Dim lTempH As Int32 = moDevice.PresentationParameters.BackBufferHeight

        'ok, now, get the scaling factor
        Dim oTex1 As Texture = goResMgr.GetTexture("DSELogo.bmp", GFXResourceManager.eGetTextureType.StartupTexture, "Misc.pak")
        Dim oTex2 As Texture = goResMgr.GetTexture("AlienwarePBO.bmp", GFXResourceManager.eGetTextureType.StartupTexture, "apbo.pak")

        If mlStartupDlgAlpha < 0 Then mlStartupDlgAlpha = 0
        'Now, do it...
        oTmpSprite.Begin(SpriteFlags.None)

        Dim lDestW As Int32 = Math.Min(lTempW, 1024)
        Dim lDestH As Int32 = 1024
        If lDestW <> 1024 Then lDestH = CInt((lDestW / 1024.0F) * 1024)
        Dim lNewH As Int32 = Math.Min(lDestH, lTempH)
        If lNewH <> lDestH Then lDestW = CInt((lNewH / lDestH) * lDestW)
        lDestH = lNewH

        'Now, position the image in the center...
        Dim fPtX As Single = (lTempW / 2.0F) - (lDestW / 2.0F)
        Dim fPtY As Single = (lTempH / 2.0F) - (lDestH / 2.0F)

        'Dim ptPos As Point = New Point(CInt((lTempW - 819) / 2), CInt((lTempH - 1024) / 2))
        Dim ptPos As Point = New Point(CInt(fPtX), CInt(fPtY))
        Dim fMultX As Single = 1024.0F / lDestW '819.0F / CSng(lTempW - ptPos.X)
        Dim fMultY As Single = 1024.0F / lDestH '1024.0F / CSng(lTempH - ptPos.Y)

        'ok, the hope is that whatever the screen resolution might be... we put our logo in the middle...
        'ptPos.X = CInt((lTempW) / 2)
        'ptPos.Y = CInt((lTempH) / 2)
        If oTex1 Is Nothing Then Return

        oTmpSprite.Draw2D(oTex1, _
                  New Rectangle(0, 0, 1024, 1024), System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lDestW, ptPos.Y + lDestH), _
                  System.Drawing.Point.Empty, 0, New Point(CInt(ptPos.X * fMultX), CInt(ptPos.Y * fMultY)), _
                  System.Drawing.Color.FromArgb(mlStartupDlgAlpha, 255, 255, 255))
        oTmpSprite.End()
        oTmpSprite.Dispose()
        oTmpSprite = Nothing


    End Sub

    Private Sub DrawLoginBackground()

        Try
            'regardless, we clear to black
            moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)

            'Now, determine our fade
            If swStartup Is Nothing Then
                'Ok, we're being told to close this screen
                mlStartupDlgAlpha -= 10
                'If moLogoFX Is Nothing = False AndAlso moLogoFX.mbEmitterStopping = False Then moLogoFX.StopEmitter()
                If mlStartupDlgAlpha <= 0 Then
                    ForceCleanupLoginBackground()
                    Exit Sub
                End If
            Else
                mlStartupDlgAlpha += 10
                If mlStartupDlgAlpha > 255 Then mlStartupDlgAlpha = 255
                'ensure our Camera is correct
                With goCamera
                    Dim bUpdateSound As Boolean = .mlCameraX <> 0 OrElse .mlCameraZ <> -10000
                    .mlCameraAtX = 0 : .mlCameraAtY = 0 : .mlCameraAtZ = 0
                    .mlCameraX = 0 : .mlCameraY = 0 : .mlCameraZ = -10000
                    .CheckUpdateListenerLoc()
                End With

            End If

            If LoginScreen.mySequenceEnded = 1 OrElse LoginScreen.mySequenceEnded = 2 Then
                If goSound Is Nothing = False AndAlso goSound.MusicOn = True Then
                    Dim sName As String = goSound.GetCurrentIntroSong()

                    If sName.ToUpper.EndsWith("INTRO2.MP3") = True Then
                        If moLoginScreen Is Nothing = False AndAlso LoginScreen.mfFadeoutAlpha > 200 Then
                            moLoginScreen = Nothing
                            LoginScreen.mfFadeoutAlpha -= 1
                            If LoginScreen.mfFadeoutAlpha <= 200 AndAlso LoginScreen.mySequenceEnded = 1 Then
                                LoginScreen.mfFadeoutAlpha = 200
                                LoginScreen.mySequenceEnded = 2
                                moLoginScreen = New LoginScreen(moDevice)
                            End If
                        End If
                    ElseIf sName.ToUpper.EndsWith("INTRO1.MP3") = True Then
                        LoginScreen.mySequenceEnded = 3
                        LoginScreen.BSDestroyed = 0
                    End If
                Else
                    Static xblLastElapsed As Int64 = swStartup.ElapsedMilliseconds
                    If swStartup.IsRunning = False Then swStartup.Start()
                    If swStartup.ElapsedMilliseconds - xblLastElapsed > 5000 Then
                        LoginScreen.mySequenceEnded = 3
                        LoginScreen.BSDestroyed = 0
                        swStartup.Reset()
                    Else
                        If moLoginScreen Is Nothing = False AndAlso LoginScreen.mfFadeoutAlpha > 200 Then
                            moLoginScreen = Nothing
                            LoginScreen.mfFadeoutAlpha -= 1
                            If LoginScreen.mfFadeoutAlpha <= 200 AndAlso LoginScreen.mySequenceEnded = 1 Then
                                LoginScreen.mfFadeoutAlpha = 200
                                LoginScreen.mySequenceEnded = 2
                                moLoginScreen = New LoginScreen(moDevice)
                            End If
                        End If
                    End If
                End If
            Else

                If LoginScreen.mySequenceEnded = 3 Then
                    LoginScreen.mfFadeoutAlpha -= 10.0F
                    If LoginScreen.mfFadeoutAlpha < 0 Then
                        LoginScreen.mfFadeoutAlpha = 0
                        LoginScreen.mySequenceEnded = 0
                    End If
                End If

                If moLoginScreen Is Nothing Then moLoginScreen = New LoginScreen(moDevice)
                SetRenderStates()
                moLoginScreen.DoRender(moDevice)

                If muSettings.PostGlowAmt > 0 Then
                    If moPost Is Nothing Then moPost = New PostShader()
                    moPost.ExecutePostProcess()
                ElseIf moPost Is Nothing = False Then
                    moPost.DisposeMe()
                    moPost = Nothing
                End If
            End If
        Catch
        End Try

    End Sub

    Public Sub ForceCleanupLoginBackground()
        'Called to ensure that our memory is cleared 
        'If moLogoFX Is Nothing = False AndAlso moLogoFX.mbEmitterStopping = False Then moLogoFX.StopEmitter()
        moLoginScreen = Nothing
    End Sub

    Public Sub RenderOnlyLogo()
        If mbInitialized = False Then Exit Sub
        moDevice.BeginScene()
        moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)
        SetRenderStates()
        moDevice.RenderState.Ambient = System.Drawing.Color.DarkGray
        If moLogo Is Nothing = False Then
            If moLogo.Render(Not glCurrentEnvirView = CurrentView.eStartupLogin, moDevice) = True Then
                moLogo = Nothing
            End If
        End If
        goUILib.RenderInterfaces(glCurrentEnvirView)

        goCamera.SetupMatrices(moDevice, glCurrentEnvirView)
        moDevice.EndScene()
        'moDevice.Present()
        PresentTheScene()
    End Sub

    'Private Sub moLogoFX_Disposed() Handles moLogoFX.Disposed
    '    moLogoFX = Nothing
    'End Sub

    Public Sub InitializeLoginScreen()
        If moLoginScreen Is Nothing Then moLoginScreen = New LoginScreen(moDevice)

        If moVid Is Nothing = True Then

            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            sPath &= "alienware.mpg"
            If Exists(sPath) = True Then

                Dim oINI As New InitFile
                Dim bShowMovie As Boolean = CInt(oINI.GetString("SETTINGS", "ShowAlienWareVideo", "1")) <> 0
                bShowMovie = False
                If bShowMovie = False Then Return


                'if Microsoft.DirectX.AudioVideoPlayback.SeekingCaps
                Dim lTmpHt As Int32 = mfrmMain.Height
                Dim lTmpWd As Int32 = mfrmMain.Width
                Try
                    moVid = New Microsoft.DirectX.AudioVideoPlayback.Video(sPath) '"C:\Documents and Settings\Matthew Campbell\Desktop\alienware.mpg")
                    moVid.Owner = mfrmMain
                    mfrmMain.Width = lTmpWd
                    mfrmMain.Height = lTmpHt
                    moVid.CurrentPosition = 0
                Catch
                End Try
            End If
        End If
    End Sub

#End Region

    Private mlLostExceptionCount As Int32 = 0
    Private Function ProcessLostDevice() As Boolean
        'ok, see if we can continue again...
        Dim lResult As Int32
        'Check the cooperative level to see if it's ok to render
        If moDevice.CheckCooperativeLevel(lResult) = False Then
            If lResult = ResultCode.DeviceLost Then
                'the device has been lost but cannot be reset at this time, so wait until it can be reset
                Threading.Thread.Sleep(50)
                Return False
            End If
        End If

        'if we are windowed, check the format to ensure that we are in the same format (backbuffer to display mode)
        'If muSettings.Windowed = True Then

        'End If 

        Try
            'now, we force the device to reset...
            Dim uParms As PresentParameters = CreatePresentationParams(mfrmMain)
            If uParms Is Nothing Then Return False
            moDevice.Reset(uParms)
            gbDeviceLost = False
        Catch exDevLost As DeviceLostException
            Threading.Thread.Sleep(50)
            Return False
        Catch ex As Exception
            'reset failed but the device wasn't lost so something bad happened. 
            'we will just throw an error as we cannot handle the event well at this point

            'Special boolean used in testing a forced crash -> Recreate Everything
            If gbDeviceLostHard Then
                RaiseEvent CriticalFailure()
                gbDeviceLostHard = False
                Return False
            End If
            Try
                Dim uParms As PresentParameters = CreatePresentationParams(mfrmMain)
                If uParms Is Nothing Then Return False
                moDevice.Reset(uParms)
                gbDeviceLost = False
            Catch exx As Exception
                mlLostExceptionCount += 1
                If mlLostExceptionCount > 10 Then
                    'ok, bad... so...
                    RaiseEvent CriticalFailure()
                End If
                Return False
            End Try
        End Try

        Return True
    End Function

    ' Private moCityCars As CityCars.CityCars = Nothing
    Public sDrawSceneLocation As String = ""

    Public bRenderSelectedRanges As Boolean = False

    Public Sub ForceDeviceLost()
        moDevice_DeviceLost(Nothing, Nothing)
    End Sub
    Public Function DrawScene() As Int32
        Dim X As Int32
        Dim yVisible As Byte
        Dim lRenderCnt As Int32 = 0
        Dim lCameraCellX As Int32
        Dim lCameraCellZ As Int32
        Dim yRel As Byte


        'If glCurrentEnvirView = CurrentView.ePlanetView Then
        '    If moCityCars Is Nothing Then
        '        moCityCars = New CityCars.CityCars(moDevice, 16)
        '        For X = 0 To goCurrentEnvir.lEntityUB
        '            If goCurrentEnvir.lEntityIdx(X) > -1 AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
        '                For Y As Int32 = 0 To goCurrentEnvir.lEntityUB
        '                    If X <> Y AndAlso goCurrentEnvir.lEntityIdx(Y) > -1 AndAlso goCurrentEnvir.oEntity(Y).ObjTypeID = ObjectType.eFacility AndAlso goCurrentEnvir.oEntity(X).OwnerID = goCurrentEnvir.oEntity(Y).OwnerID Then
        '                        moCityCars.AddParticle(goCurrentEnvir.oEntity(X), goCurrentEnvir.oEntity(Y))
        '                    End If
        '                Next Y
        '            End If
        '        Next X
        '    End If
        'End If
        ExplosionManager.CurrentAddCount = 0

        sDrawSceneLocation = "Start"

        'first, check if we initialized ourselves...
		If mbInitialized = False Then Return 0

        If gbPaused = True Then Return 0
        If gbDeviceLost = True Then
            If ProcessLostDevice() = False Then Return 0
        End If

        If mbToggleFullScreen = True Then
            DoToggleFullScreen()
            Return 0
        End If

        Dim lCoopRes As Int32 = 0
        If moDevice.CheckCooperativeLevel(lCoopRes) = False Then
            If lCoopRes = ResultCode.DeviceNotReset Then
                gbDeviceLost = True
                gbPaused = True
            End If
            Return 0
        End If

        sDrawSceneLocation = "BeginScene"
        'Try
        '    moDevice.BeginScene()
        'Catch ex As Exception
        '    moDevice.EndScene()
        '    moDevice.BeginScene()
        'End Try
        moDevice.BeginScene()



        sDrawSceneLocation = "SetRenderStates"
        SetRenderStates()
        moDevice.RenderState.Ambient = System.Drawing.Color.DarkGray

        'Now, check if we are in special views... the 2 start ups to be exact...

        If glCurrentEnvirView = CurrentView.eStartupDSELogo Then
            If moVid Is Nothing = False Then DrawDSELogo() Else DrawDSELogoNoVid()
        ElseIf glCurrentEnvirView = CurrentView.eStartupLogin Then
            sDrawSceneLocation = "DrawLoginBackground"
            DrawLoginBackground()

            If moLogo Is Nothing = False Then
                If moLogo.Render(Not glCurrentEnvirView = CurrentView.eStartupLogin, moDevice) = True Then
                    moLogo = Nothing
                End If
            End If

            'Now, draw our interface objects
            sDrawSceneLocation = "Login.RenderInterfaces"
            goUILib.RenderInterfaces(glCurrentEnvirView)
        ElseIf glCurrentEnvirView > CurrentView.eFullScreenInterface Then
            'check our interface backdrop...
            sDrawSceneLocation = "FS.Clear"
            moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)

            sDrawSceneLocation = "FS.Background"
            If goFullScreenBackground Is Nothing = False AndAlso goFullScreenBackground.Disposed = False Then
                'ok, create a quick sprite
                'Device.IsUsingEventHandlers = False
                'Dim oTmpSprite As Sprite = New Sprite(moDevice)
                'Device.IsUsingEventHandlers = True

                'Determine our parameters
                Dim lTempW As Int32 = moDevice.PresentationParameters.BackBufferWidth
                Dim lTempH As Int32 = moDevice.PresentationParameters.BackBufferHeight
                Dim ptPos As Point '= New Point(CInt((lTempW - 800) / 2), CInt((lTempH - 600) / 2))

                'ok, the hope is that whatever the screen resolution might be... we put our logo in the middle...
                ptPos.X = 0 'CInt((lTempW - 800) / 2)
                ptPos.Y = 0 'CInt((lTempH - 600) / 2)

                'Now, do it...
                Dim bFog As Boolean = False
                With moDevice.RenderState
                    .Lighting = False
                    bFog = .FogEnable
                    .FogEnable = False
                End With
                Try
                    Dim oSurfDesc As SurfaceDescription = goFullScreenBackground.GetLevelDescription(0)
                    BPSprite.Draw2DOnce(moDevice, goFullScreenBackground, New Rectangle(0, 0, oSurfDesc.Width, oSurfDesc.Height), _
                      System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lTempW, ptPos.Y + lTempH), System.Drawing.Color.FromArgb(255, 255, 255, 255), oSurfDesc.Width, oSurfDesc.Height)
                Catch
                End Try
                With moDevice.RenderState
                    .Lighting = True
                    .FogEnable = bFog
                End With

                'oTmpSprite.Begin(SpriteFlags.None)
                'Try
                '    oTmpSprite.Draw2D(goFullScreenBackground, Rectangle.Empty, _
                '      System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lTempW, ptPos.Y + lTempH), _
                '      System.Drawing.Point.Empty, 0, New Point(CInt(ptPos.X), CInt(ptPos.Y)), System.Drawing.Color.FromArgb(255, 255, 255, 255))
                'Catch
                'End Try
                'oTmpSprite.End()
                'oTmpSprite.Dispose()
                'oTmpSprite = Nothing
            End If

            'Now, draw our interface objects
            sDrawSceneLocation = "FS.RenderInterfaces"
            goUILib.RenderInterfaces(glCurrentEnvirView)
        Else
            sDrawSceneLocation = "CalculateFrustrum"
            'Do our calculate frustrum now...
            goCamera.CalculateFrustrum(moDevice)

            'moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, Color.Black, 1, 0)

            'Render like normal...
            sDrawSceneLocation = "GeoRender"
            If goGalaxy Is Nothing = False Then
                If gbMonitorPerformance = True Then gsw_Geography.Start()
                goGalaxy.Render(glCurrentEnvirView)
                If gbMonitorPerformance = True Then gsw_Geography.Stop()
            End If

            If (glCurrentEnvirView = CurrentView.eSystemView AndAlso muSettings.DrawGrid = True) Then ' Or CurrentEnvirView = CurrentView.eSystemMapView2 Then
                'Draw our grid
                goCamera.DrawGrid(moDevice)
            End If

            'before we render any units, re-enable lighting, in case it was set to off
            moDevice.SetRenderState(RenderStates.Lighting, 1)

            If glCurrentEnvirView = CurrentView.ePlanetView OrElse glCurrentEnvirView = CurrentView.eSystemView Then

                Dim lFar As Int32
                Dim lTexMod As Int32

                lFar = muSettings.EntityClipPlane

                lCameraCellX = goCamera.mlCameraAtX \ lFar ' CInt(Math.Floor(goCamera.mlCameraX / lFar))
                lCameraCellZ = goCamera.mlCameraAtZ \ lFar ' CInt(Math.Floor(goCamera.mlCameraZ / lFar))

                'If moRadarMesh Is Nothing Then
                '    Device.IsUsingEventHandlers = False
                '    moRadarMesh = Mesh.Sphere(moDevice, 20, 4, 2)
                '    Device.IsUsingEventHandlers = True
                'End If

                mlRadarAlpha += mlRadarAlphaMod
                If mlRadarAlpha < 64 Then
                    mlRadarAlpha = 64
                    mlRadarAlphaMod = 1
                ElseIf mlRadarAlpha > 200 Then
                    mlRadarAlpha = 200
                    mlRadarAlphaMod = -1
                End If
                With matRadar
                    .Ambient = System.Drawing.Color.FromArgb(mlRadarAlpha, 0, 255, 0)
                    .Diffuse = System.Drawing.Color.FromArgb(mlRadarAlpha, 0, 255, 0)
                    .Emissive = System.Drawing.Color.FromArgb(mlRadarAlpha, 0, 255, 0)
                End With

                If muSettings.RenderBurnMarks = True Then
                    sDrawSceneLocation = "BurnMark.BeginFrame"
                    goBurnMarkMgr.BeginFrame()
                End If

                'if I am in either of these views then I should be able to see the units...
                moDevice.RenderState.CullMode = Cull.None
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing = False Then
                    If (oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView = CurrentView.ePlanetView) OrElse _
                       (oEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso glCurrentEnvirView = CurrentView.eSystemView) Then

                        'Dim yMapWrapSituation As Byte = 0           '0 = no map wrap, 1 = left edge, 2 = right edge
                        'Dim lLocXMapWrapCheck As Int32 = 0

                        'If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                        '    Dim lTmpMapWrapVal As Int32 = Math.Min((oEnvir.lMaxXPos * 2) \ 3, muSettings.EntityClipPlane)
                        '    If goCamera.mlCameraX < oEnvir.lMinXPos + lTmpMapWrapVal Then
                        '        yMapWrapSituation = 1
                        '        lLocXMapWrapCheck = oEnvir.lMaxXPos - lTmpMapWrapVal
                        '    ElseIf goCamera.mlCameraX > oEnvir.lMaxXPos - lTmpMapWrapVal Then
                        '        yMapWrapSituation = 2
                        '        lLocXMapWrapCheck = oEnvir.lMinXPos + lTmpMapWrapVal
                        '    End If
                        'End If

                        If gbMonitorPerformance = True Then gsw_Models.Start()
                        If mbSupportsNewModelMethod = True Then SetModelTextureStates()
                        Dim alSelectedIdx As New ArrayList

                        If HasAliasedRights(AliasingRights.eViewUnitsAndFacilities) = True Then

                            If goSound Is Nothing = False Then
                                sDrawSceneLocation = "Sound.PrepareUnitSoundCollector"
                                goSound.PrepareUnitSoundCollector()
                            End If


                            If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                                sDrawSceneLocation = "ModelShader.PrepareToRender"
                                If moModelShader Is Nothing Then moModelShader = New ModelShader()

                                Dim vecOverallDiff As Vector3 = New Vector3(0, 0, 0)
                                Dim vecOverallSpec As Vector3 = New Vector3(0, 0, 0)
                                Dim vecOverallAmb As Vector3 = New Vector3(0, 0, 0)

                                Try
                                    For lLightIdx As Int32 = 0 To moDevice.Lights.Count - 1
                                        If moDevice.Lights(lLightIdx).Enabled = True Then
                                            With moDevice.Lights(lLightIdx)
                                                vecOverallDiff.X = Math.Max(vecOverallDiff.X, .Diffuse.R / 255.0F)
                                                vecOverallDiff.Y = Math.Max(vecOverallDiff.Y, .Diffuse.G / 255.0F)
                                                vecOverallDiff.Z = Math.Max(vecOverallDiff.Z, .Diffuse.B / 255.0F)

                                                vecOverallSpec.X = Math.Max(vecOverallSpec.X, .Specular.R / 255.0F)
                                                vecOverallSpec.Y = Math.Max(vecOverallSpec.Y, .Specular.G / 255.0F)
                                                vecOverallSpec.Z = Math.Max(vecOverallSpec.Z, .Specular.B / 255.0F)

                                                vecOverallAmb.X *= (.Ambient.R / 255.0F)
                                                vecOverallAmb.Y *= (.Ambient.G / 255.0F)
                                                vecOverallAmb.Z *= (.Ambient.B / 255.0F)
                                            End With
                                        End If
                                    Next lLightIdx
                                Catch
                                End Try

                                Dim clrDiff As System.Drawing.Color = System.Drawing.Color.FromArgb(255, Math.Min(255, Math.Max(0, CInt(vecOverallDiff.X * 255.0F))), Math.Min(255, Math.Max(0, CInt(vecOverallDiff.Y * 255.0F))), Math.Min(255, Math.Max(0, CInt(vecOverallDiff.Z * 255.0F))))
                                Dim clrSpec As System.Drawing.Color = System.Drawing.Color.FromArgb(255, Math.Min(255, Math.Max(0, CInt(vecOverallSpec.X * 255.0F))), Math.Min(255, Math.Max(0, CInt(vecOverallSpec.Y * 255.0F))), Math.Min(255, Math.Max(0, CInt(vecOverallSpec.Z * 255.0F))))
                                Dim clrAmb As System.Drawing.Color = System.Drawing.Color.FromArgb(255, Math.Min(255, Math.Max(0, CInt(vecOverallAmb.X * 255.0F))), Math.Min(255, Math.Max(0, CInt(vecOverallAmb.Y * 255.0F))), Math.Min(255, Math.Max(0, CInt(vecOverallAmb.Z * 255.0F))))


                                If glCurrentEnvirView = CurrentView.ePlanetView Then
                                    moModelShader.PrepareToRender(moDevice.Lights(0).Direction, clrDiff, clrAmb, clrSpec)
                                    'moModelShader.PrepareToRender(moDevice.Lights(0).Direction, moDevice.Lights(0).Diffuse, moDevice.Lights(0).Ambient, moDevice.Lights(0).Specular)
                                Else
                                    Dim vecTemp As Vector3 = New Vector3(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ)
                                    'moModelShader.PrepareToRender(vecTemp, moDevice.Lights(0).Diffuse, moDevice.Lights(0).Ambient, moDevice.Lights(0).Specular)
                                    moModelShader.PrepareToRender(vecTemp, clrDiff, clrAmb, clrSpec)
                                End If
                            End If

                            Dim bHasIronCurtains As Boolean = False
                            Dim bHasLongRangeRadar As Boolean = False
                            Dim bCheckGuild As Boolean = False
                            Dim oGuild As Guild = Nothing
                            If goCurrentPlayer Is Nothing = False Then
                                oGuild = goCurrentPlayer.oGuild
                                If oGuild Is Nothing = False Then
                                    bCheckGuild = (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0
                                End If
                            End If
                            UpdateCP()
                            'frmEnvirDisplay.lMyFacilities = 0
                            Dim alTethers As New ArrayList

                            Dim bSaveTacticalData As Boolean = False
                            For X = 0 To oEnvir.lEntityUB
                                If oEnvir.lEntityIdx(X) <> -1 Then
                                    sDrawSceneLocation = "GetEntityFromArray"
                                    Dim oTmpEntity As BaseEntity = oEnvir.oEntity(X)
                                    If oTmpEntity Is Nothing OrElse oTmpEntity.oMesh Is Nothing Then Continue For

                                    With oTmpEntity
                                        sDrawSceneLocation = "RenderEntity"
                                        'If .ObjTypeID = ObjectType.eFacility AndAlso .OwnerID = glPlayerID Then
                                        '    If .yProductionType = ProductionType.eSpaceStationSpecial Then frmEnvirDisplay.lMyFacilities += 50 Else frmEnvirDisplay.lMyFacilities += 1
                                        'End If

                                        .bCulled = True
                                        Dim vecTemp As Vector3 = .oMesh.vecHalfDeathSeqSize

                                        If .ObjTypeID = ObjectType.eUnit AndAlso .bSelected = True AndAlso .OwnerID = glPlayerID AndAlso .lTetherPointX <> Int32.MinValue AndAlso .lTetherPointZ <> Int32.MinValue Then
                                            sDrawSceneLocation = "TetherPtAdd"
                                            Dim ptLoc As New Point(.lTetherPointX, .lTetherPointZ)
                                            alTethers.Add(ptLoc)
                                        End If

                                        If bRenderSelectedRanges = True AndAlso .bSelected = True AndAlso .OwnerID = glPlayerID AndAlso (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                                            sDrawSceneLocation = "SelectedIdxAdd"
                                            alSelectedIdx.Add(X)
                                        End If

                                        'Dim fTmpX As Single = .LocX
                                        'If yMapWrapSituation <> 0 AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet Then
                                        '    'Ok, in a mapwrap situation
                                        '    If yMapWrapSituation = 1 Then
                                        '        'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
                                        '        If .LocX > lLocXMapWrapCheck Then
                                        '            fTmpX -= oEnvir.lMapWrapAdjustX
                                        '        End If
                                        '    Else
                                        '        'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
                                        '        If .LocX < lLocXMapWrapCheck Then
                                        '            fTmpX += oEnvir.lMapWrapAdjustX
                                        '        End If
                                        '    End If
                                        'End If
                                        '.fMapWrapLocX = fTmpX

                                        If goSound Is Nothing = False AndAlso .ObjTypeID = ObjectType.eUnit AndAlso .oMesh.sRoarSFX <> "" AndAlso (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso .yVisibility = eVisibilityType.Visible Then goSound.TestEntitySound(X, .LocX, .LocY, .LocZ)

                                        'If goCamera.CullObject(CullBox.GetCullBox(fTmpX, .LocY, .LocZ, -vecTemp.X, -vecTemp.Y, -vecTemp.Z, vecTemp.X, vecTemp.Y, vecTemp.Z)) = False Then
                                        If goCamera.CullObject(CullBox.GetCullBox(.LocX, .LocY, .LocZ, -vecTemp.X, -vecTemp.Y, -vecTemp.Z, vecTemp.X, vecTemp.Y, vecTemp.Z)) = False Then
                                            .bCulled = False
                                            lRenderCnt += 1

                                            If muSettings.RenderBurnMarks = True Then
                                                sDrawSceneLocation = "BurnMark.RenderFromObj"
                                                goBurnMarkMgr.RenderFromObj(oTmpEntity)
                                            End If

                                            'The valid entries for lTexMod are:
                                            ' 0 = Neutral
                                            ' 1 = Mine (belongs to player)
                                            ' 2 = Ally
                                            ' 3 = Enemy
                                            If .OwnerID = glPlayerID Then
                                                lTexMod = 1
                                                yVisible = eVisibilityType.Visible
                                            ElseIf (bCheckGuild = True AndAlso oGuild.MemberInGuild(.OwnerID) = True) Then
                                                lTexMod = 4
                                                yVisible = eVisibilityType.Visible
                                            Else
                                                yRel = goCurrentPlayer.GetPlayerRelScore(.OwnerID)

                                                If yRel <= elRelTypes.eWar Then
                                                    lTexMod = 3
                                                ElseIf yRel <= elRelTypes.ePeace Then       'elRelTypes.eNeutral
                                                    lTexMod = 0
                                                Else : lTexMod = 2
                                                End If
                                                yVisible = oEnvir.UnitInRadarRange(CInt(.LocX), CInt(.LocZ), .oMesh.RangeOffset)
                                            End If
                                            'MSC: we only care about the mesh portion of the modelid
                                            If (.oMesh.lModelID And 255) = 24 OrElse (.yProductionType = ProductionType.eMining AndAlso .ObjTypeID = ObjectType.eFacility) Then
                                                yVisible = 2 ' eVisibilityType.Visible
                                            End If
                                            If .yVisibility = eVisibilityType.Visible Then
                                                If yVisible <> eVisibilityType.Visible AndAlso .ObjTypeID = ObjectType.eFacility Then
                                                    yVisible = eVisibilityType.FacilityIntel
                                                End If
                                            ElseIf .yVisibility = eVisibilityType.FacilityIntel Then
                                                If yVisible = eVisibilityType.Visible Then
                                                    If .bObjectDestroyed = True Then

                                                        If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.oGeoObject Is Nothing = False AndAlso CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                                                            CType(oEnvir.oGeoObject, Planet).SetCityCreepLoc(.LocX, .LocZ, False, .oMesh.ShieldXZRadius)
                                                        End If

                                                        .ClearParticleFX()
                                                        oEnvir.lEntityIdx(X) = -1
                                                        oEnvir.oEntity(X) = Nothing

                                                        bSaveTacticalData = True
                                                        Continue For

                                                    End If
                                                Else : yVisible = eVisibilityType.FacilityIntel
                                                End If
                                            End If

                                            'TODO: For Full visibility, this is all you have to do!!!
                                            'yVisible = eVisibilityType.Visible

                                            If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso .ObjTypeID = ObjectType.eFacility Then
                                                If yVisible >= eVisibilityType.Visible AndAlso .yVisibility < eVisibilityType.Visible Then
                                                    'ok, let's ensure that our terrain vert is updated...
                                                    If oEnvir.oGeoObject Is Nothing = False AndAlso CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                                                        sDrawSceneLocation = "SetCityCreepLoc"
                                                        CType(oEnvir.oGeoObject, Planet).SetCityCreepLoc(.LocX, .LocZ, True, .oMesh.ShieldXZRadius)
                                                    End If
                                                End If
                                            End If

                                            .yVisibility = yVisible
                                            sDrawSceneLocation = "RenderModel/RenderMesh"
                                            If yVisible = eVisibilityType.Visible Then
                                                If oEnvir.IsPlayerIronCurtain(.OwnerID) = True Then
                                                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                                                        bHasIronCurtains = True
                                                    Else
                                                        If oTmpEntity.bSelected = True Then RenderModel(oTmpEntity, lTexMod, RenderModelType.eIronCurtain Or RenderModelType.eSelected) Else RenderModel(oTmpEntity, lTexMod, RenderModelType.eIronCurtain)
                                                    End If
                                                Else
                                                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                                                        moModelShader.RenderMesh(oTmpEntity, RenderModelType.eNoSpecial, lTexMod)
                                                    Else
                                                        If oTmpEntity.bSelected = True Then RenderModel(oTmpEntity, lTexMod, RenderModelType.eSelected) Else RenderModel(oTmpEntity, lTexMod, RenderModelType.eNoSpecial)
                                                    End If
                                                End If
                                            ElseIf yVisible = eVisibilityType.FacilityIntel Then
                                                If oEnvir.IsPlayerIronCurtain(.OwnerID) = True Then
                                                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                                                        bHasIronCurtains = True
                                                    Else
                                                        RenderModel(oTmpEntity, lTexMod, RenderModelType.eIronCurtain Or RenderModelType.eOldData)
                                                    End If
                                                Else
                                                    If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                                                        oTmpEntity.GetWorldMatrix()
                                                        moModelShader.RenderMesh(oTmpEntity, RenderModelType.eOldData, lTexMod)
                                                    Else
                                                        RenderModel(oTmpEntity, lTexMod, RenderModelType.eOldData)
                                                    End If
                                                End If
                                            ElseIf yVisible = eVisibilityType.InMaxRange Then
                                                If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                                                    bHasLongRangeRadar = True
                                                Else : RenderRadarObject(oTmpEntity)
                                                End If
                                            End If
                                        End If
                                    End With
                                End If
                            Next X

                            sDrawSceneLocation = "EndOfRenderEntities"
                            If bSaveTacticalData = True Then oEnvir.SaveEnvironmentTacticalData()

                            If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                                moModelShader.EndRender()

                                'TODO: Hate this double loop crap...
                                If bHasIronCurtains = True OrElse bHasLongRangeRadar = True Then
                                    sDrawSceneLocation = "RenderRadarAndInvul"
                                    For X = 0 To oEnvir.lEntityUB
                                        If oEnvir.lEntityIdx(X) <> -1 Then
                                            sDrawSceneLocation = "GetEntityFromArray"
                                            Dim oTmpEntity As BaseEntity = oEnvir.oEntity(X)
                                            If oTmpEntity Is Nothing = False AndAlso oTmpEntity.bCulled = False Then
                                                With oTmpEntity
                                                    If bHasIronCurtains = True AndAlso (.yVisibility = eVisibilityType.Visible OrElse .yVisibility = eVisibilityType.FacilityIntel) Then
                                                        If oEnvir.IsPlayerIronCurtain(.OwnerID) = True Then
                                                            sDrawSceneLocation = "RenderIronCurtainModel"
                                                            Dim lRenderModelType As RenderModelType = RenderModelType.eIronCurtain
                                                            If .yVisibility = eVisibilityType.FacilityIntel Then lRenderModelType = lRenderModelType Or RenderModelType.eOldData
                                                            If .bSelected = True Then lRenderModelType = lRenderModelType Or RenderModelType.eSelected
                                                            RenderModel(oTmpEntity, 0, lRenderModelType)
                                                        End If
                                                    ElseIf .yVisibility = eVisibilityType.InMaxRange Then
                                                        sDrawSceneLocation = "RenderRadarObject"
                                                        RenderRadarObject(oTmpEntity)
                                                    End If
                                                End With
                                            End If
                                        End If
                                    Next X
                                End If
                            End If

                            If alTethers Is Nothing = False AndAlso alTethers.Count > 0 Then
                                sDrawSceneLocation = "RenderTetherPoints"
                                RenderTetherPoints(alTethers)
                            End If
                        End If

                        sDrawSceneLocation = "FinalizeModelRender"
                        If mbSupportsNewModelMethod = True Then ResetModelTextureStates()
                        If gbMonitorPerformance = True Then gsw_Models.Stop()

                        If gbMonitorPerformance = True Then gsw_CommitEntitySoundChanges.Start()
                        sDrawSceneLocation = "CommitEntitySoundChanges"
                        If goSound Is Nothing = False Then goSound.CommitEntitySoundChanges()
                        If gbMonitorPerformance = True Then gsw_CommitEntitySoundChanges.Stop()

                        If mbCaptureScreenshot = False AndAlso bRenderSelectedRanges = True AndAlso alSelectedIdx Is Nothing = False AndAlso alSelectedIdx.Count > 0 Then
                            sDrawSceneLocation = "RenderSelectedIdxRanges"
                            RenderSelectedRanges(alSelectedIdx)
                        End If

                        'Render our mineral cache's
                        sDrawSceneLocation = "CacheRender"
                        If muSettings.RenderMineralCaches = True OrElse (goUILib Is Nothing = False AndAlso goUILib.BuildGhost Is Nothing = False) OrElse (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255) Then
                            If gbMonitorPerformance = True Then gsw_Caches.Start()
                            glCachesRendered = 0

                            'TODO: Should determine this on the Region Server before sending the mineral caches down!!!
                            If oEnvir.HasUnitsHere() = True Then

                                'Go through, if it is within the clipper, render it
                                sDrawSceneLocation = "InitializeMinCacheResources"
                                If MineralCache.CacheMesh Is Nothing Then MineralCache.Initialize3DResources()
                                moDevice.Material = MineralCache.GetCacheMaterial()

                                For X = 0 To oEnvir.lCacheUB
                                    If oEnvir.lCacheIdx(X) <> -1 Then
                                        sDrawSceneLocation = "GetCacheFromArray"
                                        Dim oCache As MineralCache = oEnvir.oCache(X)
                                        If oCache Is Nothing Then Continue For

                                        With oCache
                                            'Ensure our Y value is set...
                                            .lMapWrapLocX = .LocX

                                            'If yMapWrapSituation <> 0 AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet Then
                                            '    'Ok, in a mapwrap situation
                                            '    If yMapWrapSituation = 1 Then
                                            '        'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
                                            '        If .LocX > lLocXMapWrapCheck Then
                                            '            .lMapWrapLocX = .LocX - oEnvir.lMapWrapAdjustX
                                            '        End If
                                            '    Else
                                            '        'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
                                            '        If .LocX < lLocXMapWrapCheck Then
                                            '            .lMapWrapLocX = .LocX + oEnvir.lMapWrapAdjustX
                                            '        End If
                                            '    End If
                                            'End If

                                            If .LocY = Int32.MinValue OrElse goCamera.CullObject(CullBox.GetCullBox(.lMapWrapLocX, .LocY, .LocZ, -10, -10, -10, 10, 10, 10)) = False Then
                                                sDrawSceneLocation = "DrawMinCache"
                                                glCachesRendered += 1
                                                moDevice.Transform.World = .GetWorldMatrix()

                                                'Ok, ow, if this is a mineralcache...
                                                If .ObjTypeID = ObjectType.eMineralCache Then
                                                    sDrawSceneLocation = "DrawMinCache.MinCache"
                                                    moDevice.SetTexture(0, MineralCache.texCache)
                                                    If .CacheTypeID = MineralCacheType.eMineable Then
                                                        sDrawSceneLocation = "DrawMinCache.MinCache.Mineable"
                                                        If MineralCache.CacheMesh Is Nothing OrElse MineralCache.CacheMesh(.ModelIndex) Is Nothing Then MineralCache.Initialize3DResources()
                                                        If MineralCache.CacheMesh Is Nothing = False AndAlso MineralCache.CacheMesh(.ModelIndex) Is Nothing = False Then MineralCache.CacheMesh(.ModelIndex).DrawSubset(0)
                                                    Else
                                                        'ok, its pickupable... 
                                                        sDrawSceneLocation = "DrawMinCache.MinCache.Pickup"
                                                        If MineralCache.DebrisMesh Is Nothing = False Then MineralCache.DebrisMesh.DrawSubset(0) 'Else MineralCache.Initialize3DResources()
                                                    End If
                                                Else
                                                    'its a component cache...
                                                    sDrawSceneLocation = "DrawMinCache.CompCache"
                                                    If MineralCache.DebrisTex Is Nothing OrElse MineralCache.DebrisMesh Is Nothing Then MineralCache.Initialize3DResources()
                                                    If MineralCache.DebrisTex Is Nothing = False Then moDevice.SetTexture(0, MineralCache.DebrisTex(.lDebrisTexIdx))
                                                    If MineralCache.DebrisMesh Is Nothing = False Then MineralCache.DebrisMesh.DrawSubset(0)
                                                End If

                                            End If
                                        End With
                                    End If
                                Next X
                            End If
                            If gbMonitorPerformance = True Then gsw_Caches.Stop()
                        End If
                        sDrawSceneLocation = "EndOfCacheRender"

                        If bInCtrlMove = True Then
                            sDrawSceneLocation = "RenderDirectionalMoving"
                            RenderDirectionalMoving()
                        End If

                        'Render our BuildGhost
                        If goUILib Is Nothing = False Then
                            If goUILib.BuildGhost Is Nothing = False Then
                                sDrawSceneLocation = "RenderBuildGhost"
                                RenderBuildGhost()
                            End If
                        End If

                        'Render facility constructions...
                        If gbMonitorPerformance = True Then gsw_Models.Start()
                        sDrawSceneLocation = "RenderEntityConstructions"
                        RenderEntityConstructions()
                        If gbMonitorPerformance = True Then gsw_Models.Stop()

                        If glCurrentEnvirView = CurrentView.ePlanetView Then
                            If goCurrentEnvir Is Nothing = False Then
                                If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                                    sDrawSceneLocation = "DoWaterRenderPass"
                                    CType(goCurrentEnvir.oGeoObject, Planet).DoWaterRenderPass()
                                End If
                            End If
                        End If

                        'render our system minimap
                        If glCurrentEnvirView = CurrentView.eSystemView Then
                            If gbMonitorPerformance = True Then gsw_Minimap.Start()
                            sDrawSceneLocation = "RenderSystemMiniMap"
                            RenderSystemMiniMap()
                            If gbMonitorPerformance = True Then gsw_Minimap.Stop()
                        ElseIf glCurrentEnvirView = CurrentView.ePlanetView Then
                            If gbMonitorPerformance = True Then gsw_Minimap.Start()
                            sDrawSceneLocation = "RenderPlanetMiniMap"
                            RenderPlanetMiniMap()
                            If gbMonitorPerformance = True Then gsw_Minimap.Stop()
                        End If
                    End If
                End If


                If muSettings.ShowTargetBoxes = True Then
                    sDrawSceneLocation = "RenderEntityTargets"
                    RenderEntityTargets()
                End If

                If muSettings.RenderBurnMarks = True Then
                    sDrawSceneLocation = "BurnMark.EndFrame"
                    moDevice.Transform.World = Matrix.Identity
                    goBurnMarkMgr.EndFrame()
                End If

                If goUILib Is Nothing = False AndAlso goUILib.lUISelectState = UILib.eSelectState.eBombTargetSelection Then
                    sDrawSceneLocation = "RenderBombardmentReticle"
                    RenderBombardmentReticle()
                End If

                'If muSettings.DrawGrid = True Then DrawLargeGrid()

            ElseIf glCurrentEnvirView = CurrentView.eSystemMapView2 OrElse glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                UpdateCP()
                If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    sDrawSceneLocation = "RenderSysViewMap"
                    Dim alList As ArrayList = RenderSysViewMap()
                    If alList Is Nothing = False AndAlso alList.Count > 0 AndAlso bRenderSelectedRanges = True Then
                        RenderSelectedRanges(alList)
                    End If
                End If
            ElseIf glCurrentEnvirView = CurrentView.ePlanetMapView Then
                UpdateCP()
                If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                    Try
                        If gbMonitorPerformance = True Then gsw_Geography.Start()
                        RenderPlanetMapBlips(False)
                        If gbMonitorPerformance = True Then gsw_Geography.Stop()
                    Catch
                        'Do nothing for now
                        'Stop
                    End Try
                End If
            ElseIf glCurrentEnvirView = CurrentView.eGalaxyMapView Then
                If Not muSettings.gbGalaxyControlHideFleetMovement = True Then
                    sDrawSceneLocation = "RenderGalaxyFleetMovements"
                    RenderGalaxyFleetMovements()
                End If
                If Not muSettings.gbGalaxyControlHideWormholes = True Then RenderWormholes()
                If goGalaxy.GalaxySelectionIdx > 0 AndAlso Not goGalaxy.moSystems(goGalaxy.GalaxySelectionIdx) Is Nothing Then
                    RenderSolarSystemSelectionBox(goGalaxy.moSystems(goGalaxy.GalaxySelectionIdx))
                End If
            End If

            'Render our special fx
            sDrawSceneLocation = "SpecialFX.Wpns"
            Dim bUpdateNoRender As Boolean = (glCurrentEnvirView <> CurrentView.eSystemView) AndAlso (glCurrentEnvirView <> CurrentView.ePlanetView)
            If gbMonitorPerformance = True Then gsw_WpnFX.Start()
            goWpnMgr.RenderFX(bUpdateNoRender)
            If gbMonitorPerformance = True Then gsw_WpnFX.Stop()

            sDrawSceneLocation = "SpecialFX.Bombs"
            If gbMonitorPerformance = True Then gsw_BombFX.Start()
            BombMgr.goBombMgr.Update(bUpdateNoRender)
            If gbMonitorPerformance = True Then gsw_BombFX.Stop()

            sDrawSceneLocation = "SpecialFX.Shlds"
            If gbMonitorPerformance = True Then gsw_ShieldFX.Start()
            goShldMgr.RenderFX(bUpdateNoRender)
            If gbMonitorPerformance = True Then gsw_ShieldFX.Stop()

            sDrawSceneLocation = "SpecialFX.Burn"
            If gbMonitorPerformance = True Then gsw_BurnFX.Start()
            goPFXEngine32.Render(bUpdateNoRender)
            sDrawSceneLocation = "SpecialFX.Missile"
            If goMissileMgr Is Nothing = False Then goMissileMgr.Render(bUpdateNoRender)
            If gbMonitorPerformance = True Then gsw_BurnFX.Stop()

            sDrawSceneLocation = "SpecialFX.EntityDeath"
            If gbMonitorPerformance = True Then gsw_DeathFX.Start()
            goEntityDeath.RenderSequences(bUpdateNoRender)
            If gbMonitorPerformance = True Then gsw_DeathFX.Stop()

            sDrawSceneLocation = "SpecialFX.Fireworks"
            If gbMonitorPerformance = True Then gsw_FireworksFX.Start()
            If goFireworks Is Nothing = False Then goFireworks.Render(bUpdateNoRender)
            If gbMonitorPerformance = True Then gsw_FireworksFX.Stop()

            sDrawSceneLocation = "SpecialFX.ExplMgr"
            If gbMonitorPerformance = True Then gsw_Explosions.Start()
            If goExplMgr Is Nothing = False Then goExplMgr.Render(bUpdateNoRender)
            If gbMonitorPerformance = True Then gsw_Explosions.Stop()

            'sDrawSceneLocation = "SpecialFX.Warpoints"
            'If goRewards Is Nothing = False Then goRewards.RenderRewards(bUpdateNoRender)

            sDrawSceneLocation = "SpecialFX.AOEExplMgr"
            If gbMonitorPerformance = True Then gsw_AOEExplosions.Start()
            If AOEExplosionMgr.goMgr Is Nothing = False Then AOEExplosionMgr.goMgr.Update()
            If gbMonitorPerformance = True Then gsw_AOEExplosions.Stop()

            'If moCityCars Is Nothing = False Then moCityCars.Render(False)

            If goWormholeMgr Is Nothing = False AndAlso (glCurrentEnvirView = CurrentView.eSystemMapView1 OrElse glCurrentEnvirView = CurrentView.eSystemMapView2 OrElse glCurrentEnvirView = CurrentView.eSystemView) Then
                moDevice.Transform.World = Matrix.Identity
                goCamera.SetupMatrices(moDevice, glCurrentEnvirView)
                sDrawSceneLocation = "SpecialFX.WHMgr"
                goWormholeMgr.RenderFX(bUpdateNoRender)
            End If

            If glCurrentEnvirView < CurrentView.eFullScreenInterface AndAlso glCurrentEnvirView <> CurrentView.ePlanetMapView Then
                sDrawSceneLocation = "PostGlow"
                If muSettings.PostGlowAmt > 0 Then
                    If gbMonitorPerformance = True Then gsw_PostEffects.Start()
                    If moPost Is Nothing Then moPost = New PostShader()
                    moPost.ExecutePostProcess()
                    If gbMonitorPerformance = True Then gsw_PostEffects.Stop()
                ElseIf moPost Is Nothing = False Then
                    moPost.DisposeMe()
                    moPost = Nothing
                End If
            End If

            'Before interfaces but after all else, render our hitpoint bars
            sDrawSceneLocation = "HPBarsAndWPInds"
            If gbMonitorPerformance = True Then gsw_HPBars.Start()
            RenderHitPointBarsAndWPInds(lCameraCellX, lCameraCellZ)
            If gbMonitorPerformance = True Then gsw_HPBars.Stop()

            'Now, draw our interface objects
            sDrawSceneLocation = "Normal.RenderInterfaces"
            If gbMonitorPerformance = True Then gsw_UI.Start()
            goUILib.RenderInterfaces(glCurrentEnvirView)
            If gbMonitorPerformance = True Then gsw_UI.Stop()

            'And our selection box
            If goCamera.ShowSelectionBox = True Then
                sDrawSceneLocation = "SelectBox"

                Dim vSelBox(4) As Vector2
                vSelBox(0).X = goCamera.SelectionBoxStart.X : vSelBox(0).Y = goCamera.SelectionBoxStart.Y
                vSelBox(1).X = goCamera.SelectionBoxEnd.X : vSelBox(1).Y = goCamera.SelectionBoxStart.Y
                vSelBox(2).X = goCamera.SelectionBoxEnd.X : vSelBox(2).Y = goCamera.SelectionBoxEnd.Y
                vSelBox(3).X = goCamera.SelectionBoxStart.X : vSelBox(3).Y = goCamera.SelectionBoxEnd.Y
                vSelBox(4).X = goCamera.SelectionBoxStart.X : vSelBox(4).Y = goCamera.SelectionBoxStart.Y

                If GFXEngine.gbDeviceLost = False AndAlso GFXEngine.gbPaused = False Then
                    BPLine.DrawLine(2, True, vSelBox, System.Drawing.Color.FromArgb(255, 0, 255, 255))
                End If
            End If
        End If

        'then setup our matrices, and do it
        sDrawSceneLocation = "SetupMatrices"
        goCamera.SetupMatrices(moDevice, glCurrentEnvirView)

        sDrawSceneLocation = "EndScene"
        moDevice.EndScene()

        'check for capture screenshot request
        If mbCaptureScreenshot = True Then ' OrElse mbCaptureScenes = True Then
            sDrawSceneLocation = "DoCaptureScreenshot"
            DoCaptureScreenshot()
            mbCaptureScreenshot = False
        End If

        If gbMonitorPerformance = True Then gsw_Present.Start()
        sDrawSceneLocation = "Present"
        PresentTheScene()
        If gbMonitorPerformance = True Then gsw_Present.Stop()
        Return lRenderCnt

    End Function

    Private Sub DrawLargeGrid()

        'If goCurrentEnvir Is Nothing Then Return
        'If goCurrentEnvir.oGeoObject Is Nothing Then Return

        Dim lGridSquareSize As Int32 = 8000
        'Dim lGridsPerRow As Int32 = 0
        'Dim lEnvirSize As Int32 = 0
        'If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
        '    lEnvirSize = 10000000
        'Else
        '    lEnvirSize = CType(goCurrentEnvir.oGeoObject, Planet).CellSpacing * TerrainClass.Width
        'End If
        'lGridsPerRow = lEnvirSize \ lGridSquareSize

        'Dim lTemp As Int32 = goCamera.mlCameraAtZ

        'Dim lLargeSector As Int32 = (lTemp \ lGridSquareSize) * lGridsPerRow
        'lLargeSector += goCamera.mlCameraAtX \ lGridSquareSize

        Dim lBaseX As Int32 = goCamera.mlCameraAtX \ lGridSquareSize
        Dim lBaseZ As Int32 = goCamera.mlCameraAtZ \ lGridSquareSize

        'Now, draw our grid around us...
        lBaseX *= lGridSquareSize
        lBaseZ *= lGridSquareSize

        lBaseX -= (lGridSquareSize * 2)
        lBaseZ -= (lGridSquareSize * 2)

        Dim lFirstX As Int32 = lBaseX
        Dim lFirstZ As Int32 = lBaseZ

        'center square
        Dim lClr As Int32 = System.Drawing.Color.FromArgb(32, 255, 255, 255).ToArgb
        Dim lSqrSize As Int32 = lGridSquareSize - 10
        Dim uVerts(149) As CustomVertex.PositionColored
        Dim lIdx As Int32 = -1

        For Y As Int32 = 0 To 4
            lBaseX = lFirstX

            For X As Int32 = 0 To 4
                lIdx += 1
                uVerts(lIdx) = New CustomVertex.PositionColored(lBaseX, 0, lBaseZ, lClr)
                lIdx += 1
                uVerts(lIdx) = New CustomVertex.PositionColored(lBaseX + lSqrSize, 0, lBaseZ, lClr)
                lIdx += 1
                uVerts(lIdx) = New CustomVertex.PositionColored(lBaseX, 0, lBaseZ + lSqrSize, lClr)

                lIdx += 1
                uVerts(lIdx) = New CustomVertex.PositionColored(lBaseX + lSqrSize, 0, lBaseZ, lClr)
                lIdx += 1
                uVerts(lIdx) = New CustomVertex.PositionColored(lBaseX, 0, lBaseZ + lSqrSize, lClr)
                lIdx += 1
                uVerts(lIdx) = New CustomVertex.PositionColored(lBaseX + lSqrSize, 0, lBaseZ + lSqrSize, lClr)

                lBaseX += lGridSquareSize
            Next X

            lBaseZ += lGridSquareSize
        Next Y

        'reset identity to nothing...
        moDevice.Transform.World = Matrix.Identity
        moDevice.RenderState.CullMode = Cull.None
        moDevice.RenderState.ZBufferWriteEnable = False
        moDevice.RenderState.SourceBlend = Blend.SourceAlpha
        moDevice.RenderState.DestinationBlend = Blend.One
        moDevice.RenderState.AlphaBlendEnable = True

        moDevice.RenderState.Lighting = False
        moDevice.SetTexture(0, Nothing)
        moDevice.VertexFormat = CustomVertex.PositionColored.Format

        Dim oMat As Material
        With oMat
            .Ambient = Color.White
            .Diffuse = Color.White
            .Emissive = Color.White
            .Specular = Color.White
        End With
        moDevice.Material = oMat

        moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, 50, uVerts)

        moDevice.RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
        moDevice.RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero 
        moDevice.RenderState.ZBufferWriteEnable = True

        moDevice.RenderState.Lighting = True
    End Sub

    Private Sub RenderSelectedRanges(ByVal alSelectedIdx As ArrayList)
        If mbCaptureScreenshot = True Then Return
        If alSelectedIdx Is Nothing OrElse alSelectedIdx.Count < 1 Then Return
        'If glCurrentEnvirView <> CurrentView.eSystemView AndAlso glCurrentEnvirView <> CurrentView.ePlanetMapView AndAlso glCurrentEnvirView <> CurrentView.ePlanetView AndAlso glCurrentEnvirView <> CurrentView.eSystemMapView2 Then Return
        If glCurrentEnvirView <> CurrentView.eSystemView AndAlso glCurrentEnvirView <> CurrentView.ePlanetView AndAlso glCurrentEnvirView <> CurrentView.eSystemMapView2 Then Return

        Dim uVerts((alSelectedIdx.Count * 18) - 1) As CustomVertex.PositionColoredTextured
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        Dim lVertIdx As Int32 = 0
        Dim lPrimCnt As Int32 = 0

        Try
            Dim lDetRng As Int32 = gl_FINAL_GRID_SQUARE_SIZE * 4
            Dim fMult As Single = 1.0F

            If glCurrentEnvirView = CurrentView.eSystemMapView2 Then fMult = 1.0F / 30.0F
            For X As Int32 = 0 To alSelectedIdx.Count - 1
                Dim lIdx As Int32 = CInt(alSelectedIdx(X))
                Dim oEntity As BaseEntity = oEnvir.oEntity(lIdx)
                If oEntity Is Nothing = False AndAlso oEntity.oUnitDef Is Nothing = False Then

                    With oEntity
                        'Ok, first quad, optimum range
                        Dim lLocY As Int32 = CInt(.LocY)
                        If glCurrentEnvirView = CurrentView.eSystemMapView2 Then lLocY = 0
                        Dim fHlfVal As Single = .oUnitDef.FOW_OptRadarRange '* 0.5F
                        Dim lClrVal As Int32 = System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb
                        'triangle 1
                        uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 0, 0) : lVertIdx += 1
                        uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 0, 1) : lVertIdx += 1
                        uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 1, 0) : lVertIdx += 1
                        lPrimCnt += 1
                        'triangle 2
                        uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 0, 1) : lVertIdx += 1
                        uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 1, 0) : lVertIdx += 1
                        uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 1, 1) : lVertIdx += 1
                        lPrimCnt += 1

                        'Ok, second quad, maximum range
                        If .oUnitDef.FOW_MaxRadarRange > (.oMesh.RangeOffset * lDetRng) + 0.05F Then
                            fHlfVal = .oUnitDef.FOW_MaxRadarRange '* 0.5F
                            lClrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128).ToArgb
                            'triangle 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 0, 0) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 0, 1) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 1, 0) : lVertIdx += 1
                            lPrimCnt += 1
                            'triangle 2
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 0, 1) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 1, 0) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 1, 1) : lVertIdx += 1
                            lPrimCnt += 1
                        End If

                        'Finally, third quad, weapon range...
                        If .lMaxWpnRngValue = -1 Then
                            .lMaxWpnRngValue = 0

                            'Request the maximum Wpn Range value here...

                            Dim yMsg(7) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestMaxWpnRng).CopyTo(yMsg, 0)
                            .GetGUIDAsString.CopyTo(yMsg, 2)
                            goUILib.SendMsgToPrimary(yMsg)
                        End If

                        If .lMaxWpnRngValue > 0 Then
                            fHlfVal = .lMaxWpnRngValue '* gl_FINAL_GRID_SQUARE_SIZE
                            lClrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0).ToArgb
                            'triangle 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 0, 0) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 0, 1) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 1, 0) : lVertIdx += 1
                            lPrimCnt += 1
                            'triangle 2
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX - fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 0, 1) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ - fHlfVal) * fMult, lClrVal, 1, 0) : lVertIdx += 1
                            uVerts(lVertIdx) = New CustomVertex.PositionColoredTextured((.LocX + fHlfVal) * fMult, lLocY, (.LocZ + fHlfVal) * fMult, lClrVal, 1, 1) : lVertIdx += 1
                            lPrimCnt += 1
                        End If
                    End With
                End If
            Next X
        Catch
        End Try

        Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
        Dim lPeevLighting As Boolean = moDevice.RenderState.Lighting
        Dim lPrevCullMode As Cull = moDevice.RenderState.CullMode
        Dim lPrevFog As Boolean = moDevice.RenderState.FogEnable

        Dim lMin As TextureFilter
        Dim lMip As TextureFilter
        With moDevice.RenderState
            .Lighting = False
            .ZBufferEnable = False
            .FogEnable = False
            .CullMode = Cull.None

            If glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                lMin = moDevice.SamplerState(0).MinFilter
                lMip = moDevice.SamplerState(0).MipFilter

                moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, TextureFilter.None)
                moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, TextureFilter.None)
            End If
        End With
        'If moTex Is Nothing Then moTex = TextureLoader.FromFile(moDevice, "C:\visring.dds")


        moDevice.VertexFormat = CustomVertex.PositionColoredTextured.Format
        moDevice.SetTexture(0, goResMgr.GetTexture("visring.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "ob.pak"))
        moDevice.Transform.World = Matrix.Identity
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, lPrimCnt, uVerts)

        With moDevice.RenderState
            .Lighting = lPeevLighting 'True
            .ZBufferEnable = lPrevZBuffer 'True
            .FogEnable = lPrevFog
            .CullMode = lPrevCullMode 'Cull.CounterClockwise

            If glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, lMin)
                moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, lMip)
            End If
        End With

    End Sub

    Private mfTetherPointRotation As Single = 0.0F
    Private Sub RenderTetherPoints(ByVal alTethers As ArrayList)
        If alTethers Is Nothing OrElse alTethers.Count < 1 Then Return
        If glCurrentEnvirView <> CurrentView.eSystemView AndAlso glCurrentEnvirView <> CurrentView.ePlanetView Then Return

        Dim uVerts((alTethers.Count * 6) - 1) As CustomVertex.PositionColored
        Dim lVertIdx As Int32 = 0

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        Dim yHtLookup As Byte = 0       '0 none, 1 = planet, 2 = system
        Dim oPlanet As Planet = Nothing
        Dim oSystem As SolarSystem = Nothing

        mfTetherPointRotation += 1.0F

        Dim fExtX As Single = 50
        Dim fExtZ As Single = 0
        Dim fOppX As Single = 0
        Dim fOppZ As Single = 50
        RotatePoint(0, 0, fExtX, fExtZ, mfTetherPointRotation)
        RotatePoint(0, 0, fOppX, fOppZ, mfTetherPointRotation)

        Try
            If oEnvir Is Nothing = False AndAlso oEnvir.oGeoObject Is Nothing = False Then
                If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                    yHtLookup = 1
                    If CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID <> ObjectType.ePlanet Then Return
                    oPlanet = CType(oEnvir.oGeoObject, Planet)
                Else
                    yHtLookup = 2
                    If CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID <> ObjectType.eSolarSystem Then Return
                    oSystem = CType(oEnvir.oGeoObject, SolarSystem)
                End If
            End If

            If yHtLookup = 1 Then
                If oPlanet Is Nothing Then yHtLookup = 0
            ElseIf yHtLookup = 2 Then
                If oSystem Is Nothing Then yHtLookup = 0
            End If

            Dim lClr As Int32
            If yHtLookup = 1 Then
                Dim clrGhostClr As System.Drawing.Color = System.Drawing.Color.FromArgb(64, 255, 255, 255)
                Select Case CType(oPlanet.MapTypeID, PlanetType)
                    Case PlanetType.eAcidic
                        clrGhostClr = muSettings.AcidBuildGhost
                    Case PlanetType.eAdaptable
                        clrGhostClr = muSettings.AdaptableBuildGhost
                    Case PlanetType.eBarren
                        clrGhostClr = muSettings.BarrenBuildGhost
                    Case PlanetType.eDesert
                        clrGhostClr = muSettings.DesertBuildGhost
                    Case PlanetType.eGeoPlastic
                        clrGhostClr = muSettings.LavaBuildGhost
                    Case PlanetType.eTerran
                        clrGhostClr = muSettings.TerranBuildGhost
                    Case PlanetType.eTundra
                        clrGhostClr = muSettings.IceBuildGhost
                    Case PlanetType.eWaterWorld
                        clrGhostClr = muSettings.WaterworldBuildGhost
                End Select
                lClr = clrGhostClr.ToArgb
            Else
                lClr = System.Drawing.Color.FromArgb(64, 255, 255, 255).ToArgb
            End If

            For X As Int32 = 0 To alTethers.Count - 1
                'Ok, render the tether points...
                Dim ptLoc As Point = CType(alTethers(X), Point)

                Dim fHt As Single = 0
                If yHtLookup = 1 Then
                    fHt = oPlanet.GetHeightAtPoint(ptLoc.X, ptLoc.Y, True)
                    '                ElseIf yHtLookup = 2 Then
                    '                   fHt = oSystem.GetBaseY(ptLoc.X, ptLoc.Y)
                End If
                uVerts(lVertIdx) = New CustomVertex.PositionColored(ptLoc.X, fHt, ptLoc.Y, lClr) : lVertIdx += 1
                uVerts(lVertIdx) = New CustomVertex.PositionColored(ptLoc.X - fExtX, fHt + 150, ptLoc.Y - fExtZ, lClr) : lVertIdx += 1
                uVerts(lVertIdx) = New CustomVertex.PositionColored(ptLoc.X + fExtX, fHt + 150, ptLoc.Y + fExtZ, lClr) : lVertIdx += 1

                uVerts(lVertIdx) = New CustomVertex.PositionColored(ptLoc.X, fHt, ptLoc.Y, lClr) : lVertIdx += 1
                uVerts(lVertIdx) = New CustomVertex.PositionColored(ptLoc.X - fOppX, fHt + 150, ptLoc.Y - fOppZ, lClr) : lVertIdx += 1
                uVerts(lVertIdx) = New CustomVertex.PositionColored(ptLoc.X + fOppX, fHt + 150, ptLoc.Y + fOppZ, lClr) : lVertIdx += 1
            Next X

            Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
            Dim lPeevLighting As Boolean = moDevice.RenderState.Lighting
            Dim lPrevCullMode As Cull = moDevice.RenderState.CullMode
            Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
            Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend

            With moDevice.RenderState
                .SourceBlend = Blend.SourceAlpha
                .DestinationBlend = Blend.One
                .AlphaBlendEnable = True
                .ZBufferWriteEnable = False
                .Lighting = False
                .CullMode = Cull.None
            End With
            With moDevice
                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
            End With

            moDevice.Transform.World = Matrix.Identity
            moDevice.SetTexture(0, Nothing)
            moDevice.VertexFormat = CustomVertex.PositionColored.Format
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, alTethers.Count * 2, uVerts)

            With moDevice.RenderState
                'Then, reset our device...
                .CullMode = lPrevCullMode 'Cull.CounterClockwise
                .ZBufferWriteEnable = lPrevZBuffer 'True
                .Lighting = lPeevLighting 'True
                .SourceBlend = lPrevSrcBlnd 'Blend.SourceAlpha 'Blend.One
                .DestinationBlend = lPrevDestBlnd  ' Blend.InvSourceAlpha 'Blend.Zero
            End With
            With moDevice
                .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
            End With

        Catch
        End Try

    End Sub
    'Private mbCaptureScenes As Boolean = False
    'Public Sub ToggleCaptureScenes()
    '    mbCaptureScenes = Not mbCaptureScenes
    'End Sub

	Private Sub PresentTheScene()

        Try
            moDevice.Present()
        Catch exDevLost As DeviceLostException
            'Ok, the device is lost... set our state to lost essentially
            gbDeviceLost = True
            gbPaused = True
        Catch exDriverInternalError As DriverInternalErrorException
            gbDeviceLost = True
            gbPaused = True
        Catch
            Dim b As Boolean = False
        End Try


	End Sub

    Private Sub RenderPlanetMapBlips(ByVal bForMiniMap As Boolean)

        If gbDeviceLost = True OrElse gbPaused = True Then Return
        If HasAliasedRights(AliasingRights.eViewUnitsAndFacilities) = False Then Return

        Dim X As Int32
        Dim iGridLoc(,) As Int16
        Dim lLocX As Int32
        Dim lLocZ As Int32
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        If oEnvir.oGeoObject Is Nothing Then Return

        Dim lHalfExtent As Int32 = CType(oEnvir.oGeoObject, Planet).GetExtent \ 2
        Dim lFullExtent As Int32 = CType(oEnvir.oGeoObject, Planet).GetExtent
        Dim lCellSpacing As Int32 = lHalfExtent \ 120 ' CInt(lHalfExtent \ 120)

        Dim iVal As Int16       '0 = none, 1 = neutral, 2 = player, 4 = ally, 8 = enemy
        Dim yVisible As Byte
        Dim yRel As Byte

        'ReDim iGridLoc(TerrainClass.Width - 1, TerrainClass.Height - 1)
        ReDim iGridLoc(255, 255)

        If lCellSpacing = 0 Then Return

        Dim alSelectedIdx As New ArrayList()

        Dim bCheckGuild As Boolean = False
        Dim oGuild As Guild = Nothing
        If goCurrentPlayer Is Nothing = False Then
            oGuild = goCurrentPlayer.oGuild
            If oGuild Is Nothing = False Then
                bCheckGuild = ((goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0)
            End If
        End If
        Try
            Dim bSaveTacticalData As Boolean = False
            For X = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) <> -1 Then
                    With oEnvir.oEntity(X)
                        'lLocX = CInt(.LocX + lHalfExtent) \ lCellSpacing
                        'lLocZ = CInt(.LocZ + lHalfExtent) \ lCellSpacing
                        lLocX = CInt(((.LocX + lHalfExtent) / lFullExtent) * 256.0F)
                        lLocZ = CInt(((.LocZ + lHalfExtent) / lFullExtent) * 256.0F)

                        If bForMiniMap = False Then
                            If .OwnerID = glPlayerID Then
                                If .bSelected = True AndAlso bRenderSelectedRanges = True Then
                                    alSelectedIdx.Add(X)
                                End If
                                If Not muSettings.gbPlanetMapDontShowMyUnits Then
                                    iVal = 2
                                Else
                                    iVal = 0
                                End If
                                yVisible = eVisibilityType.Visible
                            ElseIf (bCheckGuild = True AndAlso oGuild.MemberInGuild(.OwnerID) = True) Then
                                If muSettings.gbPlanetMapDontShowGuilds = False Then
                                    iVal = 16
                                    yVisible = eVisibilityType.Visible
                                Else : Continue For
                                End If
                            ElseIf .oMesh Is Nothing = False Then
                                yVisible = oEnvir.UnitInRadarRange(CInt(.LocX), CInt(.LocZ), .oMesh.RangeOffset)
                                If yVisible = eVisibilityType.Visible OrElse .yVisibility = eVisibilityType.FacilityIntel Then
                                    If oEnvir.oEntity(X).yRelID = Byte.MaxValue OrElse oEnvir.oEntity(X).yRelID = 0 Then
                                        yRel = goCurrentPlayer.GetPlayerRelScore(oEnvir.oEntity(X).OwnerID)
                                        oEnvir.oEntity(X).yRelID = yRel
                                    Else : yRel = oEnvir.oEntity(X).yRelID
                                    End If

                                    If yRel <= elRelTypes.eWar Then
                                        iVal = 8
                                    ElseIf yRel <= elRelTypes.ePeace Then       'elRelTypes.eNeutral
                                        iVal = 1
                                    Else : iVal = 4
                                    End If
                                ElseIf yVisible = eVisibilityType.InMaxRange Then
                                    iVal = 1
                                Else : iVal = 0
                                End If
                                Try
                                    If yRel <= 0 Then
                                        If muSettings.gbPlanetMapDontShowUnknown Then iVal = 0
                                    ElseIf yRel <= elRelTypes.eWar Then
                                        If muSettings.gbPlanetMapDontShowEnemy Then iVal = 0
                                    ElseIf yRel <= elRelTypes.ePeace Then       'elRelTypes.eNeutral
                                        If muSettings.gbPlanetMapDontShowNeutral Then iVal = 0
                                    ElseIf (bCheckGuild = True AndAlso oGuild.MemberInGuild(oEnvir.oEntity(X).OwnerID)) = True Then
                                        If muSettings.gbPlanetMapDontShowGuilds Then iVal = 0
                                    Else
                                        If muSettings.gbPlanetMapDontShowAllied Then iVal = 0
                                    End If
                                Catch
                                End Try

                            Else : Continue For
                            End If

                            'TODO: For Full visibility, this is all you have to do!!!
                            'yVisible = eVisibilityType.Visible
                            'iVal = 8

                            If .yVisibility = eVisibilityType.Visible Then
                                If yVisible <> eVisibilityType.Visible AndAlso .ObjTypeID = ObjectType.eFacility Then
                                    yVisible = eVisibilityType.FacilityIntel
                                End If
                            ElseIf .yVisibility = eVisibilityType.FacilityIntel Then
                                If yVisible = eVisibilityType.Visible Then
                                    If .bObjectDestroyed = True Then

                                        If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.oGeoObject Is Nothing = False AndAlso CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                                            CType(oEnvir.oGeoObject, Planet).SetCityCreepLoc(.LocX, .LocZ, False, .oMesh.ShieldXZRadius)
                                        End If

                                        .ClearParticleFX()
                                        oEnvir.lEntityIdx(X) = -1
                                        oEnvir.oEntity(X) = Nothing
                                        bSaveTacticalData = True
                                        Continue For
                                    End If
                                Else : yVisible = eVisibilityType.FacilityIntel
                                End If
                            End If
                            'MSC: We only care about the mesh portion of the modelid
                            If oEnvir.oEntity(X).oMesh Is Nothing = False AndAlso ((oEnvir.oEntity(X).oMesh.lModelID And 255) = 24 OrElse (oEnvir.oEntity(X).yProductionType = ProductionType.eMining AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility)) Then
                                yVisible = 2 ' eVisibilityType.Visible
                            End If

                            If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso .ObjTypeID = ObjectType.eFacility Then
                                If yVisible >= eVisibilityType.Visible AndAlso .yVisibility < eVisibilityType.Visible Then
                                    'ok, let's ensure that our terrain vert is updated...
                                    If oEnvir.oGeoObject Is Nothing = False AndAlso CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                                        If .oMesh Is Nothing = False Then CType(oEnvir.oGeoObject, Planet).SetCityCreepLoc(.LocX, .LocZ, True, .oMesh.ShieldXZRadius)
                                    End If
                                End If
                            End If

                            .yVisibility = yVisible
                        Else
                            If .OwnerID = glPlayerID Then
                                iVal = 2
                                yVisible = eVisibilityType.Visible
                            Else
                                yVisible = .yVisibility
                                If yVisible = eVisibilityType.Visible OrElse .yVisibility = eVisibilityType.FacilityIntel Then
                                    If oEnvir.oEntity(X).yRelID = Byte.MaxValue OrElse oEnvir.oEntity(X).yRelID = 0 Then
                                        yRel = goCurrentPlayer.GetPlayerRelScore(oEnvir.oEntity(X).OwnerID)
                                        oEnvir.oEntity(X).yRelID = yRel
                                    Else : yRel = oEnvir.oEntity(X).yRelID
                                    End If

                                    If yRel <= elRelTypes.eWar Then
                                        iVal = 8
                                    ElseIf yRel <= elRelTypes.ePeace Then   'elRelTypes.eNeutral
                                        iVal = 1
                                    Else : iVal = 4
                                    End If
                                ElseIf yVisible = eVisibilityType.InMaxRange Then
                                    iVal = 1
                                Else : iVal = 0
                                End If
                            End If
                            'iVal = 8
                        End If

                        If lLocX > iGridLoc.GetUpperBound(0) Then lLocX = iGridLoc.GetUpperBound(0)
                        If lLocZ > iGridLoc.GetUpperBound(1) Then lLocZ = iGridLoc.GetUpperBound(1)
                        If lLocX < 0 Then lLocX = 0
                        If lLocZ < 0 Then lLocZ = 0

                        iGridLoc(lLocX, lLocZ) = iGridLoc(lLocX, lLocZ) Or iVal

                    End With
                End If
            Next X

            If bSaveTacticalData = True Then oEnvir.SaveEnvironmentTacticalData()
        Catch
        End Try

        'Now... ensure our texture is made
        If moPMapRadarTex Is Nothing Then
            Device.IsUsingEventHandlers = False
            moPMapRadarTex = New Texture(moDevice, 256, 256, 1, Usage.None, Format.A8R8G8B8, Pool.Managed)
            Device.IsUsingEventHandlers = True
        End If

        Dim pitch As Int32
        Dim yDest As Byte() = CType(moPMapRadarTex.LockRectangle(GetType(Byte), 0, LockFlags.Discard, pitch, 256 * 256 * 4), Byte())

        'Use 255 for purposes of texture mapping
        'Dim lMapX As Int32
        'Dim lMapY As Int32
        Dim Y As Int32
        If muSettings.gbPlanetMapBlinkUnits = True Then
            If msw_Utility Is Nothing Then msw_Utility = Stopwatch.StartNew
            If msw_Utility.IsRunning = False Then
                msw_Utility.Reset() : msw_Utility.Start()
            End If

            If msw_Utility.ElapsedMilliseconds > 30 Then
                mfFlashVal -= 0.1F
                If mfFlashVal < 0 Then mfFlashVal = 1.0F
                msw_Utility.Reset()
                msw_Utility.Start()
            End If
        Else
            mfFlashVal = 0.9F
        End If

        Dim fTempFlashVal As Single = mfFlashVal
        If bForMiniMap = True Then mfFlashVal = 1.0F

        If mfFlashVal > 0.98F AndAlso bForMiniMap = False Then
            For Y = 0 To 255
                'lMapY = CInt(Math.Floor((Y / 256.0F) * 240.0F))
                For lXVal As Int32 = 0 To 255
                    'For X = 0 To 255

                    'lMapX = CInt(Math.Floor((X / 256.0F) * 240.0F))

                    X = 255 - lXVal

                    '0 = none, 1 = neutral, 2 = player, 4 = ally, 8 = enemy
                    'Select Case iGridLoc(X, Y)
                    Select Case iGridLoc(lXVal, Y)
                        Case 0
                            yDest((Y * 256 + X) * 4 + 3) = 0            'a
                            yDest((Y * 256 + X) * 4 + 2) = 0            'r
                            yDest((Y * 256 + X) * 4 + 1) = 0            'g
                            yDest((Y * 256 + X) * 4 + 0) = 0            'b
                        Case Else
                            yDest((Y * 256 + X) * 4 + 3) = 255
                            yDest((Y * 256 + X) * 4 + 2) = 255            'r
                            yDest((Y * 256 + X) * 4 + 1) = 255          'g
                            yDest((Y * 256 + X) * 4 + 0) = 255            'b]

                            'If X > 0 Then
                            '	yDest((Y * 256 + (X - 1) * 4 + 3)) = 255
                            '	yDest((Y * 256 + (X - 1) * 4 + 2)) = 255
                            '	yDest((Y * 256 + (X - 1) * 4 + 1)) = 255
                            '	yDest((Y * 256 + (X - 1) * 4 + 0)) = 255
                            '	If Y > 0 Then
                            '		yDest(((Y - 1) * 256 + (X - 1) * 4 + 3)) = 255
                            '		yDest(((Y - 1) * 256 + (X - 1) * 4 + 2)) = 255
                            '		yDest(((Y - 1) * 256 + (X - 1) * 4 + 1)) = 255
                            '		yDest(((Y - 1) * 256 + (X - 1) * 4 + 0)) = 255
                            '	ElseIf Y < 255 Then
                            '		yDest(((Y + 1) * 256 + (X - 1) * 4 + 3)) = 255
                            '		yDest(((Y + 1) * 256 + (X - 1) * 4 + 2)) = 255
                            '		yDest(((Y + 1) * 256 + (X - 1) * 4 + 1)) = 255
                            '		yDest(((Y + 1) * 256 + (X - 1) * 4 + 0)) = 255
                            '	End If
                            'ElseIf X < 255 Then
                            '	yDest((Y * 256 + (X + 1) * 4 + 3)) = 255
                            '	yDest((Y * 256 + (X + 1) * 4 + 2)) = 255
                            '	yDest((Y * 256 + (X + 1) * 4 + 1)) = 255
                            '	yDest((Y * 256 + (X + 1) * 4 + 0)) = 255
                            '	If Y > 0 Then
                            '		yDest(((Y - 1) * 256 + (X + 1) * 4 + 3)) = 255
                            '		yDest(((Y - 1) * 256 + (X + 1) * 4 + 2)) = 255
                            '		yDest(((Y - 1) * 256 + (X + 1) * 4 + 1)) = 255
                            '		yDest(((Y - 1) * 256 + (X + 1) * 4 + 0)) = 255
                            '	ElseIf Y < 255 Then
                            '		yDest(((Y + 1) * 256 + (X + 1) * 4 + 3)) = 255
                            '		yDest(((Y + 1) * 256 + (X + 1) * 4 + 2)) = 255
                            '		yDest(((Y + 1) * 256 + (X + 1) * 4 + 1)) = 255
                            '		yDest(((Y + 1) * 256 + (X + 1) * 4 + 0)) = 255
                            '	End If
                            'End If
                    End Select
                Next lXVal
            Next Y
        Else

            Dim yNeutR As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.NeutralAssetColor.X * 255))))
            Dim yNeutG As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.NeutralAssetColor.Y * 255))))
            Dim yNeutB As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.NeutralAssetColor.Z * 255))))

            Dim yMyR As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.MyAssetColor.X * 255))))
            Dim yMyG As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.MyAssetColor.Y * 255))))
            Dim yMyB As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.MyAssetColor.Z * 255))))

            Dim yAllyR As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.AllyAssetColor.X * 255))))
            Dim yAllyG As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.AllyAssetColor.Y * 255))))
            Dim yAllyB As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.AllyAssetColor.Z * 255))))

            Dim yEnemyR As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.EnemyAssetColor.X * 255))))
            Dim yEnemyG As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.EnemyAssetColor.Y * 255))))
            Dim yEnemyB As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.EnemyAssetColor.Z * 255))))

            Dim yGuildR As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.GuildAssetColor.X * 255))))
            Dim yGuildG As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.GuildAssetColor.Y * 255))))
            Dim yGuildB As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.GuildAssetColor.Z * 255))))

            Dim yMyNeutR As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.MyAssetColor.X * muSettings.NeutralAssetColor.X * 255))))
            Dim yMyNeutG As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.MyAssetColor.Y * muSettings.NeutralAssetColor.Y * 255))))
            Dim yMyNeutB As Byte = CByte(Math.Max(0, Math.Min(255, CInt(muSettings.MyAssetColor.Z * muSettings.NeutralAssetColor.Z * 255))))

            For Y = 0 To 255
                'lMapY = CInt(Math.Floor((Y / 256.0F) * 240.0F))
                For lXVal As Int32 = 0 To 255
                    'For X = 0 To 255

                    'lMapX = CInt(Math.Floor((X / 256.0F) * 240.0F))

                    X = 255 - lXVal

                    '0 = none, 1 = neutral, 2 = player, 4 = ally, 8 = enemy
                    'Select Case iGridLoc(X, Y)
                    Select Case iGridLoc(lXVal, Y)
                        Case 0
                            yDest((Y * 256 + X) * 4 + 3) = 0            'a
                            yDest((Y * 256 + X) * 4 + 2) = 0            'r
                            yDest((Y * 256 + X) * 4 + 1) = 0            'g
                            yDest((Y * 256 + X) * 4 + 0) = 0            'b
                        Case 1      'Neutral
                            yDest((Y * 256 + X) * 4 + 3) = CByte(255 * mfFlashVal)          'a
                            yDest((Y * 256 + X) * 4 + 2) = yNeutR          'r 192
                            yDest((Y * 256 + X) * 4 + 1) = yNeutG          'g
                            yDest((Y * 256 + X) * 4 + 0) = yNeutB          'b 
                        Case 2, 6, 7      'Player and player with ally
                            yDest((Y * 256 + X) * 4 + 3) = CByte(255 * mfFlashVal)
                            yDest((Y * 256 + X) * 4 + 2) = yMyR            'r
                            yDest((Y * 256 + X) * 4 + 1) = yMyG          'g
                            yDest((Y * 256 + X) * 4 + 0) = yMyB            'b
                        Case 3      'Player and Neutral
                            yDest((Y * 256 + X) * 4 + 3) = CByte(255 * mfFlashVal)
                            yDest((Y * 256 + X) * 4 + 2) = yMyNeutR            'r
                            yDest((Y * 256 + X) * 4 + 1) = yMyNeutG            'g
                            yDest((Y * 256 + X) * 4 + 0) = yMyNeutB          'b
                        Case 4, 5      'Ally and Ally w/ Neutral
                            yDest((Y * 256 + X) * 4 + 3) = CByte(255 * mfFlashVal)
                            yDest((Y * 256 + X) * 4 + 2) = yAllyR            'r
                            yDest((Y * 256 + X) * 4 + 1) = yAllyG            'g
                            yDest((Y * 256 + X) * 4 + 0) = yAllyB            'b
                        Case 8          'enemy
                            yDest((Y * 256 + X) * 4 + 3) = CByte(255 * mfFlashVal)
                            yDest((Y * 256 + X) * 4 + 2) = yEnemyR            'r
                            yDest((Y * 256 + X) * 4 + 1) = yEnemyG            'g
                            yDest((Y * 256 + X) * 4 + 0) = yEnemyB            'b
                        Case Else       'guild
                            yDest((Y * 256 + X) * 4 + 3) = CByte(255 * mfFlashVal)
                            yDest((Y * 256 + X) * 4 + 2) = yGuildR            'r
                            yDest((Y * 256 + X) * 4 + 1) = yGuildG            'g
                            yDest((Y * 256 + X) * 4 + 0) = yGuildB            'b
                    End Select
                Next lXVal
            Next Y
        End If

        moPMapRadarTex.UnlockRectangle(0)
        Erase yDest

        If bForMiniMap = False Then
            'Now, draw that
            Dim lTempW As Int32 = moDevice.PresentationParameters.BackBufferWidth
            Dim lTempH As Int32 = moDevice.PresentationParameters.BackBufferHeight
            Dim lSmaller As Int32 = CInt(Math.Min(lTempW, lTempH) * 0.8)
            Dim lLarger As Int32 = Math.Max(lTempW, lTempH)
            Dim ptPos As Point
            ptPos.X = CInt((lTempW - lSmaller) / 2)
            ptPos.Y = CInt((lTempH - lSmaller) / 2)

            'Dim fMultX As Single = CSng(TerrainClass.Width / lSmaller)
            'Dim fMultY As Single = CSng(TerrainClass.Height / lSmaller)
            Dim fMultX As Single = CSng(255 / lSmaller)
            Dim fMultY As Single = CSng(255 / lSmaller)

            moDevice.RenderState.Lighting = False
            moDevice.RenderState.ZBufferWriteEnable = False
            Planet.EnsureSpritesValid(moDevice)
            Planet.moSprite.SetWorldViewLH(Matrix.Identity, moDevice.Transform.View)
            Planet.moSprite.Begin(SpriteFlags.None)
            Planet.moSprite.Draw2D(moPMapRadarTex, System.Drawing.Rectangle.FromLTRB(0, 0, 256, 256), System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lSmaller, ptPos.Y + lSmaller), System.Drawing.Point.Empty, 0, New Point(CInt(ptPos.X * fMultX), CInt(ptPos.Y * fMultY)), System.Drawing.Color.White)
            'Planet.moSprite.Draw2D(moPMapRadarTex, System.Drawing.Rectangle.FromLTRB(0, 0, TerrainClass.Width - 1, TerrainClass.Height - 1), System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lSmaller, ptPos.Y + lSmaller), System.Drawing.Point.Empty, 0, New Point(CInt(ptPos.X * fMultX), CInt(ptPos.Y * fMultY)), System.Drawing.Color.White)
            Planet.moSprite.End()
            moDevice.RenderState.ZBufferWriteEnable = True
            moDevice.RenderState.Lighting = True


            'Ok, now, render our selected...
            If alSelectedIdx Is Nothing = False AndAlso alSelectedIdx.Count > 0 AndAlso bRenderSelectedRanges = True Then
                Dim uVerts((alSelectedIdx.Count * 18) - 1) As CustomVertex.TransformedColoredTextured
                Dim lVertIdx As Int32 = 0
                Dim lPrimCnt As Int32 = 0

                Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
                Dim lPrevLighting As Boolean = moDevice.RenderState.Lighting
                Dim lPrevFog As Boolean = moDevice.RenderState.FogEnable
                Dim lPrevCullMode As Cull = moDevice.RenderState.CullMode

                With moDevice.RenderState
                    .Lighting = False
                    .ZBufferEnable = False
                    .FogEnable = False
                    .CullMode = Cull.None
                End With
                Dim lmin As TextureFilter = moDevice.SamplerState(0).MinFilter
                Dim lmip As TextureFilter = moDevice.SamplerState(0).MipFilter
                moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, TextureFilter.None)
                moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, TextureFilter.None)


                Try

                    Dim fScaler As Single = lSmaller / 256.0F
                    Dim lDetRng As Int32 = gl_FINAL_GRID_SQUARE_SIZE * 4
                    Dim rcSrc As Rectangle = New Rectangle(0, 0, 256, 256)

                    'Ok, we made the map 0 to 255
                    ' where 0 = (-cellspacing * width)
                    ' and 255 = (cellspacing * width)

                    For X = 0 To alSelectedIdx.Count - 1
                        Dim lIdx As Int32 = CInt(alSelectedIdx(X))
                        Dim oEntity As BaseEntity = oEnvir.oEntity(lIdx)
                        If oEntity Is Nothing = False AndAlso oEntity.oUnitDef Is Nothing = False Then

                            With oEntity
                                'Ok, first quad, optimum range
                                Dim fHlfVal As Single = 0 '.oUnitDef.FOW_OptRadarRange '* fScaler
                                Dim lClrVal As Int32 = System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb

                                lLocX = CInt(((.LocX + lHalfExtent) / lFullExtent) * lSmaller)
                                lLocZ = CInt(((.LocZ + lHalfExtent) / lFullExtent) * lSmaller)

                                Dim lFinalX As Int32 = (lSmaller - lLocX) + ptPos.X
                                lLocX = lFinalX
                                lLocZ += ptPos.Y

                                'Now, determine our fHlfVal
                                fHlfVal = CSng((.oUnitDef.FOW_OptRadarRange / lFullExtent) * lSmaller)

                                'triangle 1
                                uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 0, 0) : lVertIdx += 1
                                uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 0, 1) : lVertIdx += 1
                                uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 1, 0) : lVertIdx += 1
                                lPrimCnt += 1
                                'triangle 2
                                uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 0, 1) : lVertIdx += 1
                                uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 1, 0) : lVertIdx += 1
                                uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 1, 1) : lVertIdx += 1
                                lPrimCnt += 1

                                'Ok, second quad, maximum range
                                If .oUnitDef.FOW_MaxRadarRange > (.oMesh.RangeOffset * lDetRng) + 0.05F Then
                                    fHlfVal = CSng((.oUnitDef.FOW_MaxRadarRange / lFullExtent) * lSmaller)
                                    lClrVal = System.Drawing.Color.FromArgb(255, 128, 128, 128).ToArgb
                                    'triangle 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 0, 0) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 0, 1) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 1, 0) : lVertIdx += 1
                                    lPrimCnt += 1
                                    'triangle 2
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 0, 1) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 1, 0) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 1, 1) : lVertIdx += 1
                                    lPrimCnt += 1
                                End If

                                'Finally, third quad, weapon range...
                                If .lMaxWpnRngValue = -1 Then
                                    .lMaxWpnRngValue = 0

                                    'Request the maximum Wpn Range value here...

                                    Dim yMsg(7) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestMaxWpnRng).CopyTo(yMsg, 0)
                                    .GetGUIDAsString.CopyTo(yMsg, 2)
                                    goUILib.SendMsgToPrimary(yMsg)
                                End If

                                If .lMaxWpnRngValue > 0 Then
                                    fHlfVal = CSng((.lMaxWpnRngValue / lFullExtent) * lSmaller) '.lMaxWpnRngValue '* gl_FINAL_GRID_SQUARE_SIZE
                                    lClrVal = System.Drawing.Color.FromArgb(255, 255, 0, 0).ToArgb
                                    'triangle 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 0, 0) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 0, 1) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 1, 0) : lVertIdx += 1
                                    lPrimCnt += 1
                                    'triangle 2
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX - fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 0, 1) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ - fHlfVal, 0, 1, lClrVal, 1, 0) : lVertIdx += 1
                                    uVerts(lVertIdx) = New CustomVertex.TransformedColoredTextured(lLocX + fHlfVal, lLocZ + fHlfVal, 0, 1, lClrVal, 1, 1) : lVertIdx += 1
                                    lPrimCnt += 1
                                End If
                            End With
                        End If
                    Next X
                Catch
                End Try


                moDevice.VertexFormat = CustomVertex.TransformedColoredTextured.Format
                moDevice.SetTexture(0, goResMgr.GetTexture("visring.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "ob.pak"))
                moDevice.Transform.World = Matrix.Identity
                moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, lPrimCnt, uVerts)

                With moDevice.RenderState
                    .Lighting = lPrevLighting  'True
                    .ZBufferEnable = lPrevZBuffer 'True
                    .FogEnable = lPrevFog 'bFog
                    .CullMode = lPrevCullMode 'Cull.CounterClockwise
                End With
                moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, lmin)
                moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, lmip)
            End If

        End If

        mfFlashVal = fTempFlashVal
    End Sub

    Private Sub RenderRadarObject(ByVal oObj As RenderObject)
        Try
            If goCurrentPlayer.ShowSpaceDots = True OrElse goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                moDevice.Transform.World = oObj.GetWorldMatrix
                moDevice.SetTexture(0, Nothing)
                moDevice.Material = matRadar
                oObj.oMesh.oShieldMesh.DrawSubset(0)
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Function RenderSysViewMap() As ArrayList
        If HasAliasedRights(AliasingRights.eViewUnitsAndFacilities) = False Then Return Nothing
        If goCurrentEnvir Is Nothing Then Return Nothing
        'If goCurrentEnvir.lEntityUB = -1 Then Return Nothing
        If gbDeviceLost = True OrElse gbPaused = True Then Return Nothing

        If moRadarBlip Is Nothing OrElse moRadarBlip.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moRadarBlip = goResMgr.GetTexture("Blip.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
            Device.IsUsingEventHandlers = True
        End If

        Dim lPointSize As Int32

        Dim alList As ArrayList = Nothing

        If glCurrentEnvirView = CurrentView.eSystemMapView2 Then
            alList = goCurrentEnvir.ReadySysView2Points(bRenderSelectedRanges)
            lPointSize = 48
        ElseIf glCurrentEnvirView = CurrentView.eSystemMapView1 Then
            goCurrentEnvir.ReadySysView1Points()
            lPointSize = 8
        Else : Return Nothing
        End If

        'If goCurrentEnvir.moPoints Is Nothing OrElse goCurrentEnvir.moPoints.GetUpperBound(0) = -1 Then Return Nothing

        With moDevice
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
                If .DeviceCaps.MaxPointSize > lPointSize Then
                    .RenderState.PointSize = lPointSize
                Else : .RenderState.PointSize = .DeviceCaps.MaxPointSize
                End If
                .RenderState.PointScaleA = 0
                .RenderState.PointScaleB = 0
                .RenderState.PointScaleC = 0.5F
            End If
            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.One
            .RenderState.AlphaBlendEnable = True

            .Transform.World = Matrix.Identity

            .RenderState.ZBufferWriteEnable = False

            .RenderState.Lighting = False

            .VertexFormat = Microsoft.DirectX.Direct3D.CustomVertex.PositionColoredTextured.Format
            .SetTexture(0, moRadarBlip)
            If goCurrentEnvir.moPoints Is Nothing = False AndAlso goCurrentEnvir.moPoints.GetUpperBound(0) > -1 Then
                .DrawUserPrimitives(PrimitiveType.PointList, goCurrentEnvir.moPoints.GetUpperBound(0) + 1, goCurrentEnvir.moPoints)
            End If

            If goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso goCurrentEnvir.oGeoObject Is Nothing = False AndAlso CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                Dim oSystem As SolarSystem = CType(goCurrentEnvir.oGeoObject, SolarSystem)

                If oSystem.WormholeUB <> -1 Then
                    Dim uVerts(oSystem.WormholeUB) As CustomVertex.PositionColoredTextured
                    Dim fDivisor As Single = 1.0F
                    Dim fR As Single = Rnd() * 128 + 120
                    Dim fG As Single = Rnd() * 128 + 120
                    Dim fB As Single = Rnd() * 128 + 120
                    Dim lClr As Int32 = System.Drawing.Color.FromArgb(255, CInt(fR), CInt(fG), CInt(fB)).ToArgb
                    If glCurrentEnvirView = CurrentView.eSystemMapView2 Then
                        fDivisor = 30.0F
                    ElseIf glCurrentEnvirView = CurrentView.eSystemMapView1 Then
                        fDivisor = 10000.0F
                    Else : Return Nothing
                    End If

                    For X As Int32 = 0 To oSystem.WormholeUB
                        With oSystem.moWormholes(X)
                            If .System1 Is Nothing = False AndAlso .System1.ObjectID = oSystem.ObjectID Then
                                uVerts(X).Position = New Vector3(.LocX1 / fDivisor, 0.0F, .LocY1 / fDivisor)
                            Else : uVerts(X).Position = New Vector3(.LocX2 / fDivisor, 0.0F, .LocY2 / fDivisor)
                            End If
                            uVerts(X).Color = lClr
                        End With
                    Next X

                    .DrawUserPrimitives(PrimitiveType.PointList, oSystem.WormholeUB + 1, uVerts)
                End If
            End If

            .RenderState.ZBufferWriteEnable = True
            .RenderState.Lighting = True

            .RenderState.SourceBlend = Blend.SourceAlpha ' Blend.One
            .RenderState.DestinationBlend = Blend.InvSourceAlpha ' Blend.Zero
            .RenderState.AlphaBlendEnable = True
            .RenderState.PointSpriteEnable = False

            'Ok, if our device was created with mixed...
            If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                'Set us up for software vertex processing as point sprites always work in software
                moDevice.SoftwareVertexProcessing = False
            End If

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
        End With
        Return alList
    End Function

    Private Sub RenderHitPointBarsAndWPInds(ByVal lCameraCellX As Int32, ByVal lCameraCellZ As Int32)
        Dim X As Int32

        If gbDeviceLost = True OrElse gbPaused = True Then Return

        Dim uWPVerts() As CustomVertex.PositionTextured = Nothing
        Dim lWPVertUB As Int32 = -1

        If msw_Utility Is Nothing Then msw_Utility = Stopwatch.StartNew
        If msw_Utility.IsRunning = False Then
            msw_Utility.Reset() : msw_Utility.Start()
        End If
        If msw_Utility.ElapsedMilliseconds > 1000 Then
            msw_Utility.Reset() : msw_Utility.Start()
        End If
        Dim fTime As Single = msw_Utility.ElapsedMilliseconds / 1000.0F  'should be 0 to 1 every second

        Dim fTu1 As Single = fTime
        Dim fTu2 As Single = fTime + 1.0F
        Dim fDist As Single

        'in the correct view?
        If glCurrentEnvirView = CurrentView.ePlanetView OrElse glCurrentEnvirView = CurrentView.eSystemView Then
            'environment valid?
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                'Are we looking at our environment??
                If (oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso glCurrentEnvirView = CurrentView.ePlanetView) OrElse _
                   (oEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso glCurrentEnvirView = CurrentView.eSystemView) Then

                    'Dim yMapWrapSituation As Byte = 0			'0 = no map wrap, 1 = left edge, 2 = right edge
                    'Dim lLocXMapWrapCheck As Int32 = 0

                    'If oEnvir.ObjTypeID = ObjectType.ePlanet Then
                    '	Dim lTmpMapWrapVal As Int32 = Math.Min((oEnvir.lMaxXPos * 2) \ 3, muSettings.EntityClipPlane)
                    '	If goCamera.mlCameraX < oEnvir.lMinXPos + lTmpMapWrapVal Then
                    '		yMapWrapSituation = 1
                    '		lLocXMapWrapCheck = oEnvir.lMaxXPos - lTmpMapWrapVal
                    '	ElseIf goCamera.mlCameraX > oEnvir.lMaxXPos - lTmpMapWrapVal Then
                    '		yMapWrapSituation = 2
                    '		lLocXMapWrapCheck = oEnvir.lMinXPos + lTmpMapWrapVal
                    '	End If
                    'End If

                    Dim bBegun As Boolean = False

                    'set up our hp sprite
                    Dim bRenderHPBars As Boolean = muSettings.RenderHPBars AndAlso (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255)
                    If bRenderHPBars = True Then
                        If moHPStat Is Nothing Then
                            Device.IsUsingEventHandlers = False
                            moHPStat = New Sprite(moDevice)
                            'AddHandler moHPStat.Disposing, AddressOf SpriteDispose
                            'AddHandler moHPStat.Lost, AddressOf SpriteLost
                            'AddHandler moHPStat.Reset, AddressOf SpriteLost
                            Device.IsUsingEventHandlers = True
                        End If
                    End If


                    Try

                        Dim lCurUB As Int32 = -1
                        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                        For X = 0 To lCurUB
                            Dim oEntity As BaseEntity = Nothing
                            If oEnvir.lEntityIdx(X) <> -1 Then
                                oEntity = oEnvir.oEntity(X)
                            End If
                            If oEntity Is Nothing = False AndAlso oEntity.bSelected = True Then
                                With oEntity
                                    'If Math.Abs(.CellLocX - lCameraCellX) < 2 AndAlso _
                                    '   Math.Abs(.CellLocZ - lCameraCellZ) < 2 Then

                                    If .bCulled = False AndAlso bRenderHPBars = True Then
                                        If .yVisibility = eVisibilityType.Visible Then
                                            If bBegun = False Then
                                                moHPStat.SetWorldViewLH(Matrix.Identity, moDevice.Transform.View)
                                                moHPStat.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace Or SpriteFlags.SortDepthBackToFront)
                                                bBegun = True
                                            End If

                                            .RefreshHPTexture(moDevice)
                                            'moHPStat.Draw(.HPTexture, System.Drawing.Rectangle.Empty, New Vector3(16, 16, 0), New Vector3(.fMapWrapLocX, .LocY + (.oMesh.YMidPoint * 2.5F) + 64, .LocZ), System.Drawing.Color.White)
                                            moHPStat.Draw(.HPTexture, System.Drawing.Rectangle.Empty, New Vector3(16, 16, 0), New Vector3(.LocX, .LocY + (.oMesh.YMidPoint * 2.5F) + 64, .LocZ), System.Drawing.Color.White)
                                        Else
                                            .bSelected = False 'TODO: do we want to deselect the unit?
                                        End If
                                    End If

                                    If .OwnerID = glPlayerID AndAlso ((.CurrentStatus And elUnitStatus.eUnitMoving) <> 0) Then
                                        fDist = Distance(CInt(.LocX), CInt(.LocZ), .DestX, .DestZ)
                                        fDist /= 100
                                        If fDist < 1 Then fDist = 1

                                        'Now, add it to the vert list
                                        'uWPVerts
                                        lWPVertUB += 12
                                        ReDim Preserve uWPVerts(lWPVertUB)

                                        Dim fDestY As Single = .mlTargetY ' .LocY
                                        'If (.oUnitDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                        If oEnvir Is Nothing = False AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet Then
                                            If oEnvir.oGeoObject Is Nothing = False Then
                                                fDestY = CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(.DestX, .DestZ, False)
                                                .DestY = CInt(fDestY)
                                            End If
                                        End If
                                        'End If

                                        Dim lDestX As Int32 = .DestX
                                        'If yMapWrapSituation <> 0 AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet Then
                                        '    'Ok, in a mapwrap situation
                                        '    If yMapWrapSituation = 1 Then
                                        '        'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
                                        '        If lDestX > lLocXMapWrapCheck Then
                                        '            lDestX -= oEnvir.lMapWrapAdjustX
                                        '        End If
                                        '    Else
                                        '        'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
                                        '        If lDestX < lLocXMapWrapCheck Then
                                        '            lDestX += oEnvir.lMapWrapAdjustX
                                        '        End If
                                        '    End If
                                        'End If

                                        '0 1 2
                                        uWPVerts(lWPVertUB - 11) = New CustomVertex.PositionTextured(.LocX - ml_WP_IND_SIZE, .LocY, .LocZ - ml_WP_IND_SIZE, fTu1 * fDist, 0)                '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 10) = New CustomVertex.PositionTextured(.LocX + ml_WP_IND_SIZE, .LocY, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)                '.fMapWrapLocX  
                                        uWPVerts(lWPVertUB - 9) = New CustomVertex.PositionTextured(lDestX - ml_WP_IND_SIZE, fDestY, .DestZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        '1 2 3
                                        uWPVerts(lWPVertUB - 8) = New CustomVertex.PositionTextured(.LocX + ml_WP_IND_SIZE, .LocY, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)                 '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 7) = New CustomVertex.PositionTextured(lDestX - ml_WP_IND_SIZE, fDestY, .DestZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        uWPVerts(lWPVertUB - 6) = New CustomVertex.PositionTextured(lDestX + ml_WP_IND_SIZE, fDestY, .DestZ + ml_WP_IND_SIZE, fTu2 * fDist, 1)

                                        '0 1 2
                                        uWPVerts(lWPVertUB - 5) = New CustomVertex.PositionTextured(.LocX, .LocY - ml_WP_IND_SIZE, .LocZ - ml_WP_IND_SIZE, fTu1 * fDist, 0)                 '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 4) = New CustomVertex.PositionTextured(.LocX, .LocY + ml_WP_IND_SIZE, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)                 '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 3) = New CustomVertex.PositionTextured(lDestX, fDestY - ml_WP_IND_SIZE, .DestZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        '1 2 3
                                        uWPVerts(lWPVertUB - 2) = New CustomVertex.PositionTextured(.LocX, .LocY + ml_WP_IND_SIZE, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)                 '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 1) = New CustomVertex.PositionTextured(lDestX, fDestY - ml_WP_IND_SIZE, .DestZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        uWPVerts(lWPVertUB) = New CustomVertex.PositionTextured(lDestX, fDestY + ml_WP_IND_SIZE, .DestZ + ml_WP_IND_SIZE, fTu2 * fDist, 1)
                                    End If

                                    If .OwnerID = glPlayerID AndAlso .RallyX <> Int32.MinValue AndAlso .RallyZ <> Int32.MinValue Then
                                        fDist = Distance(CInt(.LocX), CInt(.LocZ), .DestX, .DestZ)
                                        fDist /= 100
                                        If fDist < 1 Then fDist = 1

                                        'Now, add it to the vert list
                                        'uWPVerts
                                        lWPVertUB += 12
                                        ReDim Preserve uWPVerts(lWPVertUB)

                                        Dim fDestY As Single = .mlTargetY '.LocY
                                        'If (.oUnitDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                        If oEnvir Is Nothing = False AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet Then
                                            If oEnvir.oGeoObject Is Nothing = False Then
                                                fDestY = CType(oEnvir.oGeoObject, Planet).GetHeightAtPoint(.RallyX, .RallyZ, True)
                                                '.DestY = CInt(fDestY)
                                            End If
                                        End If
                                        'End If

                                        Dim lDestX As Int32 = .RallyX
                                        'If yMapWrapSituation <> 0 AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet Then
                                        '	'Ok, in a mapwrap situation
                                        '	If yMapWrapSituation = 1 Then
                                        '		'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
                                        '		If lDestX > lLocXMapWrapCheck Then
                                        '			lDestX -= oEnvir.lMapWrapAdjustX
                                        '		End If
                                        '	Else
                                        '		'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
                                        '		If lDestX < lLocXMapWrapCheck Then
                                        '			lDestX += oEnvir.lMapWrapAdjustX
                                        '		End If
                                        '	End If
                                        'End If

                                        '0 1 2
                                        uWPVerts(lWPVertUB - 11) = New CustomVertex.PositionTextured(.LocX - ml_WP_IND_SIZE, .LocY, .LocZ - ml_WP_IND_SIZE, fTu1 * fDist, 0)            '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 10) = New CustomVertex.PositionTextured(.LocX + ml_WP_IND_SIZE, .LocY, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)            '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 9) = New CustomVertex.PositionTextured(lDestX - ml_WP_IND_SIZE, fDestY, .RallyZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        '1 2 3
                                        uWPVerts(lWPVertUB - 8) = New CustomVertex.PositionTextured(.LocX + ml_WP_IND_SIZE, .LocY, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)             '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 7) = New CustomVertex.PositionTextured(lDestX - ml_WP_IND_SIZE, fDestY, .RallyZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        uWPVerts(lWPVertUB - 6) = New CustomVertex.PositionTextured(lDestX + ml_WP_IND_SIZE, fDestY, .RallyZ + ml_WP_IND_SIZE, fTu2 * fDist, 1)

                                        '0 1 2
                                        uWPVerts(lWPVertUB - 5) = New CustomVertex.PositionTextured(.LocX, .LocY - ml_WP_IND_SIZE, .LocZ - ml_WP_IND_SIZE, fTu1 * fDist, 0)             '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 4) = New CustomVertex.PositionTextured(.LocX, .LocY + ml_WP_IND_SIZE, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)             '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 3) = New CustomVertex.PositionTextured(lDestX, fDestY - ml_WP_IND_SIZE, .RallyZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        '1 2 3
                                        uWPVerts(lWPVertUB - 2) = New CustomVertex.PositionTextured(.LocX, .LocY + ml_WP_IND_SIZE, .LocZ + ml_WP_IND_SIZE, fTu1 * fDist, 1)             '.fMapWrapLocX
                                        uWPVerts(lWPVertUB - 1) = New CustomVertex.PositionTextured(lDestX, fDestY - ml_WP_IND_SIZE, .RallyZ - ml_WP_IND_SIZE, fTu2 * fDist, 0)
                                        uWPVerts(lWPVertUB) = New CustomVertex.PositionTextured(lDestX, fDestY + ml_WP_IND_SIZE, .RallyZ + ml_WP_IND_SIZE, fTu2 * fDist, 1)
                                    End If
                                End With

                            End If
                        Next
                    Catch
                    End Try

                    'end our hp sprite
                    If bBegun = True AndAlso bRenderHPBars = True Then moHPStat.End()

                    'Now, render our wp verts
                    If lWPVertUB <> -1 Then
                        If moWPInd Is Nothing OrElse moWPInd.Disposed = True Then moWPInd = goResMgr.GetTexture("WPInd.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
                        With moDevice
                            .VertexFormat = CustomVertex.PositionTextured.Format
                            .RenderState.Lighting = False
                            .RenderState.ZBufferWriteEnable = False
                            .RenderState.AlphaBlendEnable = True
                            .RenderState.SourceBlend = Blend.SourceAlpha
                            .RenderState.DestinationBlend = Blend.One
                            .RenderState.AlphaBlendEnable = True
                            .Material = matRadar
                            .SetTexture(0, moWPInd)
                            .Transform.World = Matrix.Identity
                            .RenderState.CullMode = Cull.None

                            Try
                                Dim lPrimCnt As Int32 = (lWPVertUB + 1) \ 3  'CInt((lWPVertUB + 1) / 3.0F)
                                If uWPVerts Is Nothing = False AndAlso lPrimCnt > 0 Then 'AndAlso uWPVerts.Length \ 3 >= lPrimCnt Then
                                    .DrawUserPrimitives(PrimitiveType.TriangleList, lPrimCnt, uWPVerts)
                                End If
                            Catch
                            End Try
                            .RenderState.Lighting = True
                            .RenderState.ZBufferWriteEnable = True
                            .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                            .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                            .RenderState.AlphaBlendEnable = True
                        End With

                        Erase uWPVerts
                    End If

                End If
            End If

        End If
    End Sub

    Private Sub SetRenderStates()
        With moDevice.RenderState
            .Lighting = True
            .ZBufferEnable = True
            .DitherEnable = muSettings.Dither
            .SpecularEnable = muSettings.SpecularEnabled
            .ShadeMode = ShadeMode.Gouraud
            .SourceBlend = Blend.SourceAlpha
            .DestinationBlend = Blend.InvSourceAlpha
            .AlphaBlendEnable = True
            .ZBufferFunction = Compare.LessEqual
        End With

        moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, TextureFilter.Linear)
        moDevice.SetSamplerState(0, SamplerStageStates.MagFilter, TextureFilter.Linear)
        moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, TextureFilter.Linear)

        moDevice.SetSamplerState(0, SamplerStageStates.AddressU, TextureAddress.Wrap)
        moDevice.SetSamplerState(0, SamplerStageStates.AddressV, TextureAddress.Wrap)

    End Sub

    Private Sub DoCaptureScreenshot()
        Dim oSurf As Surface
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        Dim ifmt As Microsoft.DirectX.Direct3D.ImageFileFormat
        If sFile.EndsWith("\") = False Then sFile = sFile & "\"
        sFile &= "Screenshots"
        If Exists(sFile) = False Then MkDir(sFile)
        If sFile.EndsWith("\") = False Then sFile &= "\"

        sFile &= "SS_" & Now.ToString("MM_dd_yyyy_HHmmss") '& ".bmp"
        Select Case muSettings.ScreenshotFormat
            Case 1
                ifmt = ImageFileFormat.Jpg
                sFile &= ".jpg"
            Case 2
                ifmt = ImageFileFormat.Tga
                sFile &= ".tga"
            Case 3
                ifmt = ImageFileFormat.Png
                sFile &= ".png"
            Case Else
                ifmt = ImageFileFormat.Bmp
                sFile &= ".bmp"
        End Select

        oSurf = moDevice.GetBackBuffer(0, 0, BackBufferType.Mono)
        SurfaceLoader.Save(sFile, iFmt, oSurf)
        oSurf.Dispose()
        oSurf = Nothing

        If goUILib Is Nothing = False Then
            goUILib.AddNotification("Screenshot saved!", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Public Sub CaptureScreenshot()
        'this tells the engine to save the next screenshot
        mbCaptureScreenshot = True
    End Sub

    Public Sub ForcefulCleanup()
        'kill our effects engines
        If moBoxMeshTargets Is Nothing = False Then moBoxMeshTargets.Dispose()
        moBoxMeshTargets = Nothing
        If moBoxMeshSelStar Is Nothing = False Then moBoxMeshSelStar.Dispose()
        moBoxMeshSelStar = Nothing

        If moBombMesh Is Nothing = False Then moBombMesh.Dispose()
        If moBombTex Is Nothing = False Then moBombTex.Dispose()
        moBombMesh = Nothing
        moBombTex = Nothing
        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        If moRadarBlip Is Nothing = False Then moRadarBlip.Dispose()
        moRadarBlip = Nothing
        'If moRadarMesh Is Nothing = False Then moRadarMesh.Dispose()
        'moRadarMesh = Nothing
        If moHPStat Is Nothing = False Then moHPStat.Dispose()
        moHPStat = Nothing
        If moWPInd Is Nothing = False Then moWPInd.Dispose()
        moWPInd = Nothing
        If moPMapRadarTex Is Nothing = False Then moPMapRadarTex.Dispose()
        moPMapRadarTex = Nothing
        If moPost Is Nothing = False Then moPost.DisposeMe()
        moPost = Nothing

        If goUILib Is Nothing = False Then goUILib.ReleaseInterfaceTextures()


        goWpnMgr = Nothing
        goShldMgr = Nothing
        goExplMgr = Nothing
        goPFXEngine32 = Nothing
        goEntityDeath = Nothing
        'goRewards = Nothing
    End Sub

    Protected Overrides Sub Finalize()
        ForcefulCleanup()
        MyBase.Finalize()
    End Sub

    Public Sub WriteRSToFile()
        Dim oFS As New IO.FileStream("C:\" & Now.ToString("MMddyyyyHHmmss") & ".txt", IO.FileMode.Create)
        Dim oWrite As New IO.StreamWriter(oFS)

        oWrite.Write(moDevice.SamplerState(0).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.SamplerState(1).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.SamplerState(2).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.SamplerState(3).ToString.Replace(vbLf, vbCrLf))
        oWrite.WriteLine()
        oWrite.Write(moDevice.RenderState.ToString.Replace(vbLf, vbCrLf))
        oWrite.WriteLine()
        oWrite.WriteLine("DepthStencilFormat: " & moDevice.PresentationParameters.AutoDepthStencilFormat.ToString)
        oWrite.WriteLine()
        oWrite.WriteLine("Depth Size: " & moDevice.DepthStencilSurface.Description.Width & "x" & moDevice.DepthStencilSurface.Description.Height & "... MST: " & moDevice.DepthStencilSurface.Description.MultiSampleType.ToString & "... Quality: " & moDevice.DepthStencilSurface.Description.MultiSampleQuality)

        oWrite.Write(moDevice.Lights(0).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(0).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(1).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(2).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.TextureState.TextureState(3).ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.VertexDeclaration.ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(moDevice.VertexFormat.ToString.Replace(vbLf, vbCrLf))

        oWrite.Write(vbCrLf)
        oWrite.Write("PixelShaderCaps" & vbCrLf)
        oWrite.Write(moDevice.DeviceCaps.PixelShaderCaps.ToString.Replace(vbLf, vbCrLf))
        oWrite.Write(vbCrLf & "VertexShaderCaps" & vbCrLf)
        oWrite.Write(moDevice.DeviceCaps.VertexShaderCaps.ToString.Replace(vbLf, vbCrLf))
        oWrite.Close()
        oFS.Close()
        oWrite = Nothing
        oFS = Nothing
    End Sub

    Private Sub moDevice_DeviceLost(ByVal sender As Object, ByVal e As System.EventArgs)
        'On Error Resume Next
        GFXEngine.gbDeviceLost = True

        Debug.Write("GFXEngine.moDevice_DeviceLost" & vbCrLf)
        'This happens when the device is lost. we need to release any default textures and any sprites...

        If goGalaxy Is Nothing = False Then
            Try
                For X As Int32 = 0 To goGalaxy.mlSystemUB
                    Try
                        Dim oSystem As SolarSystem = goGalaxy.moSystems(X)
                        If oSystem Is Nothing = False Then
                            For Y As Int32 = 0 To oSystem.PlanetUB
                                Try
                                    Dim oPlanet As Planet = oSystem.moPlanets(Y)
                                    If oPlanet Is Nothing = False Then
                                        oPlanet.ReleaseRenderTargets()
                                    End If
                                Catch
                                End Try
                            Next Y
                        End If
                    Catch
                    End Try
                Next X
            Catch
            End Try
        End If

        Try
            If moPost Is Nothing = False Then moPost.ReleaseTextures()
        Catch
        End Try
        moPost = Nothing
        Try
            If moModelShader Is Nothing = False Then moModelShader.DisposeMe()
        Catch
        End Try
        moModelShader = Nothing
        Try
            If moLoginScreen Is Nothing = False Then
                moLoginScreen.CleanUpResources()
                'LoginScreen.mySequenceEnded = 2
                'moLoginScreen = Nothing
            End If
        Catch
        End Try

        If moLogo Is Nothing = False Then moLogo.ReleaseSprite()
        Try
            If moHPStat Is Nothing = False Then moHPStat.Dispose()
            moHPStat = Nothing
            If ctlDiplomacy.moSprite Is Nothing = False Then ctlDiplomacy.moSprite.Dispose()
            ctlDiplomacy.moSprite = Nothing
        Catch
        End Try
        StarType.ReleaseSprite()
        If goGalaxy Is Nothing = False Then goGalaxy.ReleaseSprite()
        Try
            If DeathSequence.moSprite Is Nothing = False Then DeathSequence.moSprite.Dispose()
            DeathSequence.moSprite = Nothing
        Catch
        End Try
        AgentRenderer.ReleaseSprite()
        frmTournament.ReleaseSprite()
        Planet.ReleaseSprites()

        TerrainClass.ReleaseVisibleTexture()
        TerrainClass.bReloadVertexBuffer = True
        TerrainClass.ForceWaterCreation()
        TerrainClass.ResetShaders()
        BPFont.ClearAllFonts()
        If goResMgr Is Nothing = False Then goResMgr.ReleaseDefaultPoolTextures()
        If goUILib Is Nothing = False Then goUILib.ReleaseInterfaceTextures()

    End Sub

    Private Sub moDevice_DeviceReset(ByVal sender As Object, ByVal e As System.EventArgs)
        Debug.Write("GFXEngine.moDevice_DeviceReset" & vbCrLf)
        If goResMgr Is Nothing = False Then goResMgr.RecreateDefaultPoolTextures()
    End Sub

    Private Sub moDevice_Disposing(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO: Do any final cleanup here
        Debug.Write("GFXEngine.moDevice_Disposing" & vbCrLf)
    End Sub

    Private Sub RenderBuildGhost()
        Dim X As Int32
        Dim lTexMod As Int32

        Dim matWorld As Matrix
        Dim oMat As Material

        Dim uVerts(11) As CustomVertex.PositionColoredTextured

        Dim bIgnoreSlopeTest As Boolean = False

        matWorld = Matrix.Identity
        If glCurrentEnvirView = CurrentView.ePlanetView AndAlso (goCurrentEnvir.oGeoObject Is Nothing = False) Then

            Dim lHalfVal As Int32 = CInt(Math.Ceiling(goUILib.BuildGhost.ShieldXZRadius / 2.0F))
            Dim lMinHt As Int32 = Int32.MaxValue
            Dim lMaxHt As Int32 = Int32.MinValue
            Dim lCntrPt As Int32 = 0
            Dim lTotal As Int32 = 0

            With CType(goCurrentEnvir.oGeoObject, Planet)
                Dim LocX As Int32 = CInt(goUILib.vecBuildGhostLoc.X)
                Dim LocZ As Int32 = CInt(goUILib.vecBuildGhostLoc.Z)
                Dim lTemp As Int32 = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ - lHalfVal, False))
                lTotal += lTemp
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ - lHalfVal, False))
                lTotal += lTemp
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ + lHalfVal, False))
                lTotal += lTemp
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ + lHalfVal, False))
                lTotal += lTemp
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX, LocZ, False))
                lTotal += lTemp
                lCntrPt = lTemp
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
            End With
            Dim bValid As Boolean = lMaxHt - lMinHt < 200
            'goUILib.vecBuildGhostLoc.Y = CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goUILib.vecBuildGhostLoc.X, goUILib.vecBuildGhostLoc.Z)
            If bValid = False Then lMinHt = ((lTotal \ 5) + lCntrPt) \ 2 'lCntrPt

            Dim lExtent As Int32 = (CType(goCurrentEnvir.oGeoObject, Planet).GetExtent \ 2)
            Dim lCellSpace As Int32 = CType(goCurrentEnvir.oGeoObject, Planet).CellSpacing
            If goUILib.vecBuildGhostLoc.Z + (lHalfVal * 4) > (lExtent - lCellSpace) OrElse goUILib.vecBuildGhostLoc.Z - (lHalfVal * 4) < -lExtent Then bValid = False
            If goUILib.vecBuildGhostLoc.X + (lHalfVal * 4) > (lExtent - lCellSpace) OrElse goUILib.vecBuildGhostLoc.X - (lHalfVal * 4) < -lExtent Then bValid = False

            goUILib.vecBuildGhostLoc.Y = lMinHt + goUILib.BuildGhost.PlanetYAdjust
            'matWorld.Multiply(Matrix.RotationY(DegreeToRadian((goUILib.BuildGhostAngle - 900.0F) / 10.0F)))

            Dim bNaval As Boolean = False
            For X = 0 To glEntityDefUB
                If glEntityDefIdx(X) = goUILib.BuildGhostID AndAlso goEntityDefs(X).ObjTypeID = goUILib.BuildGhostTypeID Then
                    If goEntityDefs(X).ProductionTypeID = ProductionType.eMining Then
                        bIgnoreSlopeTest = True
                    End If
                    bNaval = (goEntityDefs(X).yChassisType And ChassisType.eNavalBased) <> 0
                    'bNaval = True
                    Exit For
                End If
            Next X
            If bNaval = True Then goUILib.vecBuildGhostLoc.Y = CType(goCurrentEnvir.oGeoObject, Planet).WaterHeight

            Dim vecNormal As Vector3 = CType(goCurrentEnvir.oGeoObject, Planet).GetTriangleNormal(goUILib.vecBuildGhostLoc.X, goUILib.vecBuildGhostLoc.Z)
            With oMat
                If (bIgnoreSlopeTest = False AndAlso (Math.Abs(vecNormal.X) > 0.4 OrElse Math.Abs(vecNormal.Z) > 0.4)) OrElse ((bNaval = True AndAlso (goUILib.vecBuildGhostLoc.Y > CType(goCurrentEnvir.oGeoObject, Planet).WaterHeight Or CType(goCurrentEnvir.oGeoObject, Planet).MapTypeID = PlanetType.eAcidic Or CType(goCurrentEnvir.oGeoObject, Planet).MapTypeID = PlanetType.eGeoPlastic)) OrElse (bNaval = False AndAlso goUILib.vecBuildGhostLoc.Y < CType(goCurrentEnvir.oGeoObject, Planet).WaterHeight)) OrElse (bValid = False AndAlso bIgnoreSlopeTest = False) Then
                    .Ambient = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                    .Diffuse = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                    .Emissive = System.Drawing.Color.FromArgb(64, 255, 0, 0)

                    If NewTutorialManager.TutorialOn = True Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildGhostNonBuildable, -1, -1, -1, "")
                    End If
                Else
                    'Ok, check for intersections with other facilities
                    X = CInt(goUILib.BuildGhost.ShieldXZRadius)
                    Dim rcFac As Rectangle = Rectangle.FromLTRB(CInt(goUILib.vecBuildGhostLoc.X) - X, CInt(goUILib.vecBuildGhostLoc.Z) - X, CInt(goUILib.vecBuildGhostLoc.X) + X, CInt(goUILib.vecBuildGhostLoc.Z) + X)
                    'Now, cycle thru to see if anyone is within that rect...
                    Dim bFound As Boolean = False
                    Dim lCurUB As Int32 = Math.Min(goCurrentEnvir.lEntityUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
                    For X = 0 To lCurUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).yVisibility = eVisibilityType.Visible AndAlso (goCurrentEnvir.oEntity(X).oMesh.bLandBased = True OrElse goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility) AndAlso goCurrentEnvir.oEntity(X).bSelected = False Then
                            Dim lShieldRad As Int32 = CInt(goCurrentEnvir.oEntity(X).oMesh.ShieldXZRadius)
                            Dim rcTemp As Rectangle = Rectangle.FromLTRB(CInt(goCurrentEnvir.oEntity(X).LocX - lShieldRad), _
                            CInt(goCurrentEnvir.oEntity(X).LocZ - lShieldRad), CInt(goCurrentEnvir.oEntity(X).LocX + lShieldRad), _
                            CInt(goCurrentEnvir.oEntity(X).LocZ + lShieldRad))

                            If rcFac.IntersectsWith(rcTemp) = True Then
                                bFound = True
                                Exit For
                            End If
                        End If
                    Next X

                    If bFound = False Then

                        Dim clrGhostClr As System.Drawing.Color = System.Drawing.Color.FromArgb(64, 255, 255, 255)
                        Select Case CType(CType(goCurrentEnvir.oGeoObject, Planet).MapTypeID, PlanetType)
                            Case PlanetType.eAcidic
                                clrGhostClr = muSettings.AcidBuildGhost
                            Case PlanetType.eAdaptable
                                clrGhostClr = muSettings.AdaptableBuildGhost
                            Case PlanetType.eBarren
                                clrGhostClr = muSettings.BarrenBuildGhost
                            Case PlanetType.eDesert
                                clrGhostClr = muSettings.DesertBuildGhost
                            Case PlanetType.eGeoPlastic
                                clrGhostClr = muSettings.LavaBuildGhost
                            Case PlanetType.eTerran
                                clrGhostClr = muSettings.TerranBuildGhost
                            Case PlanetType.eTundra
                                clrGhostClr = muSettings.IceBuildGhost
                            Case PlanetType.eWaterWorld
                                clrGhostClr = muSettings.WaterworldBuildGhost
                        End Select

                        .Ambient = clrGhostClr
                        .Diffuse = clrGhostClr
                        .Emissive = clrGhostClr
                    Else
                        .Ambient = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                        .Diffuse = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                        .Emissive = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                    End If
                End If
            End With
        Else
            'Now, determine intersection distance
            Dim lMinDist As Int32 = 15000   '15k by default with a min of 6k
            If goCurrentPlayer.StationPlacementCloserToPlanet = True Then lMinDist = 6000

            Dim bInvalid As Boolean = False

            For X = 0 To glEntityDefUB
                If glEntityDefIdx(X) = goUILib.BuildGhostID AndAlso goEntityDefs(X).ObjTypeID = goUILib.BuildGhostTypeID Then
                    If (goEntityDefs(X).ModelID And 255) = 148 Then
                        If goCurrentEnvir.PositionOnPlanetRing(goUILib.vecBuildGhostLoc) = False Then
                            bInvalid = True
                        End If
                        lMinDist = 0
                    End If
                    Exit For
                End If
            Next X

            'Now, find any planets/moons nearby
            With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                For X = 0 To .PlanetUB
                    If Distance(CInt(goUILib.vecBuildGhostLoc.X), CInt(goUILib.vecBuildGhostLoc.Z), .moPlanets(X).LocX, .moPlanets(X).LocZ) < lMinDist Then
                        bInvalid = True
                        Exit For
                    End If
                Next X
                For X = 0 To .WormholeUB

                    Dim lTmpX As Int32
                    Dim lTmpZ As Int32
                    If .moWormholes(X).System1.ObjectID = .ObjectID Then
                        lTmpX = .moWormholes(X).LocX1 : lTmpZ = .moWormholes(X).LocY1
                    Else : lTmpX = .moWormholes(X).LocX2 : lTmpZ = .moWormholes(X).LocY2
                    End If

                    If Distance(CInt(goUILib.vecBuildGhostLoc.X), CInt(goUILib.vecBuildGhostLoc.Z), lTmpX, lTmpZ) < 1000 Then
                        bInvalid = True
                        Exit For
                    End If
                Next X
            End With

            'Now, check for intersection with other facilities
            'Ok, check for intersections with other facilities
            X = CInt(goUILib.BuildGhost.ShieldXZRadius)
            Dim rcFac As Rectangle = Rectangle.FromLTRB(CInt(goUILib.vecBuildGhostLoc.X) - X, CInt(goUILib.vecBuildGhostLoc.Z) - X, CInt(goUILib.vecBuildGhostLoc.X) + X, CInt(goUILib.vecBuildGhostLoc.Z) + X)
            'Now, cycle thru to see if anyone is within that rect...
            Dim lCurUB As Int32 = Math.Min(goCurrentEnvir.lEntityUB, goCurrentEnvir.lEntityIdx.GetUpperBound(0))
            For X = 0 To lCurUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).yVisibility = eVisibilityType.Visible AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility Then
                    Dim lShieldRad As Int32 = CInt(goCurrentEnvir.oEntity(X).oMesh.ShieldXZRadius)
                    Dim rcTemp As Rectangle = Rectangle.FromLTRB(CInt(goCurrentEnvir.oEntity(X).LocX - lShieldRad), _
                    CInt(goCurrentEnvir.oEntity(X).LocZ - lShieldRad), CInt(goCurrentEnvir.oEntity(X).LocX + lShieldRad), _
                    CInt(goCurrentEnvir.oEntity(X).LocZ + lShieldRad))

                    If rcFac.IntersectsWith(rcTemp) = True Then
                        bInvalid = True
                        Exit For
                    End If
                End If
            Next X

            'Ok, set up our material
            With oMat
                If bInvalid = True Then
                    .Ambient = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                    .Diffuse = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                    .Emissive = System.Drawing.Color.FromArgb(64, 255, 0, 0)
                Else
                    .Ambient = System.Drawing.Color.FromArgb(64, 255, 255, 255)
                    .Diffuse = System.Drawing.Color.FromArgb(64, 255, 255, 255)
                    .Emissive = System.Drawing.Color.FromArgb(64, 255, 255, 255)
                End If
            End With
        End If

        'now create out directional plane to show which way the build ghost faces
        Dim vecLoc As Vector3 = goUILib.vecBuildGhostLoc

        With vecLoc
            Dim fSize As Single = 1500
            Dim fX As Single = .X - fSize
            Dim fZ As Single = .Z - fSize
            Dim lClr As Int32
            RotatePoint(.X, .Z, fX, fZ, CSng(goUILib.BuildGhostAngle / 10))
            'triangle 1 rear
            lClr = Color.FromArgb(255, 0, 255, 0).ToArgb
            uVerts(0) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 1, 0)
            uVerts(1) = New CustomVertex.PositionColoredTextured(.X, .Y, .Z, lClr, 0.5, 0.5)
            RotatePoint(.X, .Z, fX, fZ, 90.0F)
            uVerts(2) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 0, 0)
            'triangle 2 left
            lClr = Color.FromArgb(255, 255, 0, 255).ToArgb
            uVerts(3) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 0, 0)
            uVerts(4) = New CustomVertex.PositionColoredTextured(.X, .Y, .Z, lClr, 0.5, 0.5)
            RotatePoint(.X, .Z, fX, fZ, 90.0F)
            uVerts(5) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 0, 1)
            'triangle 3 front
            lClr = Color.FromArgb(255, 0, 0, 255).ToArgb
            uVerts(6) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 0, 1)
            uVerts(7) = New CustomVertex.PositionColoredTextured(.X, .Y, .Z, lClr, 0.5, 0.5)
            RotatePoint(.X, .Z, fX, fZ, 90.0F)
            uVerts(8) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 1, 1)
            'triangle 4 right
            lClr = Color.FromArgb(255, 255, 0, 0).ToArgb
            uVerts(9) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 1, 1)
            uVerts(10) = New CustomVertex.PositionColoredTextured(.X, .Y, .Z, lClr, 0.5, 0.5)
            RotatePoint(.X, .Z, fX, fZ, 90.0F)
            uVerts(11) = New CustomVertex.PositionColoredTextured(fX, .Y, fZ, lClr, 1, 0)

            BPFont.DrawTextAtPoint(goUILib.oBuildGhostFont, CStr(goUILib.BuildGhostAngle / 10), frmMain.mlMouseX, frmMain.mlMouseY + 20, muSettings.InterfaceBorderColor)
        End With

        'Dim yMapWrapSituation As Byte
        'Dim lLocXMapWrapCheck As Int32
        'If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
        '	Dim lTmpMapWrapVal As Int32 = Math.Min((goCurrentEnvir.lMaxXPos * 2) \ 3, muSettings.EntityClipPlane)
        '	If goCamera.mlCameraX < goCurrentEnvir.lMinXPos + lTmpMapWrapVal Then
        '		yMapWrapSituation = 1
        '		lLocXMapWrapCheck = goCurrentEnvir.lMaxXPos - lTmpMapWrapVal
        '	ElseIf goCamera.mlCameraX > goCurrentEnvir.lMaxXPos - lTmpMapWrapVal Then
        '		yMapWrapSituation = 2
        '		lLocXMapWrapCheck = goCurrentEnvir.lMinXPos + lTmpMapWrapVal
        '	End If
        'End If

        Dim vecRenderLoc As Vector3 = goUILib.vecBuildGhostLoc
        'If yMapWrapSituation <> 0 AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
        '	'Ok, in a mapwrap situation
        '	If yMapWrapSituation = 1 Then
        '		'left edge.. in left edge mapwrap situations, entities on the right edge shift -= mapadjx
        '		If vecRenderLoc.X > lLocXMapWrapCheck Then vecRenderLoc.X -= goCurrentEnvir.lMapWrapAdjustX
        '	Else
        '		'right edge... in a right edge mapwrap situation, entites on the left edge shift +=
        '		If vecRenderLoc.X < lLocXMapWrapCheck Then vecRenderLoc.X += goCurrentEnvir.lMapWrapAdjustX
        '	End If
        'End If

        matWorld.Multiply(Matrix.RotationY(DegreeToRadian((goUILib.BuildGhostAngle - 900.0F) / 10.0F)))
        matWorld.Multiply(Matrix.Translation(vecRenderLoc))
        moDevice.Transform.World = matWorld

        Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
        Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
        Dim lPrevAlplaBlnd As Boolean = moDevice.RenderState.AlphaBlendEnable

        Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
        Dim lPeevLighting As Boolean = moDevice.RenderState.Lighting
        Dim lPrevCullMode As Cull = moDevice.RenderState.CullMode
        Dim lPrevFog As Boolean = moDevice.RenderState.FogEnable

        'Ensure proper renderstates...
        With moDevice.RenderState
            .SourceBlend = Blend.SourceAlpha
            .DestinationBlend = Blend.One
            .AlphaBlendEnable = True
        End With
        moDevice.Material = oMat

        'render build ghost
        lTexMod = 0
        For X = 0 To goUILib.BuildGhost.NumOfMaterials - 1
            moDevice.SetTexture(0, goUILib.BuildGhost.Textures((X * 4) + lTexMod))
            goUILib.BuildGhost.oMesh.DrawSubset(X)
        Next X

        'Reset our renderstates and change to rendering directional plane
        With moDevice.RenderState
            .SourceBlend = lPrevSrcBlnd
            .DestinationBlend = lPrevDestBlnd
            .AlphaBlendEnable = lPrevAlplaBlnd

            .Lighting = False
            .ZBufferEnable = False
            .FogEnable = False
            .CullMode = Cull.None
        End With

        'render directional plane
        moDevice.VertexFormat = CustomVertex.PositionColoredTextured.Format
        moDevice.SetTexture(0, goResMgr.GetTexture("visring.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "ob.pak"))
        moDevice.Transform.World = Matrix.Identity
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, 4, uVerts)

        'reset our true renderstates
        With moDevice.RenderState
            .Lighting = lPeevLighting 'True
            .ZBufferEnable = lPrevZBuffer 'True
            .FogEnable = lPrevFog
            .CullMode = lPrevCullMode 'Cull.CounterClockwise
        End With

    End Sub

    Private Sub RenderSystemMiniMap()
        If goUILib Is Nothing = False AndAlso goUILib.yRenderUI = 0 Then Return
        If muSettings.ShowMiniMap = False AndAlso (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255) Then Return
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        If moDevice Is Nothing Then Return
        If goCamera Is Nothing Then Return

        If mlLastSystemMinimapUpdate = Int32.MinValue OrElse glCurrentCycle - mlLastSystemMinimapUpdate > 2 Then
            mlLastSystemMinimapUpdate = glCurrentCycle

            Dim oOrigSurf As Surface = moDevice.GetRenderTarget(0)
            Dim X As Int32
            Dim matWorld As Matrix
            Dim yRel As Byte = Byte.MaxValue
            Dim oTmpRel As PlayerRel = Nothing

            If mbSysMiniMapInit = False OrElse moSystemMinimapTexture Is Nothing OrElse moSystemMinimapTexture.Disposed = True Then
                ReDim moSysMiniMapMats(4)
                With moSysMiniMapMats(0)  'Neutral
                    .Ambient = System.Drawing.Color.DarkGray
                    .Diffuse = System.Drawing.Color.FromArgb(192, 64, 64, 64)
                    .Emissive = .Diffuse
                    .Specular = System.Drawing.Color.Black
                    .SpecularSharpness = 0
                End With
                With moSysMiniMapMats(1)  'Mine
                    .Ambient = System.Drawing.Color.DarkGray
                    .Diffuse = System.Drawing.Color.FromArgb(192, 8, 255, 8)
                    .Emissive = .Diffuse
                    .Specular = System.Drawing.Color.Black
                    .SpecularSharpness = 0
                End With
                With moSysMiniMapMats(2)  'Ally
                    .Ambient = System.Drawing.Color.DarkGray
                    .Diffuse = System.Drawing.Color.FromArgb(192, 8, 255, 255)
                    .Emissive = .Diffuse
                    .Specular = System.Drawing.Color.Black
                    .SpecularSharpness = 0
                End With
                With moSysMiniMapMats(3)  'Enemy
                    .Ambient = System.Drawing.Color.DarkGray
                    .Diffuse = System.Drawing.Color.FromArgb(192, 255, 8, 8)
                    .Emissive = .Diffuse
                    .Specular = System.Drawing.Color.Black
                    .SpecularSharpness = 0
                End With
                With moSysMiniMapMats(4)    'guild
                    .Ambient = System.Drawing.Color.DarkGray
                    .Diffuse = System.Drawing.Color.FromArgb(192, 255, 8, 255)
                    .Emissive = .Diffuse
                    .Specular = System.Drawing.Color.Black
                    .SpecularSharpness = 0
                End With

                Device.IsUsingEventHandlers = False
                moSystemMinimapTexture = New Texture(moDevice, 128, 128, 1, Usage.RenderTarget, Format.A8R8G8B8, Pool.Default)
                Device.IsUsingEventHandlers = True

                mbSysMiniMapInit = True
            End If

            moDevice.SetRenderTarget(0, moSystemMinimapTexture.GetSurfaceLevel(0))
            moDevice.Clear(ClearFlags.Target, System.Drawing.Color.Black, 1, 0)
            moDevice.SetTexture(0, Nothing)

            'Now... set up our matrices
            goCamera.SetupSystemMinimapMatrices(moDevice)

            moDevice.RenderState.ZBufferWriteEnable = False
            moDevice.RenderState.SourceBlend = Blend.SourceAlpha
            moDevice.RenderState.DestinationBlend = Blend.One
            moDevice.RenderState.BlendOperation = BlendOperation.Max
            moDevice.RenderState.AlphaBlendEnable = True

            glMinimapItemsRendered = 0
            'Now, render...
            Dim bCheckGuild As Boolean = False
            Dim oGuild As Guild = Nothing
            If goCurrentPlayer Is Nothing = False Then
                oGuild = goCurrentPlayer.oGuild
                If oGuild Is Nothing = False Then
                    bCheckGuild = (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0
                End If
            End If

            Try
                Dim oEnvir As BaseEnvironment = goCurrentEnvir
                If oEnvir Is Nothing = False Then
                    For X = 0 To oEnvir.lEntityUB
                        If oEnvir.lEntityIdx(X) <> -1 Then
                            Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                            If oEntity Is Nothing Then Continue For
                            With oEntity
                                If .LocX > goCamera.mlCameraX - 10000 AndAlso .LocX < goCamera.mlCameraX + 10000 Then
                                    If .LocZ > goCamera.mlCameraZ - 10000 AndAlso .LocZ < goCamera.mlCameraZ + 10000 Then
                                        glMinimapItemsRendered += 1

                                        matWorld = .GetWorldMatrix()
                                        'matWorld = Matrix.Identity
                                        'matWorld.Multiply(Matrix.Translation(.LocX, 0, .LocZ))
                                        moDevice.Transform.World = matWorld

                                        If oEnvir.oEntity(X).OwnerID = glPlayerID Then
                                            moDevice.Material = moSysMiniMapMats(1)
                                        ElseIf (bCheckGuild = True AndAlso oGuild.MemberInGuild(.OwnerID) = True) Then
                                            moDevice.Material = moSysMiniMapMats(4)
                                        Else
                                            Dim yVisible As Byte = oEnvir.UnitInRadarRange(CInt(.LocX), CInt(.LocZ), .oMesh.RangeOffset)
                                            'MSC: we only care about the mesh portion of the modelid
                                            If (.oMesh.lModelID And 255) = 24 OrElse (.yProductionType = ProductionType.eMining AndAlso .ObjTypeID = ObjectType.eFacility) Then
                                                yVisible = 2 ' eVisibilityType.Visible
                                            End If
                                            If yVisible = eVisibilityType.Visible Then
                                                If .yRelID = Byte.MaxValue OrElse .yRelID = 0 Then
                                                    yRel = goCurrentPlayer.GetPlayerRelScore(.OwnerID)
                                                    .yRelID = yRel
                                                Else : yRel = .yRelID
                                                End If

                                                If yRel <= elRelTypes.eWar Then
                                                    moDevice.Material = moSysMiniMapMats(3)
                                                ElseIf yRel <= elRelTypes.ePeace Then       'elRelTypes.eNeutral
                                                    moDevice.Material = moSysMiniMapMats(0)
                                                Else : moDevice.Material = moSysMiniMapMats(2)
                                                End If
                                            ElseIf yVisible = eVisibilityType.InMaxRange Then
                                                moDevice.Material = moSysMiniMapMats(0)
                                            Else : Continue For
                                            End If

                                        End If
                                        .oMesh.oShieldMesh.DrawSubset(0)
                                    End If
                                End If
                            End With
                        End If
                    Next X
                End If
            Catch
            End Try

            moDevice.RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
            moDevice.RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
            moDevice.RenderState.BlendOperation = BlendOperation.Add
            moDevice.RenderState.AlphaBlendEnable = True


            'Render our camera view...
            Dim uVerts(2) As CustomVertex.PositionColored
            Dim vecTemp As Vector3 = New Vector3(goCamera.mlCameraAtX - goCamera.mlCameraX, 0, goCamera.mlCameraAtZ - goCamera.mlCameraZ)
            vecTemp.Normalize()
            vecTemp.Multiply(2000)

            uVerts(0) = New CustomVertex.PositionColored(goCamera.mlCameraX, 0, goCamera.mlCameraZ, System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb)

            Dim fX As Single = vecTemp.X
            Dim fZ As Single = vecTemp.Z
            RotatePoint(0, 0, fX, fZ, -22.5F)
            uVerts(1) = New CustomVertex.PositionColored(goCamera.mlCameraX + fX, 0, goCamera.mlCameraZ + fZ, System.Drawing.Color.FromArgb(64, 255, 255, 255).ToArgb)

            fX = vecTemp.X
            fZ = vecTemp.Z
            RotatePoint(0, 0, fX, fZ, 22.5F)
            uVerts(2) = New CustomVertex.PositionColored(goCamera.mlCameraX + fX, 0, goCamera.mlCameraZ + fZ, System.Drawing.Color.FromArgb(64, 255, 255, 255).ToArgb)

            moDevice.Transform.World = Matrix.Identity
            moDevice.VertexFormat = CustomVertex.PositionColored.Format
            moDevice.RenderState.Lighting = False
            moDevice.RenderState.CullMode = Cull.None
            moDevice.SetTexture(0, Nothing)
            moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, 1, uVerts)
            moDevice.RenderState.CullMode = Cull.CounterClockwise
            moDevice.RenderState.Lighting = True


            moDevice.RenderState.ZBufferWriteEnable = True


            'Now, restore our original surface
            moDevice.SetRenderTarget(0, oOrigSurf)

            oOrigSurf.Dispose()
            oOrigSurf = Nothing

            'Now, reset the view...
            goCamera.SetupMatrices(moDevice, glCurrentEnvirView)
            'and the render states
            SetRenderStates()
        End If

        ''Now, draw it to screen
        'If moHPStat Is Nothing Then
        '    Device.IsUsingEventHandlers = False
        '    moHPStat = New Sprite(moDevice)
        '    'AddHandler moHPStat.Disposing, AddressOf SpriteDispose
        '    'AddHandler moHPStat.Lost, AddressOf SpriteLost
        '    'AddHandler moHPStat.Reset, AddressOf SpriteLost
        '    Device.IsUsingEventHandlers = True
        'End If
        'With moHPStat
        '    .Begin(SpriteFlags.None)
        '    .Draw2D(moSystemMinimapTexture, New Point(0, 0), 0, New Point(0, 0), System.Drawing.Color.White)
        '    .End()
        'End With
        Dim bFog As Boolean = False
        With moDevice.RenderState
            .Lighting = False
            bFog = .FogEnable
            .FogEnable = False
        End With
        BPSprite.Draw2DOnce(moDevice, moSystemMinimapTexture, New Rectangle(0, 0, 128, 128), New Rectangle(0, 0, 128, 128), System.Drawing.Color.White, 128, 128)
        With moDevice.RenderState
            .Lighting = True
            .FogEnable = bFog
        End With
    End Sub

    Private Sub RenderPlanetMiniMap()
        If muSettings.ShowMiniMap = False AndAlso (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255) Then Return

        glMinimapItemsRendered = 0

        'Now, Call RenderPlanetMapBlips passing true in... this causes an update to the moPMapRadarTex texture but no render
        If glCurrentCycle Mod 10 = 0 Then
            RenderPlanetMapBlips(True)
        Else
            If muSettings.gbPlanetMapBlinkUnits = True Then
                If msw_Utility Is Nothing Then msw_Utility = Stopwatch.StartNew
                If msw_Utility.IsRunning = False Then
                    msw_Utility.Reset() : msw_Utility.Start()
                End If

                If msw_Utility.ElapsedMilliseconds > 30 Then
                    mfFlashVal -= 0.1F
                    If mfFlashVal < 0 Then mfFlashVal = 1.0F
                    msw_Utility.Reset()
                    msw_Utility.Start()
                End If
            Else
                mfFlashVal = 0.9F
            End If
        End If

        If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
            Planet.GetTutorialPlanet().RenderMiniMap(moPMapRadarTex, CByte(255 * mfFlashVal))
        Else
            'Now, tell our planet to render its minimap
            Dim lSysIdx As Int32 = goGalaxy.CurrentSystemIdx
            If lSysIdx = -1 Then Return
            Dim lPlanetIdx As Int32 = goGalaxy.moSystems(lSysIdx).CurrentPlanetIdx
            goGalaxy.moSystems(lSysIdx).moPlanets(lPlanetIdx).RenderMiniMap(moPMapRadarTex, CByte(255 * mfFlashVal))
        End If
    End Sub

    Private Sub RenderGalaxyFleetMovements()

        If goCurrentPlayer.mlUnitGroupUB = -1 Then Return

        Dim oMat As Material
        With oMat
            .Ambient = Color.Black
            .Diffuse = Color.Black
            .Emissive = Color.Teal
            .Specular = Color.Black
        End With

        Dim lLineColor As Int32 = System.Drawing.Color.FromArgb(192, 255, 255, 255).ToArgb
        Dim lIndColor As Int32 = System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb

        'set our states, materials, textures
        Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
        Dim lPrevLighting As Boolean = moDevice.RenderState.Lighting
        Dim lPrevCullMode As Cull = moDevice.RenderState.CullMode
        Dim lPrevVertexFormat As VertexFormats = moDevice.VertexFormat
        Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
        Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
        Dim lPrevAlplaBlnd As Boolean = moDevice.RenderState.AlphaBlendEnable


        With moDevice
            .SetTexture(0, Nothing)
            .Material = oMat
            .RenderState.CullMode = Cull.None
            '.RenderState.ZBufferWriteEnable = False
            .RenderState.ZBufferEnable = False
            .Transform.World = Matrix.Identity
            .VertexFormat = CustomVertex.PositionColored.Format

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.One
            .RenderState.AlphaBlendEnable = True
        End With

        'moDevice.RenderState.Lighting = False
        Dim fCameraAngle As Single = goCamera.GetCameraAngleDegrees()
        Dim fSize As Single = 1.5F
        Dim fSpriteOffsetY As Single = 8.0F '16.0F

        'Go ahead and do our rotation of the points for Billboard effect
        Dim fIndX As Single = fSize * 5
        Dim fIndZ As Single = 0
        RotatePoint(0, 0, fIndX, fIndZ, fCameraAngle - 90)

        For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
            If goCurrentPlayer.mlUnitGroupIdx(X) <> -1 AndAlso goCurrentPlayer.moUnitGroups(X).lInterSystemOriginID <> -1 AndAlso goCurrentPlayer.moUnitGroups(X).lInterSystemTargetID <> -1 AndAlso goCurrentPlayer.moUnitGroups(X).iParentTypeID = ObjectType.eGalaxy Then
                With goCurrentPlayer.moUnitGroups(X)
                    'Ok... let's get the two points
                    Dim vecFrom As Vector3
                    Dim vecTo As Vector3
                    Dim lTargetID As Int32 = .lInterSystemTargetID
                    Dim lOriginID As Int32 = .lInterSystemOriginID

                    'vecFrom.X = .lOriginX
                    'vecFrom.Y = .lOriginY
                    'vecFrom.Z = .lOriginZ

                    For lSysIdx As Int32 = 0 To goGalaxy.mlSystemUB
                        With goGalaxy.moSystems(lSysIdx)
                            If .ObjectID = lTargetID Then
                                vecTo.X = .LocX : vecTo.Y = .LocY : vecTo.Z = .LocZ
                                lTargetID = -1
                            ElseIf .ObjectID = lOriginID Then
                                vecFrom.X = .LocX : vecFrom.Y = .LocY : vecFrom.Z = .LocZ
                                lOriginID = -1
                            End If
                        End With
                        If lTargetID = -1 AndAlso lOriginID = -1 Then Exit For
                    Next lSysIdx

                    'If vecFrom.X = vecTo.X AndAlso vecFrom.Y = vecTo.Y AndAlso vecFrom.Z = vecTo.Z Then vecFrom.X += 1

                    'Now... find out the mult
                    Dim fMult As Single = CSng(.lInterSystemMoveCyclesRemaining / .InterSystemTotalCycles)
                    If fMult = Single.PositiveInfinity Then fMult = 0
                    Dim vecDiff As Vector3 = Vector3.Subtract(vecFrom, vecTo)
                    vecDiff.Multiply(fMult)
                    vecDiff = vecTo + vecDiff ' vecDiff.Add(vecFrom)

                    'Now, draw the connecting line... set up our verts...
                    Dim uVerts1(3) As CustomVertex.PositionColored
                    Dim uVerts2(3) As CustomVertex.PositionColored

                    uVerts1(0) = New CustomVertex.PositionColored(vecFrom.X - fSize, vecFrom.Y + fSpriteOffsetY, vecFrom.Z - (fSize / 2.0F), lLineColor)
                    uVerts1(1) = New CustomVertex.PositionColored(vecFrom.X + fSize, vecFrom.Y + fSpriteOffsetY, vecFrom.Z - (fSize / 2.0F), lLineColor)
                    uVerts1(2) = New CustomVertex.PositionColored(vecTo.X - fSize, vecTo.Y + fSpriteOffsetY, vecTo.Z + (fSize / 2.0F), lLineColor)
                    uVerts1(3) = New CustomVertex.PositionColored(vecTo.X + fSize, vecTo.Y + fSpriteOffsetY, vecTo.Z - (fSize / 2.0F), lLineColor)

                    uVerts2(0) = New CustomVertex.PositionColored(vecFrom.X + (fSize / 2.0F), vecFrom.Y + fSize + fSpriteOffsetY, vecFrom.Z, lLineColor)
                    uVerts2(1) = New CustomVertex.PositionColored(vecFrom.X - (fSize / 2.0F), vecFrom.Y - fSize + fSpriteOffsetY, vecFrom.Z, lLineColor)
                    uVerts2(2) = New CustomVertex.PositionColored(vecTo.X + (fSize / 2.0F), vecTo.Y + fSize + fSpriteOffsetY, vecTo.Z, lLineColor)
                    uVerts2(3) = New CustomVertex.PositionColored(vecTo.X - (fSize / 2.0F), vecTo.Y - fSize + fSpriteOffsetY, vecTo.Z, lLineColor)

                    moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts1)
                    moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts2)

                    'Now, draw the indicator triangle
                    uVerts1(0) = New CustomVertex.PositionColored(vecDiff.X, vecDiff.Y + fSpriteOffsetY, vecDiff.Z, lIndColor)
                    uVerts1(1) = New CustomVertex.PositionColored(vecDiff.X - fIndX, vecDiff.Y + (fSize * 10) + fSpriteOffsetY, vecDiff.Z - fIndZ, lIndColor)
                    uVerts1(2) = New CustomVertex.PositionColored(vecDiff.X + fIndX, vecDiff.Y + (fSize * 10) + fSpriteOffsetY, vecDiff.Z + fIndZ, lIndColor)
                    moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, 1, uVerts1)
                End With
            End If
        Next X

        'Reset our renderstates
        With moDevice
            .RenderState.ZBufferEnable = lPrevZBuffer 'True
            .RenderState.Lighting = lPrevLighting 'True
            .SetTexture(0, Nothing)
            '.Material = oMat
            .RenderState.CullMode = lPrevCullMode
            .VertexFormat = lPrevVertexFormat

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

            .RenderState.SourceBlend = lPrevSrcBlnd
            .RenderState.DestinationBlend = lPrevDestBlnd
            .RenderState.AlphaBlendEnable = lPrevAlplaBlnd

        End With
    End Sub

    Private Sub RenderWormholes()
        Dim lPrevLighting As Boolean = moDevice.RenderState.Lighting
        Dim lPrevVertexFormat As VertexFormats = moDevice.VertexFormat
        Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend


        If WpnFXManager.moTexture Is Nothing OrElse WpnFXManager.moTexture.Disposed = True Then WpnFXManager.moTexture = goResMgr.GetTexture("WpnFire.dds", GFXResourceManager.eGetTextureType.NoSpecifics)

        With moDevice
            .Transform.World = Matrix.Identity
            .SetTexture(0, Nothing) 'WpnFXManager.moTexture)
            .VertexFormat = CustomVertex.PositionTextured.Format
            .RenderState.Lighting = False
            .RenderState.DestinationBlend = Blend.One
        End With

        Dim oMat As Material
        With oMat
            '.Ambient = Color.Black
            .Emissive = Color.White
            .Diffuse = System.Drawing.Color.FromArgb(128, 255, 255, 255)
        End With

        moDevice.Material = oMat


        Dim fTuLow As Single = 0
        Dim fTuHi As Single = 0.5F
        Dim fTvLow As Single = 0.25F
        Dim fTvHi As Single = 0.5F

        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X) Is Nothing = False Then
                For Y As Int32 = 0 To goGalaxy.moSystems(X).WormholeUB
                    If goGalaxy.moSystems(X).moWormholes(Y) Is Nothing = False Then
                        With goGalaxy.moSystems(X).moWormholes(Y)
                            If .System1 Is Nothing = False AndAlso .System1.ObjectID = goGalaxy.moSystems(X).ObjectID AndAlso .System2 Is Nothing = False Then

                                Dim uVerts(3) As CustomVertex.PositionTextured

                                Dim fSize As Single = 1
                                Dim fCurrentX As Single = .System1.LocX
                                Dim fCurrentY As Single = .System1.LocY + 8
                                Dim fCurrentZ As Single = .System1.LocZ
                                Dim fDestX As Single = .System2.LocX
                                Dim fDestY As Single = .System2.LocY + 8
                                Dim fDestZ As Single = .System2.LocZ

                                uVerts(0) = New CustomVertex.PositionTextured(fCurrentX - fSize, fCurrentY, fCurrentZ + (fSize / 2), fTuLow, fTvLow)
                                uVerts(1) = New CustomVertex.PositionTextured(fCurrentX + fSize, fCurrentY, fCurrentZ - (fSize / 2), fTuLow, fTvHi)
                                uVerts(2) = New CustomVertex.PositionTextured(fDestX - fSize, fDestY, fDestZ + (fSize / 2), fTuHi, fTvLow)
                                uVerts(3) = New CustomVertex.PositionTextured(fDestX + fSize, fDestY, fDestZ - (fSize / 2), fTuHi, fTvHi)
                                moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

                                uVerts(0) = New CustomVertex.PositionTextured(fCurrentX + (fSize / 2), fCurrentY + fSize, fCurrentZ, fTuLow, fTvLow)
                                uVerts(1) = New CustomVertex.PositionTextured(fCurrentX - (fSize / 2), fCurrentY - fSize, fCurrentZ, fTuLow, fTvHi)
                                uVerts(2) = New CustomVertex.PositionTextured(fDestX + (fSize / 2), fDestY + fSize, fDestZ, fTuHi, fTvLow)
                                uVerts(3) = New CustomVertex.PositionTextured(fDestX - (fSize / 2), fDestY - fSize, fDestZ, fTuHi, fTvHi)
                                moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

                            End If
                        End With
                    End If
                Next Y
            End If
        Next X

        'Reset our renderstates
        With moDevice
            .RenderState.Lighting = lPrevLighting 'True
            .SetTexture(0, Nothing)

            .VertexFormat = lPrevVertexFormat

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

            .RenderState.DestinationBlend = lPrevDestBlnd
        End With
    End Sub

    Private Sub RenderSolarSystemSelectionBox(ByVal oSystem As SolarSystem)
        'Exit Sub
        If oSystem Is Nothing Then Exit Sub

        Dim oMat As Material
        With oMat
            .Ambient = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            .Diffuse = .Ambient
            .Emissive = .Ambient
        End With

        If moBoxMeshSelStar Is Nothing OrElse moBoxMeshSelStar.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moBoxMeshSelStar = Nothing
            moBoxMeshSelStar = goResMgr.CreateTexturedBox(25, 25, 25)
            Device.IsUsingEventHandlers = True
        End If

        Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
        'Dim lPrevLighting As Boolean = moDevice.RenderState.Lighting
        'Dim lPrevCullMode As Cull = moDevice.RenderState.CullMode
        'Dim lPrevVertexFormat As VertexFormats = moDevice.VertexFormat
        'Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
        'Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
        'Dim lPrevAlplaBlnd As Boolean = moDevice.RenderState.AlphaBlendEnable

        moDevice.SetTexture(0, goResMgr.GetTexture("SelectedTex.dds", GFXResourceManager.eGetTextureType.NoSpecifics, ""))
        moDevice.RenderState.ZBufferEnable = False


        Dim matTemp As Matrix = Matrix.Identity
        'oSystem.SquareCenterX()
        matTemp.Multiply(Matrix.Translation(oSystem.LocX, oSystem.LocY, oSystem.LocZ))
        moDevice.Transform.World = matTemp
        moDevice.Material = oMat
        moBoxMeshSelStar.DrawSubset(0)
        matTemp = Nothing
        'oBoxMesh = Nothing
        oMat = Nothing

        'Reset our renderstates
        With moDevice
            .RenderState.ZBufferEnable = lPrevZBuffer 'True
            '.RenderState.Lighting = lPrevLighting 'True
            .SetTexture(0, Nothing)

            '.RenderState.CullMode = lPrevCullMode
            '.VertexFormat = lPrevVertexFormat

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

            '.RenderState.SourceBlend = lPrevSrcBlnd
            '.RenderState.DestinationBlend = lPrevDestBlnd
            '.RenderState.AlphaBlendEnable = lPrevAlplaBlnd
        End With
    End Sub

    Private Sub RenderEntityConstructions()
        Dim lAA0 As Int32 = moDevice.GetTextureStageStateInt32(0, TextureStageStates.AlphaArgument0)
        Dim lAA1 As Int32 = moDevice.GetTextureStageStateInt32(0, TextureStageStates.AlphaArgument1)
        Dim lAO As Int32 = moDevice.GetTextureStageStateInt32(0, TextureStageStates.AlphaOperation)
        Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
        Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
        Dim lPrevBlndOp As BlendOperation = moDevice.RenderState.BlendOperation

        Try
            moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument0, TextureArgument.TextureColor)
            moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
            moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

            If goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit AndAlso goCurrentEnvir.oEntity(X).bProducing = True Then
                        Dim fPerc As Single = goCurrentEnvir.oEntity(X).GetProductionStatus()

                        If fPerc = -1 Then
                            goCurrentEnvir.oEntity(X).bProducing = False
                            Continue For
                        End If

                        If fPerc = Single.NegativeInfinity Then fPerc = 0.0F
                        If fPerc > 1.0F Then
                            fPerc = 1.0F
                            goCurrentEnvir.oEntity(X).bProducing = False
                        End If

                        Dim lVal As Int32
                        lVal = CInt(fPerc * 255)

                        If lVal < 0 Then lVal = 0
                        If lVal > 255 Then lVal = 255

                        Dim matProgress As Material
                        Dim oMesh As BaseMesh = goCurrentEnvir.oEntity(X).GetProducingModel()
                        If oMesh Is Nothing = False Then
                            If lVal = 255 Then
                                matProgress = oMesh.Materials(0)
                            Else
                                With oMesh.Materials(0)
                                    matProgress.Ambient = System.Drawing.Color.FromArgb(lVal, .Ambient.R, .Ambient.G, .Ambient.B)
                                    matProgress.Diffuse = System.Drawing.Color.FromArgb(lVal, .Diffuse.R, .Diffuse.G, .Diffuse.B)
                                    matProgress.Specular = System.Drawing.Color.FromArgb(lVal, .Specular.R, .Specular.G, .Specular.B)
                                    matProgress.SpecularSharpness = .SpecularSharpness
                                End With
                            End If

                            Dim matWorld As Matrix = Matrix.Identity
                            Dim lTempY As Int32 = 0
                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                                lTempY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(goCurrentEnvir.oEntity(X).LocX, goCurrentEnvir.oEntity(X).LocZ, True))
                            End If
                            matWorld.Multiply(Matrix.RotationY(DegreeToRadian((goCurrentEnvir.oEntity(X).iProductionAngle - 900) / 10.0F)))
                            'matWorld.Multiply(Matrix.Translation(goCurrentEnvir.oEntity(X).fMapWrapLocX, lTempY, goCurrentEnvir.oEntity(X).LocZ))
                            matWorld.Multiply(Matrix.Translation(goCurrentEnvir.oEntity(X).LocX, lTempY, goCurrentEnvir.oEntity(X).LocZ))
                            moDevice.Transform.World = matWorld

                            With oMesh
                                For Y As Int32 = 0 To .NumOfMaterials - 1
                                    moDevice.Material = matProgress
                                    moDevice.SetTexture(0, .Textures(Y * 4))
                                    .oMesh.DrawSubset(Y)
                                Next Y
                            End With


                            moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Disable)

                            moDevice.RenderState.SourceBlend = Blend.InvSourceAlpha
                            moDevice.RenderState.DestinationBlend = Blend.One
                            moDevice.RenderState.BlendOperation = BlendOperation.Add

                            With oMesh
                                For Y As Int32 = 0 To .NumOfMaterials - 1
                                    moDevice.Material = matProgress
                                    moDevice.SetTexture(0, .Textures(Y * 4))
                                    .oMesh.DrawSubset(Y)
                                Next Y
                            End With

                            moDevice.RenderState.SourceBlend = lPrevSrcBlnd
                            moDevice.RenderState.DestinationBlend = lPrevDestBlnd
                            moDevice.RenderState.BlendOperation = lPrevBlndOp
                            moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
                        End If

                    End If
                Next X
            End If

            If goCurrentEnvir Is Nothing = False Then
                Dim matProgress As Material
                With matProgress
                    Dim lVal As Int32 = 0
                    matProgress.Ambient = System.Drawing.Color.FromArgb(lVal, 255, 255, 255)
                    matProgress.Diffuse = System.Drawing.Color.FromArgb(lVal, 255, 255, 255)
                    matProgress.Specular = System.Drawing.Color.FromArgb(lVal, 255, 255, 255)
                    matProgress.SpecularSharpness = .SpecularSharpness
                End With
                For X As Int32 = 0 To goCurrentEnvir.lUnitQueueItemUB
                    If goCurrentEnvir.oUnitQueueItem(X) Is Nothing = False Then

                        With goCurrentEnvir.oUnitQueueItem(X)
                            If .oModel Is Nothing Then
                                If .iModelID = 0 Then
                                    For Y As Int32 = 0 To glEntityDefUB
                                        If glEntityDefIdx(Y) = .lProdID AndAlso goEntityDefs(Y).ObjTypeID = .iProdTypeID Then
                                            .iModelID = goEntityDefs(Y).ModelID
                                            Exit For
                                        End If
                                    Next Y

                                    If .iModelID = 0 Then Continue For
                                End If
                                .oModel = goResMgr.GetMesh(.iModelID)
                                If .oModel Is Nothing Then Continue For
                            End If

                            'Now, let's render...
                            Dim matWorld As Matrix = Matrix.Identity
                            Dim lTempY As Int32 = 0
                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso goCurrentEnvir.oGeoObject Is Nothing = False Then
                                lTempY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(.lLocX, .lLocZ, True))
                            End If
                            matWorld.Multiply(Matrix.RotationY(DegreeToRadian((.iLocA - 900) / 10.0F)))
                            'matWorld.Multiply(Matrix.Translation(goCurrentEnvir.oEntity(X).fMapWrapLocX, lTempY, goCurrentEnvir.oEntity(X).LocZ))
                            matWorld.Multiply(Matrix.Translation(.lLocX, lTempY, .lLocZ))
                            moDevice.Transform.World = matWorld

                            With .oModel
                                For Y As Int32 = 0 To .NumOfMaterials - 1
                                    moDevice.Material = matProgress
                                    moDevice.SetTexture(0, .Textures(Y * 4))
                                    .oMesh.DrawSubset(Y)
                                Next Y
                            End With


                            moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Disable)

                            moDevice.RenderState.SourceBlend = Blend.InvSourceAlpha
                            moDevice.RenderState.DestinationBlend = Blend.One
                            moDevice.RenderState.BlendOperation = BlendOperation.Add

                            With .oModel
                                For Y As Int32 = 0 To .NumOfMaterials - 1
                                    moDevice.Material = matProgress
                                    moDevice.SetTexture(0, .Textures(Y * 4))
                                    .oMesh.DrawSubset(Y)
                                Next Y
                            End With

                            moDevice.RenderState.SourceBlend = lPrevSrcBlnd
                            moDevice.RenderState.DestinationBlend = lPrevDestBlnd
                            moDevice.RenderState.BlendOperation = lPrevBlndOp
                            moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
                        End With
                    End If
                Next X
            End If

        Catch
        End Try

        moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument0, lAA0)
        moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument1, lAA1)
        moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, lAO)
        moDevice.RenderState.SourceBlend = lPrevSrcBlnd
        moDevice.RenderState.DestinationBlend = lPrevDestBlnd
        moDevice.RenderState.BlendOperation = lPrevBlndOp
    End Sub

    Private Sub RenderBombardmentReticle()
        If moBombMesh Is Nothing OrElse moBombMesh.Disposed = True Then
            moBombMesh = goResMgr.LoadScratchMeshNoTextures("reticle.x", "misc.pak")
        End If
        If moBombTex Is Nothing OrElse moBombTex.Disposed = True Then
            moBombTex = goResMgr.GetTexture("reticle.dds", GFXResourceManager.eGetTextureType.ModelTexture, "ob.pak")
        End If

        Dim lRange As Int32 = 10000
        Select Case goUILib.yBombardType
            Case BombardType.eHighYield_BT
                '20k
                lRange = 20000
            Case BombardType.eNormal_BT
                '10k
                lRange = 10000
            Case BombardType.ePrecision_BT
                '1k
                lRange = 5000
            Case BombardType.eTargeted_BT
                '5k
                lRange = 5000
        End Select

        lRange \= 50

        Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
        Dim lPrevLighting As Boolean = moDevice.RenderState.Lighting
        Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
        Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
        Dim lPrevAlplaBlnd As Boolean = moDevice.RenderState.AlphaBlendEnable
        With moDevice.RenderState
            .SourceBlend = Blend.SourceAlpha
            .DestinationBlend = Blend.InvSourceAlpha
            .AlphaBlendEnable = True
            .ZBufferEnable = False
            .Lighting = False
        End With

        'ok, now, place it...
        Dim matWorld As Matrix = Matrix.Identity
        Dim vecTemp As Vector3 = goUILib.vecBombardLoc
        'vecTemp.X /= lRange
        'vecTemp.Y = 0
        'vecTemp.Z /= lRange
        mfBombRot += 0.001F
        matWorld.Multiply(Matrix.Scaling(New Vector3(lRange, lRange, lRange)))
        matWorld.Multiply(Matrix.RotationY(mfBombRot))
        matWorld.Multiply(Matrix.Translation(vecTemp))
        moDevice.Transform.World = matWorld

        moDevice.SetTexture(0, moBombTex)
        Dim oMat As Material
        With oMat
            .Ambient = Color.White
            .Diffuse = Color.White
            .Specular = Color.Black
        End With
        moDevice.Material = oMat
        moBombMesh.DrawSubset(0)

        With moDevice.RenderState
            .ZBufferEnable = lPrevZBuffer
            .Lighting = lPrevLighting
            .SourceBlend = lPrevSrcBlnd
            .DestinationBlend = lPrevDestBlnd
            .AlphaBlendEnable = lPrevAlplaBlnd
        End With
    End Sub

    Private Sub RenderEntityTargets()


        Dim oRelMats(3) As Material
        With oRelMats(0)
            .Ambient = System.Drawing.Color.FromArgb(255, 0, 0, 255)
            .Diffuse = .Ambient
            .Emissive = .Ambient
        End With
        With oRelMats(1)
            .Ambient = System.Drawing.Color.FromArgb(255, 255, 0, 255)
            .Diffuse = .Ambient
            .Emissive = .Ambient
        End With
        With oRelMats(2)
            .Ambient = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            .Diffuse = .Ambient
            .Emissive = .Ambient
        End With
        With oRelMats(3)
            .Ambient = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .Diffuse = .Ambient
            .Emissive = .Diffuse
        End With

        If moBoxMeshTargets Is Nothing OrElse moBoxMeshTargets.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moBoxMeshTargets = Nothing
            moBoxMeshTargets = goResMgr.CreateTexturedBox(10, 10, 10)
            Device.IsUsingEventHandlers = True
        End If

        moDevice.SetTexture(0, goResMgr.GetTexture("SelectedTex.dds", GFXResourceManager.eGetTextureType.NoSpecifics, ""))

        Dim lPrevZBuffer As Boolean = moDevice.RenderState.ZBufferEnable
        moDevice.RenderState.ZBufferEnable = False
        Try
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                For X As Int32 = 0 To oEnvir.lEntityUB
                    If oEnvir.lEntityIdx(X) <> -1 Then
                        Dim oTmpEntity As BaseEntity = oEnvir.oEntity(X)
                        If oTmpEntity Is Nothing Then Continue For
                        With oTmpEntity
                            If oTmpEntity.bSelected = True Then
                                If oTmpEntity.lTargetIdx Is Nothing = False Then
                                    For Y As Int32 = 0 To 3
                                        If oTmpEntity.lTargetIdx(Y) > -1 Then
                                            Dim oNew As BaseEntity = oEnvir.oEntity(oTmpEntity.lTargetIdx(Y))
                                            If oNew Is Nothing = False Then
                                                'ReDim Preserve uVerts(lVertUB + 6)
                                                With oNew
                                                    Dim mat As Matrix = Matrix.Identity
                                                    Dim fScaling As Single = .oMesh.ShieldXZRadius / 10.0F
                                                    mat.Multiply(Matrix.Scaling(fScaling, fScaling, fScaling))
                                                    If .oMesh.bLandBased = True Then
                                                        'mat.Multiply(Matrix.Translation(.fMapWrapLocX, .LocY, .LocZ))
                                                        mat.Multiply(Matrix.Translation(.LocX, .LocY, .LocZ))
                                                    Else
                                                        'mat.Multiply(Matrix.Translation(.fMapWrapLocX, .LocY - .oMesh.vecHalfDeathSeqSize.Y, .LocZ))
                                                        mat.Multiply(Matrix.Translation(.LocX, .LocY - .oMesh.vecHalfDeathSeqSize.Y, .LocZ))
                                                    End If

                                                    moDevice.Transform.World = mat

                                                    moDevice.Material = oRelMats(Y)
                                                    moBoxMeshTargets.DrawSubset(0)

                                                End With
                                            End If
                                        End If
                                    Next Y
                                End If
                            End If
                        End With
                    End If
                Next X
            End If
        Catch
        End Try

        moDevice.RenderState.ZBufferEnable = lPrevZBuffer


    End Sub


    Public Shared Sub ToggleFullScreen()
        mbToggleFullScreen = True
    End Sub
    Public Sub DoToggleFullScreen()
        mbToggleFullScreen = False
        If muSettings.Windowed = True Then
            'ok, going from windowed to full screen
            muSettings.Windowed = False
            mfrmMain.FormBorderStyle = FormBorderStyle.None
        Else
            'ok, going from full screen to windowed
            muSettings.Windowed = True
            mfrmMain.FormBorderStyle = FormBorderStyle.FixedSingle
        End If

        gbPaused = True
        gbDeviceLost = True
        gbPaused = False



    End Sub
End Class
