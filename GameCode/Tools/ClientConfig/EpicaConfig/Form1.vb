Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmMain

    Private moDevice As Device
	Private moLogoFX As BPLogo 'EpicaLogoFX
    Private moGlowFX As PostShader
    Private moTerrain As TerrainClass
    Private moMesh As EpicaMesh
    Private moMeshNormal As Texture
    Private moMeshIllum As Texture

    Private mbLoading As Boolean = False

    Private mbRightDrag As Boolean = False
    Private mbRightDown As Boolean = False
    Private mlMouseX As Int32
    Private mlMouseY As Int32

    Private moOcean As OceanRender

    Private mbDataChanged As Boolean = False

    Public Sub New()
        Form.CheckForIllegalCrossThreadCalls = False

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        muSettings.LoadSettings()
        goCamera = New Camera
        FillAdapterList()
    End Sub

    Public Function InitD3D(ByRef picBox As PictureBox, ByVal lAdapter As Int32) As Boolean
        Dim uParms As PresentParameters
        Dim uDispMode As DisplayMode
        Dim oINI As InitFile
        Dim bWindowed As Boolean
        Dim bRes As Boolean

        Try
            Dim lCreateFlags As CreateFlags
            Dim bReduced As Boolean = False
            uDispMode = Manager.Adapters.Default.CurrentDisplayMode
            oINI = New InitFile("BPClient.ini")
            uParms = New PresentParameters()
            With uParms
                bWindowed = True '= CBool(Val(oINI.GetString("GRAPHICS", "Windowed", "1")) <> 0)
                .Windowed = bWindowed
                .SwapEffect = SwapEffect.Discard
                .BackBufferCount = 1
                If bWindowed Then
                    .BackBufferFormat = uDispMode.Format
                    .BackBufferHeight = picBox.ClientSize.Height
                    .BackBufferWidth = picBox.ClientSize.Width
                Else
                    'TODO: Change our resolution to whatever the settings are
                    .BackBufferFormat = uDispMode.Format
                    .BackBufferHeight = uDispMode.Height
                    .BackBufferWidth = uDispMode.Width
                End If

                Dim uDevCaps As Caps = Manager.GetDeviceCaps(lAdapter, DeviceType.Hardware)

                Dim lPrsntInt As Int32 = PresentInterval.Default
                If bWindowed = True Then
                    If (uDevCaps.PresentationIntervals And PresentInterval.Immediate) = PresentInterval.Immediate Then
                        lPrsntInt = PresentInterval.Immediate
                    End If
                Else
                    If (uDevCaps.PresentationIntervals And PresentInterval.Immediate) = PresentInterval.Immediate Then
                        lPrsntInt = PresentInterval.Immediate
                    ElseIf (uDevCaps.PresentationIntervals And PresentInterval.One) = PresentInterval.One Then
                        lPrsntInt = PresentInterval.One
                    End If
                End If 
                Dim lTemp As Int32 = CInt(Val(oINI.GetString("GRAPHICS", "PresentInterval", (0).ToString)))
                If lTemp <> 0 Then lPrsntInt = lTemp
                .PresentationInterval = CType(lPrsntInt, PresentInterval)
                '.PresentFlag = PresentFlag.DiscardDepthStencil

                Dim bD32 As Boolean = Manager.CheckDepthStencilMatch(lAdapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D32, 0)
                Dim bD24X8 As Boolean = Manager.CheckDepthStencilMatch(lAdapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24X8, 0)
                Dim bD24S8 As Boolean = Manager.CheckDepthStencilMatch(lAdapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24S8, 0)
                Dim bD24X4S4 As Boolean = Manager.CheckDepthStencilMatch(lAdapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D24X4S4, 0)
                Dim bD16 As Boolean = Manager.CheckDepthStencilMatch(lAdapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D16, 0)
                Dim bD15S1 As Boolean = Manager.CheckDepthStencilMatch(lAdapter, DeviceType.Hardware, uDispMode.Format, uDispMode.Format, DepthFormat.D15S1, 0)

                Dim lDpthFmt As Int32
				'If cboZBuffer.Items.Count = 0 Then
				'	If bD32 = True Then cboZBuffer.Items.Add("32-Bit")
				'	If bD24X8 = True Then cboZBuffer.Items.Add("24X8")
				'	If bD24S8 = True Then cboZBuffer.Items.Add("24S8")
				'	If bD16 = True Then cboZBuffer.Items.Add("16")
				'End If

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
				lDpthFmt = CInt(Val(oINI.GetString("GRAPHICS", "DepthBuffer", lDpthFmt.ToString)))

				'Select Case lDpthFmt
				'	Case DepthFormat.D32
				'		cboZBuffer.Text = "32-Bit"
				'	Case DepthFormat.D24X8
				'		cboZBuffer.Text = "24X8"
				'	Case DepthFormat.D24S8
				'		cboZBuffer.Text = "24S8"
				'	Case DepthFormat.D16
				'		cboZBuffer.Text = "D16"
				'End Select

				.AutoDepthStencilFormat = CType(lDpthFmt, DepthFormat)
				.EnableAutoDepthStencil = True

				lCreateFlags = CType(Val(oINI.GetString("GRAPHICS", "VertexProcessing", CInt(CreateFlags.MixedVertexProcessing).ToString)), CreateFlags)
				If uDevCaps.DeviceCaps.SupportsHardwareTransformAndLight = False Then
					If lCreateFlags = CreateFlags.HardwareVertexProcessing Then
						lCreateFlags = CreateFlags.SoftwareVertexProcessing
						bReduced = True
					End If
				End If

			End With
            If muSettings.VertexProcessing <> 0 Then lCreateFlags = CType(muSettings.VertexProcessing, CreateFlags)

            Try
                moDevice = New Device(lAdapter, DeviceType.Hardware, picBox.Handle, lCreateFlags, uParms)
            Catch
                bReduced = True
                If (lCreateFlags And CreateFlags.HardwareVertexProcessing) = CreateFlags.HardwareVertexProcessing Then
                    lCreateFlags = (lCreateFlags Xor CreateFlags.HardwareVertexProcessing) Or CreateFlags.MixedVertexProcessing
                    Try
                        moDevice = New Device(lAdapter, DeviceType.Hardware, picBox.Handle, lCreateFlags, uParms)
                    Catch
                        lCreateFlags = (lCreateFlags Xor CreateFlags.MixedVertexProcessing) Or CreateFlags.SoftwareVertexProcessing
                        Try
                            moDevice = New Device(lAdapter, DeviceType.Hardware, picBox.Handle, lCreateFlags, uParms)
                        Catch ex As Exception
                            MsgBox("The client will not run on this computer because the video hardware is not sufficient.", MsgBoxStyle.OkOnly, "Error")
                            End
                        End Try
                    End Try
                ElseIf (lCreateFlags And CreateFlags.MixedVertexProcessing) = CreateFlags.MixedVertexProcessing Then
                    lCreateFlags = (lCreateFlags Xor CreateFlags.MixedVertexProcessing) Or CreateFlags.SoftwareVertexProcessing
                    Try
                        moDevice = New Device(lAdapter, DeviceType.Hardware, picBox.Handle, lCreateFlags, uParms)
                    Catch ex As Exception
                        MsgBox("The client will not run on this computer because the video hardware is not sufficient.", MsgBoxStyle.OkOnly, "Error")
                        End
                    End Try
                End If
            End Try
            muSettings.VertexProcessing = lCreateFlags

            'If bReduced = True Then MsgBox("Epica was required to alter settings that will affect performance in order to run the game.", MsgBoxStyle.OkOnly, "Error")
            oINI.WriteString("GRAPHICS", "VertexProcessing", CInt(lCreateFlags).ToString)

            bRes = Not moDevice Is Nothing

            'If muSettings.MiniMapLocX > uParms.BackBufferWidth Then muSettings.MiniMapLocX = uParms.BackBufferWidth - muSettings.MiniMapWidthHeight
            'If muSettings.MiniMapLocY > uParms.BackBufferHeight Then muSettings.MiniMapLocY = uParms.BackBufferHeight - muSettings.MiniMapWidthHeight
        Catch
            MsgBox(Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Error in initialize Direct3D")
            bRes = False
        Finally
            uParms = Nothing
            uDispMode = Nothing
        End Try

        With moDevice.DeviceCaps.VertexShaderVersion
            If Not (.Major > 1 OrElse (.Major = 1 AndAlso .Minor = 1)) Then
                muSettings.PostGlowAmt = 0.0F
                cboGlowFX.Enabled = False
            End If
        End With

        moResMgr = New EpicaResourceManager(moDevice)

        'mbSupportsNewModelMethod = moDevice.DeviceCaps.TextureOperationCaps.SupportsBlendTextureAlpha = True AndAlso _
        '   moDevice.DeviceCaps.MaxTextureBlendStages > 2 AndAlso moDevice.DeviceCaps.MaxSimultaneousTextures > 1

        'Instantiate our FX Managers
        'If goWpnMgr Is Nothing Then goWpnMgr = New WpnFXManager(moDevice)
        'If goShldMgr Is Nothing Then goShldMgr = New ShieldFXManager(moDevice)
        'If goExplMgr Is Nothing Then goExplMgr = New ExplosionFXManager(moDevice)
        'If goPFXEngine32 Is Nothing Then goPFXEngine32 = New BurnFX.ParticleEngine(moDevice, 32) '32 for size of points
        'If goEntityDeath Is Nothing Then goEntityDeath = New DeathSequenceMgr(moDevice)
        'If goMissileMgr Is Nothing Then goMissileMgr = New MissileMgr(moDevice)

        'InitializeFXColors()

        'AddHandler moDevice.DeviceLost, AddressOf moDevice_DeviceLost
        'AddHandler moDevice.DeviceReset, AddressOf moDevice_DeviceReset
        'AddHandler moDevice.Disposing, AddressOf moDevice_Disposing

        'mbInitialized = bRes

        Return bRes
    End Function

    Private Sub moDevice_DeviceLost(ByVal sender As Object, ByVal e As System.EventArgs)
        Debug.Write("GFXEngine.moDevice_DeviceLost" & vbCrLf)
    End Sub

    Private Sub moDevice_DeviceReset(ByVal sender As Object, ByVal e As System.EventArgs)
        Debug.Write("GFXEngine.moDevice_DeviceReset" & vbCrLf)
    End Sub

    Private Sub moDevice_Disposing(ByVal sender As Object, ByVal e As System.EventArgs)
        'TODO: Do any final cleanup here
        Debug.Write("GFXEngine.moDevice_Disposing" & vbCrLf)
    End Sub

    Private Sub FillAdapterList()
        cboDeviceAdapter.Items.Clear()
        For Each oAdapterInfo As AdapterInformation In Manager.Adapters
            cboDeviceAdapter.Items.Add(oAdapterInfo.Information.Description)
        Next

        If cboDeviceAdapter.Items.Count > 0 Then
            cboDeviceAdapter.SelectedIndex = 0
        ElseIf muSettings.AdapterOrdinal < cboDeviceAdapter.Items.Count Then
            cboDeviceAdapter.SelectedIndex = muSettings.AdapterOrdinal
        End If
    End Sub

    Private Sub cboDeviceAdapter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboDeviceAdapter.SelectedIndexChanged
        If cboDeviceAdapter.SelectedIndex > -1 Then
            Dim lAdapter As Int32 = Manager.Adapters(cboDeviceAdapter.SelectedIndex).Adapter

            'mbdatachanged = cbodeviceadapter.SelectedIndex <> musettings.

            mbLoading = True
            tmrRedraw.Enabled = False

            mbRunning = False
            While mbStopped = False
                Application.DoEvents()
            End While

			'If moLogoFX Is Nothing = False Then moLogoFX.DisposeMe()
            moLogoFX = Nothing
            If moGlowFX Is Nothing = False Then moGlowFX.DisposeMe()
            moGlowFX = Nothing
            If moTerrain Is Nothing = False Then moTerrain.CleanResources()
            moTerrain = Nothing
            If moDevice Is Nothing = False Then moDevice.Dispose()
            If moModelShader Is Nothing = False Then moModelShader.DisposeMe()
            moModelShader = Nothing
			moDevice = Nothing
			moMesh = Nothing
            TerrainClass.moDisk = Nothing

            InitD3D(picDisplay, lAdapter)

            cboFullScreenRes.Items.Clear()
            For Each cFmt As DisplayMode In Manager.Adapters(lAdapter).SupportedDisplayModes
                If cFmt.Width >= 1024 AndAlso cFmt.Height >= 768 Then
                    Dim sFmt As String = GetFormatNameFromFormat(cFmt.Format)
                    If sFmt = "32-bit" Then cboFullScreenRes.Items.Add(cFmt.Width & "x" & cFmt.Height & " " & cFmt.RefreshRate & "hz")
                End If
            Next

            cboVertProc.Items.Clear()
            cboVertProc.Items.Add("Software")
            cboVertProc.Items.Add("Mixed")
            If moDevice.DeviceCaps.DeviceCaps.SupportsHardwareTransformAndLight = True Then cboVertProc.Items.Add("Hardware")

            If moDevice.CreationParameters.Behavior.HardwareVertexProcessing = True Then
                cboVertProc.SelectedIndex = 2
            ElseIf moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
                cboVertProc.SelectedIndex = 1
            Else : cboVertProc.SelectedIndex = 0
            End If

            'Now, our texture resolutions
            cboFOWRes.Items.Clear()
            cboIllumMaps.Items.Clear()
            cboTextureResolution.Items.Clear()

            cboIllumMaps.Items.Add("Off - No Illumination Maps")
            Dim bSupportsIllumMaps As Boolean = moDevice.DeviceCaps.PixelShaderVersion.Major >= 2
            If bSupportsIllumMaps = True Then cboIllumMaps.Items.Add("Very Low Resolution")

            Dim lMaxTexRes As Int32 = Math.Min(moDevice.DeviceCaps.MaxTextureWidth, moDevice.DeviceCaps.MaxTextureHeight)
            cboFOWRes.Items.Add("Off - No FOW Shading")
            cboTextureResolution.Items.Add("Very Low Resolution")
            If lMaxTexRes >= 128 Then
                cboTextureResolution.Items.Add("Low Resolution")
                If bSupportsIllumMaps = True Then cboIllumMaps.Items.Add("Low Resolution")
            End If
            If lMaxTexRes >= 256 Then
                cboTextureResolution.Items.Add("Normal Resolution")
                If bSupportsIllumMaps = True Then cboIllumMaps.Items.Add("Normal Resolution")
            End If
            If lMaxTexRes >= 512 Then
                cboFOWRes.Items.Add("Normal Resolution")
                cboTextureResolution.Items.Add("High Resolution")
                If bSupportsIllumMaps = True Then cboIllumMaps.Items.Add("High Resolution")
            End If

            cboTerrTexRes.Items.Clear()
            cboTerrTexRes.Items.Add("Very Low Resolution")
            cboTerrTexRes.Items.Add("Low Resolution")
            cboTerrTexRes.Items.Add("Normal Resolution")

            cboBurnFX.Items.Clear() : cboPlanetFX.Items.Clear() : cboEngineFX.Items.Clear()
            cboBurnFX.Items.Add("Off") : cboPlanetFX.Items.Add("Off") : cboEngineFX.Items.Add("Off")
            cboBurnFX.Items.Add("Low") : cboPlanetFX.Items.Add("Low") : cboEngineFX.Items.Add("Low")
            cboBurnFX.Items.Add("Medium") : cboPlanetFX.Items.Add("Medium") : cboEngineFX.Items.Add("Medium")
            cboBurnFX.Items.Add("High") : cboPlanetFX.Items.Add("High") : cboEngineFX.Items.Add("High")
            cboBurnFX.Items.Add("Full") : cboPlanetFX.Items.Add("Full") : cboEngineFX.Items.Add("Full")


            cboWaterRes.Items.Clear()
            cboWaterRes.Items.Add("Low Res Water")
            If moDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then cboWaterRes.Items.Add("Shader 2.0")

            cboGlowFX.Items.Clear()
            With moDevice.DeviceCaps.VertexShaderVersion
                If Not (.Major > 1 OrElse (.Major = 1 AndAlso .Minor = 1)) Then
                    cboGlowFX.Enabled = False
                    muSettings.PostGlowAmt = 0.0F
                Else
                    For X As Int32 = 0 To 10
                        If X = 0 Then
                            cboGlowFX.Items.Add("Off")
                        Else
                            cboGlowFX.Items.Add("Level " & X.ToString)
                        End If
                    Next X
                End If
            End With

            hscrClipPlane.Value = muSettings.EntityClipPlane \ 100
            chkTripleBuffer.Checked = muSettings.TripleBuffer
            chkVerticalSync.Checked = muSettings.VSync
            chkDeepCosmos.Checked = muSettings.RenderCosmos
            chkPlayerCustom.Checked = muSettings.RenderPlayerCustomization
            chkBumpTerrain.Checked = muSettings.BumpMapTerrain
            chkBumpPlanet.Checked = muSettings.BumpMapPlanetModel
            chkIllumPlanet.Checked = muSettings.IlluminationMapTerrain

            If moDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
                chkPlayerCustom.Enabled = True
                chkBumpTerrain.Enabled = True
                chkBumpPlanet.Enabled = True
                chkIllumPlanet.Enabled = True
            Else
                chkPlayerCustom.Enabled = False : chkPlayerCustom.Checked = False
                chkBumpTerrain.Enabled = False : chkBumpTerrain.Checked = False
                chkBumpPlanet.Enabled = False : chkBumpPlanet.Checked = False
                chkIllumPlanet.Enabled = False : chkIllumPlanet.Enabled = False
            End If

            cboLightQuality.Items.Clear()
            cboLightQuality.Items.Add("Low Quality")
            If moDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
                cboLightQuality.Items.Add("Medium Quality")
                cboLightQuality.Items.Add("High Quality")
            End If
            If muSettings.LightQuality < cboLightQuality.Items.Count Then
                cboLightQuality.SelectedIndex = muSettings.LightQuality
            End If




            'If moLogoFX Is Nothing = False Then moLogoFX.DisposeMe()
            moLogoFX = Nothing
            moLogoFX = New BPLogo(moDevice, 19) ' EpicaLogoFX(moDevice, New Vector3(0, 300, 0), 1880)

            mbLoading = False

            SetSettings()

            mbRunning = True
            tmrRedraw.Enabled = True
        End If
    End Sub

    Private Function GetFormatNameFromFormat(ByVal uFormat As Format) As String
        Select Case uFormat
            Case Format.A8R8G8B8, Format.X8B8G8R8, Format.X8R8G8B8
                Return "32-bit"
            Case Format.R5G6B5
                Return "16-bit"
            Case Format.R8G8B8
                Return "24-bit"
        End Select
        Return ""
    End Function

    Private mbRunning As Boolean = True
    Private mbStopped As Boolean = True
    Private Sub tmrRedraw_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrRedraw.Tick
        mbStopped = False
        tmrRedraw.Enabled = False

		If moDevice Is Nothing = False Then
			moDevice.BeginScene()

			moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, Color.Black, 1, 0)

			SetRenderStates()
			moDevice.RenderState.Ambient = System.Drawing.Color.DarkGray

            goCamera.CalculateFrustrum(moDevice)

			Dim vecLightDir As Vector3
			vecLightDir = New Vector3(500, -1000, 500)

			With moDevice.Lights(0)
				.Diffuse = System.Drawing.Color.White
				.Ambient = System.Drawing.Color.DarkGray
				.Type = LightType.Directional
				'.Direction = New Vector3(-500, -500, -250)
				.Direction = vecLightDir
				.Range = 100000
				.Specular = System.Drawing.Color.White
				.Attenuation0 = 1
				.Attenuation1 = 0
				.Attenuation2 = 0
				.Falloff = 0.3
				.Enabled = True
				.Update()
			End With
			moDevice.RenderState.Lighting = True
			moDevice.Transform.World = Matrix.Identity
			If moTerrain Is Nothing Then moTerrain = New TerrainClass(moDevice, 12041)
			moTerrain.Render(20000)

			'then setup our matrices, and do it
			goCamera.SetupMatrices(moDevice)
			If moMesh Is Nothing Then moMesh = moResMgr.GetMesh(10)
			Dim matWorld As Matrix = Matrix.Identity
			matWorld.Multiply(Matrix.Translation(7200, 390, 4200))
            moDevice.Transform.World = matWorld

            If muSettings.LightQuality = EngineSettings.LightQualitySetting.VSPS1 Then
                moDevice.SetTexture(0, moMesh.Textures(0))
                moDevice.Material = moMesh.Materials(0)
                moMesh.oMesh.DrawSubset(0)
            Else
                If moModelShader Is Nothing Then moModelShader = New ModelShader(moDevice)
                moModelShader.PrepareToRender(Vector3.Normalize(moDevice.Lights(0).Direction), moDevice.Lights(0).Diffuse, moDevice.Lights(0).Ambient, moDevice.Lights(0).Specular)
                If moMeshNormal Is Nothing Then
                    moMeshNormal = moResMgr.GetTexture("bld-mt_n.dds", EpicaResourceManager.eGetTextureType.ModelTexture, "textures2.pak")
                End If
                If moMeshIllum Is Nothing Then
                    moMeshIllum = moResMgr.GetTexture("bld-mt_i.dds", EpicaResourceManager.eGetTextureType.ModelTexture, "textures2.pak")
                End If
                moModelShader.RenderMesh(moMesh.oMesh, moMesh.Textures(0), moMeshNormal, moMeshIllum)
                moModelShader.EndRender()
            End If
            

			'Always render our logo if it is available
			moDevice.Transform.World = Matrix.Identity
			If moLogoFX Is Nothing = False Then
				moLogoFX.Render(False, moDevice)
			End If
			moDevice.Transform.World = Matrix.Identity
			If muSettings.PostGlowAmt > 0 Then
				If moGlowFX Is Nothing Then moGlowFX = New PostShader(moDevice)
				moGlowFX.ExecutePostProcess()
			End If

			'then setup our matrices, and do it
			goCamera.SetupMatrices(moDevice)

			moDevice.EndScene()
			moDevice.Present()
		End If

 
        tmrRedraw.Enabled = True
        mbStopped = True 

    End Sub
    Private moModelShader As ModelShader

    Private Sub SetRenderStates()
        With moDevice.RenderState
            .Lighting = True
            .ZBufferEnable = True
            '.DitherEnable = muSettings.Dither
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

    End Sub

    Private Sub cboVertProc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVertProc.SelectedIndexChanged
        If mbLoading = False Then
            Dim oINI As New InitFile("BPClient.ini")
            Dim lCreateFlags As Int32
            Select Case cboVertProc.SelectedIndex
                Case 2
                    lCreateFlags = CreateFlags.HardwareVertexProcessing
                Case 1
                    lCreateFlags = CreateFlags.MixedVertexProcessing
                Case Else
                    lCreateFlags = CreateFlags.SoftwareVertexProcessing
            End Select
            muSettings.VertexProcessing = lCreateFlags
            oINI.WriteString("GRAPHICS", "VertexProcessing", CInt(lCreateFlags).ToString)
            oINI = Nothing

            cboDeviceAdapter_SelectedIndexChanged(sender, e)
            mbDataChanged = True
        End If
    End Sub

    Private Sub SetSettings()
        mbLoading = True

        If muSettings.ShowFOWTerrainShading = False Then
            cboFOWRes.SelectedIndex = cboFOWRes.FindStringExact("Off - No FOW Shading")
        Else
            Select Case muSettings.FOWTextureResolution
                Case 512
                    cboFOWRes.SelectedIndex = cboFOWRes.FindStringExact("Low Resolution")
                Case 1024
                    cboFOWRes.SelectedIndex = cboFOWRes.FindStringExact("Medium Resolution")
                Case 2048
                    cboFOWRes.SelectedIndex = cboFOWRes.FindStringExact("High Resolution")
            End Select
        End If
        
        Select Case muSettings.ModelTextureResolution
            Case EngineSettings.eTextureResOptions.eVeryLowResTextures
                cboTextureResolution.SelectedIndex = cboTextureResolution.FindStringExact("Very Low Resolution")
            Case EngineSettings.eTextureResOptions.eLowResTextures
                cboTextureResolution.SelectedIndex = cboTextureResolution.FindStringExact("Low Resolution")
            Case EngineSettings.eTextureResOptions.eNormResTextures
                cboTextureResolution.SelectedIndex = cboTextureResolution.FindStringExact("Normal Resolution")
            Case EngineSettings.eTextureResOptions.eHiResTextures
                cboTextureResolution.SelectedIndex = cboTextureResolution.FindStringExact("High Resolution")
        End Select

        Select Case muSettings.TerrainTextureResolution
            Case 1
                cboTerrTexRes.SelectedIndex = cboTerrTexRes.FindStringExact("Very Low Resolution")
            Case 2
                cboTerrTexRes.SelectedIndex = cboTerrTexRes.FindStringExact("Low Resolution")
            Case 3
                cboTerrTexRes.SelectedIndex = cboTerrTexRes.FindStringExact("Normal Resolution")
        End Select

        Select Case muSettings.IlluminationMap
            Case 0
                cboIllumMaps.SelectedIndex = 0
            Case 64
                cboIllumMaps.SelectedIndex = 1
            Case 128
                cboIllumMaps.SelectedIndex = 2
            Case 256
                cboIllumMaps.SelectedIndex = 3
            Case 512
                cboIllumMaps.SelectedIndex = 4
        End Select

        cboWaterRes.SelectedIndex = muSettings.WaterRenderMethod
        chkSmoothFOW.Checked = muSettings.SmoothFOW

        cboEngineFX.SelectedIndex = muSettings.EngineFXParticles
        cboBurnFX.SelectedIndex = muSettings.BurnFXParticles
        cboPlanetFX.SelectedIndex = muSettings.PlanetFXParticles
        chkShieldFX.Checked = muSettings.RenderShieldFX
        hscrStarPlanet.Value = muSettings.StarfieldParticlesPlanet \ 10
        hscrStarSpace.Value = muSettings.StarfieldParticlesSpace \ 100

        Dim lGlowFX As Int32 = CInt(muSettings.PostGlowAmt * 10)
        Select Case lGlowFX
            Case 0
                cboGlowFX.SelectedIndex = cboGlowFX.FindStringExact("Off")
            Case Else
                cboGlowFX.SelectedIndex = cboGlowFX.FindStringExact("Level " & lGlowFX)
        End Select

        chkHiResPlanet.Checked = muSettings.HiResPlanetTexture

        chkFullScreen.Checked = Not muSettings.Windowed
        Dim sTemp As String = muSettings.FullScreenResX & "x" & muSettings.FullScreenResY & " " & muSettings.FullScreenRefreshRate & "hz"
        cboFullScreenRes.SelectedIndex = cboFullScreenRes.FindStringExact(sTemp)

        mbLoading = False
    End Sub

    Private Sub cboGlowFX_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboGlowFX.SelectedIndexChanged
        If mbLoading = True Then Return
        muSettings.PostGlowAmt = cboGlowFX.SelectedIndex / 10.0F
        mbDataChanged = True
    End Sub

    'Private Sub doWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picDisplay.MouseWheel
    '    'I've left this here instead of hte camera class because I may want to capture these events for other things
    '    '  besides just zooming in and out... such as scrolling a listbox or something
    '    If e.Delta < 0 Then MouseWheelDown() Else MouseWheelUp()
    'End Sub

    'Private Sub MouseWheelDown()
    '    goCamera.ModifyZoom(150)
    'End Sub
    'Private Sub MouseWheelUp()
    '    goCamera.ModifyZoom(-150)
    'End Sub
    'Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picDisplay.MouseDown
    '    mbRightDown = e.Button = Windows.Forms.MouseButtons.Right
    'End Sub
    'Private Sub frmMain_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picDisplay.MouseMove
    '    Dim deltaX As Int32
    '    Dim deltaY As Int32

    '    If mbRightDown Then
    '        mbRightDrag = True
    '        deltaX = e.X - mlMouseX
    '        deltaY = e.Y - mlMouseY

    '        goCamera.RotateCamera(deltaX, deltaY)
    '    Else
    '        With goCamera
    '            .bScrollDown = False : .bScrollLeft = False : .bScrollRight = False : .bScrollUp = False

    '            If e.X < 10 Then .bScrollLeft = True
    '            If e.X > Me.ClientSize.Width - 10 Then .bScrollRight = True
    '            If e.Y < 10 Then .bScrollUp = True
    '            If e.Y > Me.ClientSize.Height - 10 Then .bScrollDown = True
    '        End With

    '    End If

    '    mlMouseX = e.X
    '    mlMouseY = e.Y
    'End Sub
    'Private Sub frmMain_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picDisplay.MouseUp
    '    mbRightDrag = False
    '    mbRightDown = False
    'End Sub

    Private Sub btnToggleDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToggleDisplay.Click
        If btnToggleDisplay.Text = "Show Compare Image" Then
            btnToggleDisplay.Text = "Show Real-Time Display"
            picPreRendered.Visible = True
            picPreRendered.BringToFront()
        Else
            btnToggleDisplay.Text = "Show Compare Image"
            picPreRendered.Visible = False
        End If
    End Sub

    Private Sub cboFOWRes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFOWRes.SelectedIndexChanged
        If mbLoading = True Then Return
        If cboFOWRes.SelectedIndex < 1 Then
            muSettings.ShowFOWTerrainShading = False
            muSettings.FOWTextureResolution = 0
        Else
            muSettings.ShowFOWTerrainShading = True
            Select Case cboFOWRes.SelectedIndex
                Case 1
                    muSettings.FOWTextureResolution = 512
                Case 2
                    muSettings.FOWTextureResolution = 1024
                Case 3
                    muSettings.FOWTextureResolution = 2048
            End Select
        End If
        TerrainClass.ReleaseVisibleTexture()
        mbDataChanged = True
    End Sub

    Private Sub cboTerrTexRes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTerrTexRes.SelectedIndexChanged
        If mbLoading = True Then Return
        muSettings.TerrainTextureResolution = Math.Max(1, cboTerrTexRes.SelectedIndex + 1)
        TerrainClass.bReloadVertexBuffer = True
        mbDataChanged = True
    End Sub

    Private Sub cboWaterRes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWaterRes.SelectedIndexChanged
        'do nothing for now
		If mbLoading = True Then Return
        muSettings.WaterRenderMethod = cboWaterRes.SelectedIndex
        mbDataChanged = True
    End Sub

    Private Sub cboTextureResolution_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTextureResolution.SelectedIndexChanged
        If mbLoading = True Then Return
        Select Case cboTextureResolution.SelectedIndex
            Case 0
                muSettings.ModelTextureResolution = EngineSettings.eTextureResOptions.eVeryLowResTextures
            Case 1
                muSettings.ModelTextureResolution = EngineSettings.eTextureResOptions.eLowResTextures
            Case 2
                muSettings.ModelTextureResolution = EngineSettings.eTextureResOptions.eNormResTextures
            Case Else
                muSettings.ModelTextureResolution = EngineSettings.eTextureResOptions.eHiResTextures
        End Select
        moResMgr.UpdateModelTextureResolution()
        mbDataChanged = True
    End Sub

    Private Sub chkHiResPlanet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHiResPlanet.CheckedChanged
        If mbLoading = True Then Return
        muSettings.HiResPlanetTexture = chkHiResPlanet.Checked
        mbDataChanged = True
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        muSettings.SaveSettings()
        mbDataChanged = False
    End Sub

    Private Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        muSettings.LoadSettings()
        SetSettings()
    End Sub
 
    Private Sub btnLaunch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLaunch.Click
        If CheckSaveChanges() = False Then Return
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"
        Shell(sPath & "UpdaterClient.exe", AppWinStyle.NormalFocus)
        End
    End Sub

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If CheckSaveChanges() = False Then e.Cancel = True
    End Sub

    Private Function CheckSaveChanges() As Boolean
        If mbDataChanged = True Then
            Dim uResult As MsgBoxResult = MsgBox("You have changed your settings, do you wish to save them?", MsgBoxStyle.YesNoCancel, "Change settings")
            If uResult = MsgBoxResult.Yes Then
                muSettings.SaveSettings()
                Return True
            ElseIf uResult = MsgBoxResult.Cancel Then
                Return False
            Else : Return True
            End If
        Else : Return True
        End If
    End Function

	'Private Sub cboZBuffer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboZBuffer.SelectedIndexChanged
	'	Dim sText As String = cboZBuffer.Text
	'	If sText Is Nothing = False Then
	'		sText = sText.ToUpper
	'		Dim oINI As InitFile = New InitFile("BPClient.ini")
	'		Select Case sText
	'			Case "32-BIT"
	'				oINI.WriteString("GRAPHICS", "DepthBuffer", CInt(DepthFormat.D32).ToString)
	'			Case "24X8"
	'				oINI.WriteString("GRAPHICS", "DepthBuffer", CInt(DepthFormat.D24X8).ToString)
	'			Case "24S8"
	'				oINI.WriteString("GRAPHICS", "DepthBuffer", CInt(DepthFormat.D24S8).ToString)
	'			Case "16"
	'				oINI.WriteString("GRAPHICS", "DepthBuffer", CInt(DepthFormat.D16).ToString)
	'		End Select

	'		cboDeviceAdapter_SelectedIndexChanged(Nothing, e)
	'	End If
	'End Sub

	Private Sub chkSmoothFOW_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSmoothFOW.CheckedChanged
		If mbLoading = True Then Return
		muSettings.SmoothFOW = chkSmoothFOW.Checked
		mbDataChanged = True
	End Sub

    Private Sub chkIllumPlanet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIllumPlanet.CheckedChanged
        If mbLoading = True Then Return
        muSettings.IlluminationMapTerrain = chkIllumPlanet.Checked
        mbDataChanged = True
    End Sub

    Private Sub chkBumpPlanet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBumpPlanet.CheckedChanged
        If mbLoading = True Then Return
        muSettings.BumpMapPlanetModel = chkBumpPlanet.Checked
        mbDataChanged = True
    End Sub

    Private Sub chkBumpTerrain_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBumpTerrain.CheckedChanged
        If mbLoading = True Then Return
        muSettings.BumpMapTerrain = chkBumpTerrain.Checked
        mbDataChanged = True
    End Sub

    Private Sub chkPlayerCustom_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPlayerCustom.CheckedChanged
        If mbLoading = True Then Return
        muSettings.RenderPlayerCustomization = chkPlayerCustom.Checked
        mbDataChanged = True
    End Sub

    Private Sub cboIllumMaps_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboIllumMaps.SelectedIndexChanged
        If mbLoading = True Then Return
        'off, verylow, low, normal, high
        Select Case cboIllumMaps.SelectedIndex
            Case 1
                muSettings.IlluminationMap = EngineSettings.eTextureResOptions.eVeryLowResTextures
            Case 2
                muSettings.IlluminationMap = EngineSettings.eTextureResOptions.eLowResTextures
            Case 3
                muSettings.IlluminationMap = EngineSettings.eTextureResOptions.eNormResTextures
            Case 4
                muSettings.IlluminationMap = EngineSettings.eTextureResOptions.eHiResTextures
            Case Else
                muSettings.IlluminationMap = -1
        End Select
        'Reload the illumination maps
        If moMeshIllum Is Nothing = False Then moMeshIllum.Dispose()
        moMeshIllum = Nothing
        mbDataChanged = True
    End Sub

    Private Sub cboLightQuality_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLightQuality.SelectedIndexChanged
        If mbLoading = True Then Return
        'Low, Medium, High
        Select Case cboLightQuality.SelectedIndex
            Case 1
                muSettings.LightQuality = EngineSettings.LightQualitySetting.PerPixel
            Case 2
                muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap
            Case Else
                muSettings.LightQuality = EngineSettings.LightQualitySetting.VSPS1
        End Select
        'Load required resources
        mbDataChanged = True
    End Sub

    Private Sub hscrClipPlane_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscrClipPlane.Scroll
        If mbLoading = True Then Return
        muSettings.EntityClipPlane = hscrClipPlane.Value * 100
        mbDataChanged = True
    End Sub

    Private Sub btnResetInterface_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetInterface.Click
        If mbLoading = True Then Return
        With muSettings
            .InterfaceBorderColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .InterfaceFillColor = System.Drawing.Color.FromArgb(128, 32, 64, 128)
            .InterfaceTextBoxFillColor = System.Drawing.Color.FromArgb(255, 32, 64, 92)
            .InterfaceTextBoxForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .InterfaceButtonColor = System.Drawing.Color.FromArgb(255, 192, 220, 255)
        End With
        mbDataChanged = True
        btnResetInterface.Enabled = False
    End Sub

    Private Sub chkVerticalSync_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkVerticalSync.CheckedChanged
        If mbLoading = True Then Return
        muSettings.VSync = chkVerticalSync.Checked
        mbDataChanged = True
    End Sub

    Private Sub chkTripleBuffer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTripleBuffer.CheckedChanged
        If mbLoading = True Then Return
        muSettings.TripleBuffer = chkTripleBuffer.Checked
        mbDataChanged = True
    End Sub

    Private Sub chkFullScreen_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFullScreen.CheckedChanged
        If mbLoading = True Then Return
        muSettings.Windowed = Not chkFullScreen.Checked
        mbDataChanged = True
    End Sub

    Private Sub cboFullScreenRes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFullScreenRes.SelectedIndexChanged
        If mbLoading = True Then Return
        'cFmt.Width & "x" & cFmt.Height & " " & cFmt.RefreshRate & "hz"
        Dim sTemp As String = cboFullScreenRes.Text
        If sTemp Is Nothing Then Return

        Dim lXLoc As Int32 = sTemp.IndexOf("x"c)
        If lXLoc > -1 Then
            Dim sWidth As String = sTemp.Substring(0, lXLoc) '- 1)
            sTemp = sTemp.Substring(lXLoc + 1)
            lXLoc = sTemp.IndexOf(" "c)
            If lXLoc > -1 Then
                Dim sHeight As String = sTemp.Substring(0, lXLoc) ' - 1)
                Dim sHz As String = sTemp.Substring(lXLoc + 1)
                sHz = sHz.Replace("hz", "")

                muSettings.FullScreenResX = CInt(sWidth)
                muSettings.FullScreenResY = CInt(sHeight)
                muSettings.FullScreenRefreshRate = CInt(sHz)
                mbDataChanged = True
            End If
        End If
    End Sub

    Private Sub chkDeepCosmos_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDeepCosmos.CheckedChanged
        If mbLoading = True Then Return
        muSettings.RenderCosmos = chkDeepCosmos.Checked
        mbDataChanged = True
    End Sub

    Private Sub hscrStarPlanet_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscrStarPlanet.Scroll
        If mbLoading = True Then Return
        muSettings.StarfieldParticlesPlanet = hscrStarPlanet.Value * 10
        mbDataChanged = True
    End Sub

    Private Sub hscrStarSpace_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscrStarSpace.Scroll
        If mbLoading = True Then Return
        muSettings.StarfieldParticlesSpace = hscrStarSpace.Value * 100
        mbDataChanged = True
    End Sub

    Private Sub chkShieldFX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShieldFX.CheckedChanged
        If mbLoading = True Then Return
        muSettings.RenderShieldFX = chkShieldFX.Checked
        mbDataChanged = True
    End Sub

    Private Sub cboPlanetFX_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPlanetFX.SelectedIndexChanged
        If mbLoading = True Then Return
        'Off, Low, Medium, High, Full
        muSettings.PlanetFXParticles = cboPlanetFX.SelectedIndex
        mbDataChanged = True
    End Sub

    Private Sub cboBurnFX_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboBurnFX.SelectedIndexChanged
        If mbLoading = True Then Return
        'Off, Low, Medium, High, Full
        muSettings.BurnFXParticles = cboBurnFX.SelectedIndex
        mbDataChanged = True
    End Sub

    Private Sub cboEngineFX_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEngineFX.SelectedIndexChanged
        If mbLoading = True Then Return
        'Off, Low, Medium, High, Full
        muSettings.EngineFXParticles = cboEngineFX.SelectedIndex
        mbDataChanged = True
    End Sub

    Private Sub btnDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefaults.Click
        'reset all defaults

        With muSettings
            .Windowed = True
            .VertexProcessing = CreateFlags.HardwareVertexProcessing
            .TripleBuffer = False
            .VSync = False
            .FullScreenResX = -1
            .FullScreenResY = -1
            .FullScreenRefreshRate = -1
            .RenderHPBars = True
            .RenderBurnMarks = True

            .LightQuality = EngineSettings.LightQualitySetting.BumpMap
            .RenderPlayerCustomization = True
            .IlluminationMap = 512
            .RenderCosmos = True
            .StarfieldParticlesPlanet = 500
            .StarfieldParticlesSpace = 10000
            .BumpMapTerrain = False
            .BumpMapPlanetModel = False
            .IlluminationMapTerrain = False

            .SmoothFOW = True
            .WaterRenderMethod = 1
            .WaterTextureRes = 256
            .EntityClipPlane = 20000
            .ModelTextureResolution = EngineSettings.eTextureResOptions.eNormResTextures
            .ShowFOWTerrainShading = True
            .FOWTextureResolution = 512
            .TerrainTextureResolution = 3
            .PlanetModelTextureWH = 256

            .EngineFXParticles = 4
            .BurnFXParticles = 4
            .RenderShieldFX = True
            .PlanetFXParticles = 4
            .PostGlowAmt = 0.5F
            .HiResPlanetTexture = True

            .InterfaceBorderColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .InterfaceFillColor = System.Drawing.Color.FromArgb(128, 32, 64, 128)
            .InterfaceTextBoxFillColor = System.Drawing.Color.FromArgb(255, 32, 64, 92)
            .InterfaceTextBoxForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .InterfaceButtonColor = System.Drawing.Color.FromArgb(255, 192, 220, 255)
        End With

        mbDataChanged = True
    End Sub

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

    End Sub
End Class

Public Class InitFile
    ' API functions
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
      (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    Private Declare Ansi Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
    Private strFilename As String

    ' Constructor, accepting a filename
    Public Sub New(Optional ByVal Filename As String = "")
        If Filename = "" Then
            'Ok, use the app.path
            strFilename = System.AppDomain.CurrentDomain.BaseDirectory()
            If Right$(strFilename, 1) <> "\" Then strFilename = strFilename & "\"
            strFilename = strFilename & Replace$(System.AppDomain.CurrentDomain.FriendlyName().ToLower, ".exe", ".ini")
        Else
            strFilename = Filename
        End If
    End Sub

    ' Read-only filename property
    ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property

    Public Function GetString(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As String) As String
        ' Returns a string from your INI file
        Dim intCharCount As Integer
        Dim objResult As New System.Text.StringBuilder(2048)
        intCharCount = GetPrivateProfileString(Section, Key, _
           [Default], objResult, objResult.Capacity, strFilename)
        If intCharCount > 0 Then Return Left(objResult.ToString, intCharCount) Else Return ""
    End Function

    Public Function GetInteger(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Integer) As Integer
        ' Returns an integer from your INI file
        Return GetPrivateProfileInt(Section, Key, _
           [Default], strFilename)
    End Function

    Public Function GetBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal [Default] As Boolean) As Boolean
        ' Returns a boolean from your INI file
        Return (GetPrivateProfileInt(Section, Key, _
           CInt([Default]), strFilename) = 1)
    End Function

    Public Sub WriteString(ByVal Section As String, _
      ByVal Key As String, ByVal Value As String)
        ' Writes a string to your INI file
        WritePrivateProfileString(Section, Key, Value, strFilename)
        Flush()
    End Sub

    Public Sub WriteInteger(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Integer)
        ' Writes an integer to your INI file
        WriteString(Section, Key, CStr(Value))
        Flush()
    End Sub

    Public Sub WriteBoolean(ByVal Section As String, _
      ByVal Key As String, ByVal Value As Boolean)
        ' Writes a boolean to your INI file
        WriteString(Section, Key, CStr(CInt(Value)))
        Flush()
    End Sub

    Private Sub Flush()
        ' Stores all the cached changes to your INI file
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub

End Class

Public Module GlobalVars
    Public goCamera As Camera
    Public moResMgr As EpicaResourceManager

    Public Function Exists(ByVal sFilename As String) As Boolean
        If Trim(sFilename).Length > 0 Then
            On Error Resume Next
            sFilename = Dir$(sFilename, FileAttribute.Archive Or FileAttribute.Directory Or FileAttribute.Hidden Or FileAttribute.Normal Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
            Return Err.Number = 0 And sFilename.Length > 0
        Else
            Return False
        End If

    End Function
#Region " Trigonometric/Geometric Functions "
    'For Geometric Calculations
    Public Const gdPi As Single = 3.14159265358979
    Public Const gdHalfPie As Single = gdPi / 2.0F
    Public Const gdPieAndAHalf As Single = gdPi * 1.5F
    Public Const gdTwoPie As Single = gdPi * 2.0F
    Public Const gdDegreePerRad As Single = 180.0F / gdPi
    Public Const gdRadPerDegree As Single = Math.PI / 180.0F
    Public Function RadianToDegree(ByVal fRads As Single) As Single
        Return fRads * gdDegreePerRad
    End Function

    Public Function DegreeToRadian(ByVal fDegree As Single) As Single
        Return fDegree * gdRadPerDegree
    End Function

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Int32, ByRef lEndY As Int32, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * CSng(Math.PI / 180.0F)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CInt((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CInt((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Sub RotatePoint(ByVal lAxisX As Int32, ByVal lAxisY As Int32, ByRef lEndX As Single, ByRef lEndY As Single, ByVal dDegree As Single)
        Dim dDX As Single
        Dim dDY As Single
        Dim dRads As Single

        dRads = dDegree * CSng(Math.PI / 180.0F)
        dDX = lEndX - lAxisX
        dDY = lEndY - lAxisY
        lEndX = lAxisX + CSng((dDX * Math.Cos(dRads)) + (dDY * Math.Sin(dRads)))
        lEndY = lAxisY + -CSng((dDX * Math.Sin(dRads)) - (dDY * Math.Cos(dRads)))
    End Sub

    Public Function LineAngleDegrees(ByVal lX1 As Int32, ByVal lY1 As Int32, ByVal lX2 As Int32, ByVal lY2 As Int32) As Single
        Dim dDeltaX As Single
        Dim dDeltaY As Single
        Dim dAngle As Single

        dDeltaX = lX2 - lX1
        dDeltaY = lY2 - lY1

        If dDeltaX = 0 Then     'vertical
            If dDeltaY < 0 Then
                dAngle = gdHalfPie
            Else
                dAngle = gdPieAndAHalf
            End If
        ElseIf dDeltaY = 0 Then     'horizontal
            If dDeltaX < 0 Then
                dAngle = gdPi
            Else
                dAngle = 0
            End If
        Else    'angled
            dAngle = CSng(System.Math.Atan(System.Math.Abs(dDeltaY / dDeltaX)))
            'Correct for VB's reversed Y... VB Upper Right is ok... but the other quads are not
            If dDeltaX > -1 And dDeltaY > -1 Then       'VB Lower Right
                dAngle = gdTwoPie - dAngle
            ElseIf dDeltaX < 0 And dDeltaY > -1 Then    'VB Lower Left
                dAngle = gdPi + dAngle
            ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left
                dAngle = gdPi - dAngle
            End If
        End If

        Return CInt(dAngle * gdDegreePerRad)

    End Function

    Public Function Distance(ByVal lX1 As Int32, ByVal lZ1 As Int32, ByVal lX2 As Int32, ByVal lZ2 As Int32) As Single
        Dim fX As Single = lX2 - lX1
        Dim fZ As Single = lZ2 - lZ1
        fX *= fX
        fZ *= fZ
        Return CSng(Math.Sqrt(fX + fZ))
    End Function

#End Region

#Region " Settings Management "
    Public Structure EngineSettings
        Public Enum LightQualitySetting As Int32
            VSPS1 = 0
            PerPixel = 1
            BumpMap = 2
        End Enum

#Region "  Graphics Configuration  "
        Public AdapterOrdinal As Int32
        Public BumpMapTerrain As Boolean
        Public BumpMapPlanetModel As Boolean
        Public BurnFXParticles As Int32
        Public DepthBuffer As Int32
        Public Dither As Boolean
        Public DrawGrid As Boolean
        Public EngineFXParticles As Int32
        Public EntityClipPlane As Int32
        Public FarClippingPlane As Int32
        Public FullScreenRefreshRate As Int32
        Public FullScreenResX As Int32
        Public FullScreenResY As Int32
        Public FOWTextureResolution As Int32
        Public HiResPlanetTexture As Boolean
        Public IlluminationMap As Int32
        Public IlluminationMapTerrain As Boolean
        Public LightQuality As LightQualitySetting
        Public MaxLights As Int32
        Public ModelTextureResolution As eTextureResOptions
        Public NearClippingPlane As Int32
        Public PlanetFXParticles As Int32
        Public PlanetModelTextureWH As Int32
        Public PostGlowAmt As Single
        Public RenderCosmos As Boolean
        Public RenderMineralCaches As Boolean
        Public RenderPlayerCustomization As Boolean
        Public RenderShieldFX As Boolean
        Public ScreenShakeEnabled As Boolean
        Public ShowFOWTerrainShading As Boolean
        Public ShowMiniMap As Boolean
        Public SmoothFOW As Boolean
        Public SpecularEnabled As Boolean
        Public StarfieldParticlesPlanet As Int32
        Public StarfieldParticlesSpace As Int32
        Public TerrainFarClippingPlane As Int32
        Public TerrainNearClippingPlane As Int32
        Public TerrainTextureResolution As Int32
        Public TripleBuffer As Boolean
        Public SavedTripleBuffer As Boolean
        Public VertexProcessing As Int32
        Public VSync As Boolean
        Public SavedVSync As Boolean
        Public WaterRenderMethod As Int32
        Public WaterTextureRes As Int32
        Public Windowed As Boolean
#End Region

#Region "  User Interface Specifics  "
        Public InterfaceBorderColor As System.Drawing.Color '= System.Drawing.Color.FromArgb(255, 255, 255, 255)
        Public InterfaceButtonColor As System.Drawing.Color
        Public InterfaceFillColor As System.Drawing.Color '= System.Drawing.Color.FromArgb(128, 32, 64, 128)
        Public InterfaceTextBoxFillColor As System.Drawing.Color
        Public InterfaceTextBoxForeColor As System.Drawing.Color
        Public NotificationDisplayTime As Int32

        Public FilterBadWords As Boolean
#End Region

#Region "  Window Position and Other Settings  "
        Public AdvanceDisplayLocX As Int32
        Public AdvanceDisplayLocY As Int32
        Public BehaviorLocX As Int32
        Public BehaviorLocY As Int32
        Public BuildWindowLeft As Int32
        Public BuildWindowTop As Int32
        Public ChatWindowLocX As Int32
        Public ChatWindowLocY As Int32
        Public ContentsLocX As Int32
        Public ContentsLocY As Int32
        Public EnvirDisplayLocX As Int32
        Public EnvirDisplayLocY As Int32
        Public MiniMapLocX As Int32
        Public MiniMapLocY As Int32
        Public MiniMapWidthHeight As Int32
        Public PlanetMinimapZoomLevel As Byte
        Public ProdStatusLocX As Int32
        Public ProdStatusLocY As Int32
        Public ResearchWindowLeft As Int32
        Public ResearchWindowTop As Int32
        Public SelectLocX As Int32
        Public SelectLocY As Int32
#End Region

#Region "  Chat Window Colors  "
        Public DefaultChatColor As System.Drawing.Color
        Public AlertChatColor As System.Drawing.Color
        Public StatusChatColor As System.Drawing.Color
        Public LocalChatColor As System.Drawing.Color
        Public GuildChatColor As System.Drawing.Color
        Public SenateChatColor As System.Drawing.Color
        Public PMChatColor As System.Drawing.Color
        Public ChannelChatColor As System.Drawing.Color
        'Public Ch2ChatColor As System.Drawing.Color
        'Public Ch3ChatColor As System.Drawing.Color
        'Public Ch4ChatColor As System.Drawing.Color
        'Public Ch5ChatColor As System.Drawing.Color
        'Public Ch6ChatColor As System.Drawing.Color
        'Public Ch7ChatColor As System.Drawing.Color
        'Public Ch8ChatColor As System.Drawing.Color
        'Public Ch9ChatColor As System.Drawing.Color
        'Public Ch10ChatColor As System.Drawing.Color
#End Region

#Region "  Planet Type Color Config  "
        Public AcidBuildGhost As System.Drawing.Color
        Public AcidMineralCache As System.Drawing.Color
        Public AcidMinimapAngle As System.Drawing.Color
        Public AdaptableBuildGhost As System.Drawing.Color
        Public AdaptableMineralCache As System.Drawing.Color
        Public AdaptableMinimapAngle As System.Drawing.Color
        Public BarrenBuildGhost As System.Drawing.Color
        Public BarrenMineralCache As System.Drawing.Color
        Public BarrenMinimapAngle As System.Drawing.Color
        Public DesertBuildGhost As System.Drawing.Color
        Public DesertMineralCache As System.Drawing.Color
        Public DesertMinimapAngle As System.Drawing.Color
        Public IceBuildGhost As System.Drawing.Color
        Public IceMineralCache As System.Drawing.Color
        Public IceMinimapAngle As System.Drawing.Color
        Public LavaBuildGhost As System.Drawing.Color
        Public LavaMineralCache As System.Drawing.Color
        Public LavaMinimapAngle As System.Drawing.Color
        Public TerranBuildGhost As System.Drawing.Color
        Public TerranMineralCache As System.Drawing.Color
        Public TerranMinimapAngle As System.Drawing.Color
        Public WaterworldBuildGhost As System.Drawing.Color
        Public WaterworldMineralCache As System.Drawing.Color
        Public WaterworldMinimapAngle As System.Drawing.Color
#End Region

#Region "  Asset Colorization  "
        Public NeutralAssetColor As Microsoft.DirectX.Vector4
        Public MyAssetColor As Microsoft.DirectX.Vector4
        Public EnemyAssetColor As Microsoft.DirectX.Vector4
        Public AllyAssetColor As Microsoft.DirectX.Vector4
        Public GuildAssetColor As Microsoft.DirectX.Vector4
#End Region

        Public AudioOn As Boolean
        Public bRanBefore As Boolean

        Public ShowIntro As Boolean

        Public CtrlQExits As Boolean
        Public ShowTargetBoxes As Boolean
        Public RenderHPBars As Boolean
        Public RenderBurnMarks As Boolean

        Public bDoNotShowEngineerCancelAlert As Boolean

        Public lCurrentResearchFilter As Int32          'NOT SAVED OR LOADED
        Public lPlanetViewCameraX As Int32              'NOT SAVED OR LOADED
        Public lPlanetViewCameraY As Int32              'NOT SAVED OR LOADED
        Public lPlanetViewCameraZ As Int32              'NOT SAVED OR LOADED
        Public yCurrentContentsView As Byte             'NOT SAVED OR LOADED
        Public ExpandedColonyStatsScreen As Boolean     'Not saved or loaded

        Public Enum eTextureResOptions As Integer
            eHiResTextures = 512
            eNormResTextures = 256
            eLowResTextures = 128
            eVeryLowResTextures = 64
        End Enum

        Public Sub LoadSettings()
            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            Dim oINI As New InitFile(sPath & "BPClient.ini")

            Dim lA As Int32
            Dim lR As Int32
            Dim lG As Int32
            Dim lB As Int32

            AdapterOrdinal = CInt(oINI.GetString("GRAPHICS", "AdapterOrdinal", "0"))
            Windowed = CBool(Val(oINI.GetString("GRAPHICS", "Windowed", "1")) <> 0)
            VertexProcessing = CInt(Val(oINI.GetString("GRAPHICS", "VertexProcessing", "64")))   'hardware
            TripleBuffer = CBool(Val(oINI.GetString("GRAPHICS", "TripleBuffer", "0")) <> 0)
            SavedTripleBuffer = TripleBuffer
            VSync = CBool(Val(oINI.GetString("GRAPHICS", "VSync", "0")) <> 0)
            SavedVSync = VSync
            FullScreenResX = CInt(oINI.GetString("GRAPHICS", "FullScreenResX", "-1"))
            FullScreenResY = CInt(oINI.GetString("GRAPHICS", "FullScreenResY", "-1"))
            FullScreenRefreshRate = CInt(oINI.GetString("GRAPHICS", "FullScreenRefreshRate", "-1"))

            FilterBadWords = CInt(Val(oINI.GetString("INTERFACE", "FilterBadWords", "1"))) <> 0
            RenderHPBars = CInt(Val(oINI.GetString("SETTINGS", "RenderHPBars", "1"))) <> 0
            RenderBurnMarks = CInt(Val(oINI.GetString("GRAPHICS", "RenderBurnMarks", "1"))) <> 0

            bDoNotShowEngineerCancelAlert = CInt(Val(oINI.GetString("INTERFACE", "DoNotShowEngineerCancelAlert", "0"))) <> 0
            CtrlQExits = CInt(Val(oINI.GetString("SETTINGS", "CtrlQExits", "1"))) <> 0

            LightQuality = CType(CInt(Val(oINI.GetString("LIGHTING", "Quality", CInt(LightQualitySetting.BumpMap).ToString))), LightQualitySetting)
            RenderPlayerCustomization = CBool(CInt(Val(oINI.GetString("GRAPHICS", "PlayerCustomization", "1"))) <> 0)
            IlluminationMap = Math.Min(Math.Max(64, CInt(Val(oINI.GetString("LIGHTING", "IlluminationMap", "512")))), 1024)
            RenderCosmos = CInt(Val(oINI.GetString("GRAPHICS", "RenderCosmos", "1"))) <> 0
            StarfieldParticlesPlanet = Math.Max(100, Math.Min(4000, CInt(Val(oINI.GetString("GRAPHICS", "StarfieldParticlesPlanet", 500.ToString)))))
            StarfieldParticlesSpace = Math.Max(1000, Math.Min(40000, CInt(Val(oINI.GetString("GRAPHICS", "StarfieldParticlesSpace", 10000.ToString)))))
            BumpMapTerrain = CInt(Val(oINI.GetString("LIGHTING", "BumpMapTerrain", "0"))) <> 0
            BumpMapPlanetModel = CInt(Val(oINI.GetString("LIGHTING", "BumpMapPlanetModel", "0"))) <> 0
            IlluminationMapTerrain = CInt(Val(oINI.GetString("LIGHTING", "IllumMapTerrain", "0"))) <> 0

            'NeutralAssetColor = New Microsoft.DirectX.Vector4(1, 1, 1, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "NeutralR", "255")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "NeutralG", "255")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "NeutralB", "255")))
            lA = 0
            NeutralAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'MyAssetColor = New Microsoft.DirectX.Vector4(0, 1, 0, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "PlayerR", "0")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "PlayerG", "255")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "PlayerB", "0")))
            lA = 0
            MyAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'EnemyAssetColor = New Microsoft.DirectX.Vector4(0, 1, 0, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "EnemyR", "255")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "EnemyG", "0")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "EnemyB", "0")))
            lA = 0
            EnemyAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'AllyAssetColor = New Microsoft.DirectX.Vector4(0, 1, 1, 0)
            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "AllyR", "0")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "AllyG", "255")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "AllyB", "255")))
            lA = 0
            AllyAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            lR = CInt(Val(oINI.GetString("IDENTIFICATION", "GuildR", "192")))
            lG = CInt(Val(oINI.GetString("IDENTIFICATION", "GuildG", "32")))
            lB = CInt(Val(oINI.GetString("IDENTIFICATION", "GuildB", "128")))
            lA = 0
            GuildAssetColor = New Microsoft.DirectX.Vector4(lR / 255.0F, lG / 255.0F, lB / 255.0F, lA / 255.0F)

            'DefaultChatColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "DefaultB", "255")))
            DefaultChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'AlertChatColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "AlertR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "AlertG", "0")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "AlertB", "0")))
            AlertChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'StatusChatColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "StatusR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "StatusG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "StatusB", "0")))
            StatusChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'LocalChatColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "LocalR", "0")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "LocalG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "LocalB", "255")))
            LocalChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'GuildChatColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "GuildR", "0")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "GuildG", "255")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "GuildB", "0")))
            GuildChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'PMChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 0)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "PMR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "PMG", "128")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "PMB", "0")))
            PMChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'SenateChatColor = System.Drawing.Color.FromArgb(255, 192, 192, 192)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "SenateR", "192")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "SenateG", "192")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "SenateB", "192")))
            SenateChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'ChannelChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            lR = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelR", "255")))
            lG = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelG", "128")))
            lB = CInt(Val(oINI.GetString("CHATCOLOR", "ChannelB", "255")))
            ChannelChatColor = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            'Ch2ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch3ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch4ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch5ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch6ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch7ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch8ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch9ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)
            'Ch10ChatColor = System.Drawing.Color.FromArgb(255, 255, 128, 255)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "AcidR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "AcidG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "AcidB", "255")))
            AcidBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "AcidR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "AcidG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "AcidB", "255")))
            AcidMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "AcidR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "AcidG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "AcidB", "255")))
            AcidMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "AdaptableR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "AdaptableG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "AdaptableB", "255")))
            AdaptableBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "AdaptableR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "AdaptableG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "AdaptableB", "255")))
            AdaptableMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "AdaptableR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "AdaptableG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "AdaptableB", "255")))
            AdaptableMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "BarrenR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "BarrenG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "BarrenB", "255")))
            BarrenBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "BarrenR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "BarrenG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "BarrenB", "255")))
            BarrenMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "BarrenR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "BarrenG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "BarrenB", "255")))
            BarrenMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "DesertR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "DesertG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "DesertB", "255")))
            DesertBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "DesertR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "DesertG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "DesertB", "255")))
            DesertMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "DesertR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "DesertG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "DesertB", "255")))
            DesertMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "IceR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "IceG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "IceB", "255")))
            IceBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "IceR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "IceG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "IceB", "255")))
            IceMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "IceR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "IceG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "IceB", "255")))
            IceMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "LavaR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "LavaG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "LavaB", "255")))
            LavaBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "LavaR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "LavaG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "LavaB", "255")))
            LavaMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "LavaR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "LavaG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "LavaB", "255")))
            LavaMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "TerranR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "TerranG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "TerranB", "255")))
            TerranBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "TerranR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "TerranG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "TerranB", "255")))
            TerranMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "TerranR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "TerranG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "TerranB", "255")))
            TerranMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)

            lR = CInt(Val(oINI.GetString("BUILDGHOST", "WaterworldR", "255")))
            lG = CInt(Val(oINI.GetString("BUILDGHOST", "WaterworldG", "255")))
            lB = CInt(Val(oINI.GetString("BUILDGHOST", "WaterworldB", "255")))
            WaterworldBuildGhost = System.Drawing.Color.FromArgb(64, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINERALCACHE", "WaterworldR", "255")))
            lG = CInt(Val(oINI.GetString("MINERALCACHE", "WaterworldG", "255")))
            lB = CInt(Val(oINI.GetString("MINERALCACHE", "WaterworldB", "255")))
            WaterworldMineralCache = System.Drawing.Color.FromArgb(255, lR, lG, lB)
            lR = CInt(Val(oINI.GetString("MINIMAPANGLE", "WaterworldR", "255")))
            lG = CInt(Val(oINI.GetString("MINIMAPANGLE", "WaterworldG", "255")))
            lB = CInt(Val(oINI.GetString("MINIMAPANGLE", "WaterworldB", "255")))
            WaterworldMinimapAngle = System.Drawing.Color.FromArgb(255, lR, lG, lB)


            'ALWAYS SET TO THIS ON EVERY STARTUP
            lCurrentResearchFilter = -1
            lPlanetViewCameraX = 0
            lPlanetViewCameraY = 1700
            lPlanetViewCameraZ = 1000
            ExpandedColonyStatsScreen = True
            '====== end of same startup =======

            ShowIntro = CBool(Val(oINI.GetString("SETTINGS", "ShowIntroScreen", "1")) <> 0)
            ShowTargetBoxes = CBool(Val(oINI.GetString("SETTINGS", "ShowTargetBoxes", "1")) <> 0)
            SmoothFOW = CBool(Val(oINI.GetString("GRAPHICS", "SmoothFOW", "1")) <> 0)
            WaterRenderMethod = CInt(Val(oINI.GetString("GRAPHICS", "WaterRenderMethod", "1")))
            WaterTextureRes = Math.Min(256, Math.Max(64, CInt(Val(oINI.GetString("GRAPHICS", "WaterTextureRes", "256")))))
            PlanetMinimapZoomLevel = CByte(Math.Min(3, Math.Max(0, Val(oINI.GetString("SETTINGS", "PlanetMinimapZoomLevel", "3")))))
            MaxLights = CInt(Val(oINI.GetString("GRAPHICS", "MaxLights", "0")))
            FarClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "FarClippingPlane", "1000000000")))
            NearClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "NearClippingPlane", "1")))
            TerrainNearClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "TerrainNearClippingPlane", "100")))
            TerrainFarClippingPlane = CInt(Val(oINI.GetString("GRAPHICS", "TerrainFarClippingPlane", "25000")))
            ShowMiniMap = CBool(Val(oINI.GetString("PLANET_VIEW", "ShowMiniMap", "1")))
            EntityClipPlane = CInt(Val(oINI.GetString("GRAPHICS", "EntityClipPlane", "20000")))
            ModelTextureResolution = CType((Math.Min(512, Math.Max(64, Val(oINI.GetString("GRAPHICS", "ModelTextureResolution", eTextureResOptions.eNormResTextures.ToString()))))), eTextureResOptions)
            If CInt(ModelTextureResolution) = 0 Then ModelTextureResolution = eTextureResOptions.eNormResTextures
            DrawGrid = CBool(Val(oINI.GetString("GRAPHICS", "DrawGrid", "0")))
            'DrawProceduralWater = CBool(Val(oINI.GetString("GRAPHICS", "DrawProceduralWater", "0")))
            ShowFOWTerrainShading = CBool(Val(oINI.GetString("GRAPHICS", "ShowFOWTerrainShading", "1")))
            FOWTextureResolution = CInt(Val(oINI.GetString("GRAPHICS", "FOWTextureResolution", "512")))
            If FOWTextureResolution < 256 Then FOWTextureResolution = 256
            MiniMapLocX = CInt(Val(oINI.GetString("HUD", "MiniMapLocX", "0")))
            MiniMapLocY = CInt(Val(oINI.GetString("HUD", "MiniMapLocY", "0")))
            MiniMapWidthHeight = CInt(Val(oINI.GetString("HUD", "MiniMapWidthHeight", "120")))
            RenderMineralCaches = CBool(Val(oINI.GetString("GRAPHICS", "RenderMineralCaches", "1")))
            Dither = CBool(Val(oINI.GetString("GRAPHICS", "DITHER", "0")))
            SpecularEnabled = CBool(Val(oINI.GetString("GRAPHICS", "SpecularEnabled", "1")))
            NotificationDisplayTime = CInt(Val(oINI.GetString("SETTINGS", "NotificationDisplayTime", "5000")))
            AudioOn = CBool(Val(oINI.GetString("AUDIO", "AudioEnabled", "1")))
            'WaterQuadWidthHeight = CInt(Val(oINI.GetString("GRAPHICS", "WaterQuadRes", "32")))
            ScreenShakeEnabled = CBool(Val(oINI.GetString("SETTINGS", "ScreenShakeEnable", "1")) <> 0)
            TerrainTextureResolution = Math.Max(1, CInt(Val(oINI.GetString("GRAPHICS", "TerrainTextureResolution", "3"))))
            PlanetModelTextureWH = Math.Max(64, CInt(Val(oINI.GetString("GRAPHICS", "PlanetModelTextureResolution", "256"))))

            BuildWindowLeft = CInt(Val(oINI.GetString("HUD", "BuildWindowLeft", "-1")))
            BuildWindowTop = CInt(Val(oINI.GetString("HUD", "BuildWindowTop", "-1")))
            ResearchWindowLeft = CInt(Val(oINI.GetString("HUD", "ResearchWindowLeft", "-1")))
            ResearchWindowTop = CInt(Val(oINI.GetString("HUD", "ResearchWindowTop", "-1")))

            EngineFXParticles = CInt(Val(oINI.GetString("GRAPHICS", "EngineFXParticles", "4")))
            If EngineFXParticles > 4 OrElse EngineFXParticles < 0 Then EngineFXParticles = 4
            BurnFXParticles = CInt(Val(oINI.GetString("GRAPHICS", "BurnFXParticles", "4")))
            If BurnFXParticles > 4 OrElse BurnFXParticles < 0 Then BurnFXParticles = 4
            RenderShieldFX = CBool(Val(oINI.GetString("GRAPHICS", "RenderShieldFX", "1")) <> 0)
            'StarfieldParticlesSpace = CInt(Val(oINI.GetString("GRAPHICS", "StarfieldParticles", "10000")))
            If StarfieldParticlesSpace < 1000 Then StarfieldParticlesSpace = 1000
            If StarfieldParticlesSpace > 40000 Then StarfieldParticlesSpace = 40000
            PlanetFXParticles = CInt(Val(oINI.GetString("GRAPHICS", "PlanetFXParticles", "4")))
            If PlanetFXParticles > 4 OrElse PlanetFXParticles < 0 Then PlanetFXParticles = 4

            PostGlowAmt = CSng(Val(oINI.GetString("GRAPHICS", "PostGlowAmount", "0.5")))
            HiResPlanetTexture = CBool(Val(oINI.GetString("GRAPHICS", "HiResPlanetTexture", "1")) <> 0)

            bRanBefore = CBool(Val(oINI.GetString("SETTINGS", "RanBefore", "0")) <> 0)

            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_R", "255"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_G", "255"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "BorderColor_B", "255"))), 0), 255)
            InterfaceBorderColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_A", "128"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_R", "32"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_G", "64"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "FillColor_B", "128"))), 0), 255)
            InterfaceFillColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_R", "32"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_G", "64"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxFillColor_B", "92"))), 0), 255)
            InterfaceTextBoxFillColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_R", "255"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_G", "255"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "TextBoxForeColor_B", "255"))), 0), 255)
            InterfaceTextBoxForeColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)
            lA = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_A", "255"))), 0), 255)
            lR = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_R", "192"))), 0), 255)
            lG = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_G", "220"))), 0), 255)
            lB = Math.Min(Math.Max(CInt(Val(oINI.GetString("INTERFACE", "ButtonColor_B", "255"))), 0), 255)
            InterfaceButtonColor = System.Drawing.Color.FromArgb(lA, lR, lG, lB)

            'Hud Position Settings
            AdvanceDisplayLocX = CInt(Val(oINI.GetString("INTERFACE", "AdvanceDisplayLocX", "-1")))
            AdvanceDisplayLocY = CInt(Val(oINI.GetString("INTERFACE", "AdvanceDisplayLocY", "-1")))
            BehaviorLocX = CInt(Val(oINI.GetString("INTERFACE", "BehaviorLocX", "-1")))
            BehaviorLocY = CInt(Val(oINI.GetString("INTERFACE", "BehaviorLocY", "-1")))
            ContentsLocX = CInt(Val(oINI.GetString("INTERFACE", "ContentsLocX", "-1")))
            ContentsLocY = CInt(Val(oINI.GetString("INTERFACE", "ContentsLocY", "-1")))
            EnvirDisplayLocX = CInt(Val(oINI.GetString("INTERFACE", "EnvirDisplayLocX", "-1")))
            EnvirDisplayLocY = CInt(Val(oINI.GetString("INTERFACE", "EnvirDisplayLocY", "-1")))
            SelectLocX = CInt(Val(oINI.GetString("INTERFACE", "SelectLocX", "-1")))
            SelectLocY = CInt(Val(oINI.GetString("INTERFACE", "SelectLocY", "-1")))
            ProdStatusLocX = CInt(Val(oINI.GetString("INTERFACE", "ProdStatusLocX", "-1")))
            ProdStatusLocY = CInt(Val(oINI.GetString("INTERFACE", "ProdStatusLocY", "-1")))
            ChatWindowLocX = CInt(Val(oINI.GetString("INTERFACE", "ChatWindowLocX", "-1")))
            ChatWindowLocY = CInt(Val(oINI.GetString("INTERFACE", "ChatWindowLocY", "-1")))
            'End of HUD Position settings

            oINI = Nothing
        End Sub

        Public Sub SaveSettings()
            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            Dim oINI As New InitFile(sPath & "BPClient.ini")


            oINI.WriteString("GRAPHICS", "AdapterOrdinal", AdapterOrdinal.ToString)
            If Windowed = True Then
                oINI.WriteString("GRAPHICS", "Windowed", "1")
            Else : oINI.WriteString("GRAPHICS", "Windowed", "0")
            End If
            oINI.WriteString("GRAPHICS", "VertexProcessing", VertexProcessing.ToString)
            If SavedTripleBuffer = True Then
                oINI.WriteString("GRAPHICS", "TripleBuffer", "1")
            Else : oINI.WriteString("GRAPHICS", "TripleBuffer", "0")
            End If
            If SavedVSync = True Then
                oINI.WriteString("GRAPHICS", "VSync", "1")
            Else : oINI.WriteString("GRAPHICS", "VSync", "0")
            End If
            oINI.WriteString("GRAPHICS", "FullScreenResX", FullScreenResX.ToString)
            oINI.WriteString("GRAPHICS", "FullScreenResY", FullScreenResY.ToString)
            oINI.WriteString("GRAPHICS", "FullScreenRefreshRate", FullScreenRefreshRate.ToString)

            If RenderBurnMarks = True Then
                oINI.WriteString("GRAPHICS", "RenderBurnMarks", "1")
            Else : oINI.WriteString("GRAPHICS", "RenderBurnMarks", "0")
            End If

            oINI.WriteString("LIGHTING", "Quality", CInt(LightQuality).ToString)
            If RenderPlayerCustomization = True Then
                oINI.WriteString("GRAPHICS", "PlayerCustomization", "1")
            Else : oINI.WriteString("GRAPHICS", "PlayerCustomization", "0")
            End If
            oINI.WriteString("LIGHTING", "IlluminationMap", IlluminationMap.ToString)
            If RenderCosmos = True Then
                oINI.WriteString("GRAPHICS", "RenderCosmos", "1")
            Else : oINI.WriteString("GRAPHICS", "RenderCosmos", "0")
            End If
            oINI.WriteString("GRAPHICS", "StarfieldParticlesPlanet", StarfieldParticlesPlanet.ToString)
            oINI.WriteString("GRAPHICS", "StarfieldParticlesSpace", StarfieldParticlesSpace.ToString)
            If BumpMapTerrain = True Then
                oINI.WriteString("LIGHTING", "BumpMapTerrain", "1")
            Else : oINI.WriteString("LIGHTING", "BumpMapTerrain", "0")
            End If
            'BumpMapPlanetModel = CInt(Val(oINI.GetString("LIGHTING", "BumpMapPlanetModel", "0"))) <> 0
            If BumpMapPlanetModel = True Then
                oINI.WriteString("LIGHTING", "BumpMapPlanetModel", "1")
            Else : oINI.WriteString("LIGHTING", "BumpMapPlanetModel", "0")
            End If
            If IlluminationMapTerrain = True Then
                oINI.WriteString("LIGHTING", "IllumMapTerrain", "1")
            Else : oINI.WriteString("LIGHTING", "IllumMapTerrain", "0")
            End If

            'Dim lR As Int32
            'Dim lG As Int32
            'Dim lB As Int32
            'lR = Math.Max(Math.Min(255, CInt(NeutralAssetColor.X * 255)), 0)
            'lG = Math.Max(Math.Min(255, CInt(NeutralAssetColor.Y * 255)), 0)
            'lB = Math.Max(Math.Min(255, CInt(NeutralAssetColor.Z * 255)), 0)
            'oINI.WriteString("IDENTIFICATION", "NeutralR", lR.ToString)
            'oINI.WriteString("IDENTIFICATION", "NeutralG", lG.ToString)
            'oINI.WriteString("IDENTIFICATION", "NeutralB", lB.ToString)

            'lR = Math.Max(Math.Min(255, CInt(MyAssetColor.X * 255)), 0)
            'lG = Math.Max(Math.Min(255, CInt(MyAssetColor.Y * 255)), 0)
            'lB = Math.Max(Math.Min(255, CInt(MyAssetColor.Z * 255)), 0)
            'oINI.WriteString("IDENTIFICATION", "PlayerR", lR.ToString)
            'oINI.WriteString("IDENTIFICATION", "PlayerG", lG.ToString)
            'oINI.WriteString("IDENTIFICATION", "PlayerB", lB.ToString)

            'lR = Math.Max(Math.Min(255, CInt(EnemyAssetColor.X * 255)), 0)
            'lG = Math.Max(Math.Min(255, CInt(EnemyAssetColor.Y * 255)), 0)
            'lB = Math.Max(Math.Min(255, CInt(EnemyAssetColor.Z * 255)), 0)
            'oINI.WriteString("IDENTIFICATION", "EnemyR", lR.ToString)
            'oINI.WriteString("IDENTIFICATION", "EnemyG", lG.ToString)
            'oINI.WriteString("IDENTIFICATION", "EnemyB", lB.ToString)

            'lR = Math.Max(Math.Min(255, CInt(AllyAssetColor.X * 255)), 0)
            'lG = Math.Max(Math.Min(255, CInt(AllyAssetColor.Y * 255)), 0)
            'lB = Math.Max(Math.Min(255, CInt(AllyAssetColor.Z * 255)), 0)
            'oINI.WriteString("IDENTIFICATION", "AllyR", lR.ToString)
            'oINI.WriteString("IDENTIFICATION", "AllyG", lG.ToString)
            'oINI.WriteString("IDENTIFICATION", "AllyB", lB.ToString)

            'lR = Math.Max(Math.Min(255, CInt(GuildAssetColor.X * 255)), 0)
            'lG = Math.Max(Math.Min(255, CInt(GuildAssetColor.Y * 255)), 0)
            'lB = Math.Max(Math.Min(255, CInt(GuildAssetColor.Z * 255)), 0)
            'oINI.WriteString("IDENTIFICATION", "GuildR", lR.ToString)
            'oINI.WriteString("IDENTIFICATION", "GuildG", lG.ToString)
            'oINI.WriteString("IDENTIFICATION", "GuildB", lB.ToString)

            'With DefaultChatColor
            '    oINI.WriteString("CHATCOLOR", "DefaultR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "DefaultG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "DefaultB", .B.ToString)
            'End With
            'With AlertChatColor
            '    oINI.WriteString("CHATCOLOR", "AlertR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "AlertG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "AlertB", .B.ToString)
            'End With
            'With StatusChatColor
            '    oINI.WriteString("CHATCOLOR", "StatusR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "StatusG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "StatusB", .B.ToString)
            'End With
            'With LocalChatColor
            '    oINI.WriteString("CHATCOLOR", "LocalR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "LocalG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "LocalB", .B.ToString)
            'End With
            'With GuildChatColor
            '    oINI.WriteString("CHATCOLOR", "GuildR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "GuildG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "GuildB", .B.ToString)
            'End With
            'With PMChatColor
            '    oINI.WriteString("CHATCOLOR", "PMR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "PMG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "PMB", .B.ToString)
            'End With
            'With SenateChatColor
            '    oINI.WriteString("CHATCOLOR", "SenateR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "SenateG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "SenateB", .B.ToString)
            'End With
            'With ChannelChatColor
            '    oINI.WriteString("CHATCOLOR", "ChannelR", .R.ToString)
            '    oINI.WriteString("CHATCOLOR", "ChannelG", .G.ToString)
            '    oINI.WriteString("CHATCOLOR", "ChannelB", .B.ToString)
            'End With

            'With AcidBuildGhost
            '    oINI.WriteString("BUILDGHOST", "AcidR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "AcidG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "AcidB", .B.ToString)
            'End With
            'With AcidMineralCache
            '    oINI.WriteString("MINERALCACHE", "AcidR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "AcidG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "AcidB", .B.ToString)
            'End With
            'With AcidMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "AcidR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "AcidG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "AcidB", .B.ToString)
            'End With


            'With AdaptableBuildGhost
            '    oINI.WriteString("BUILDGHOST", "AdaptableR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "AdaptableG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "AdaptableB", .B.ToString)
            'End With
            'With AdaptableMineralCache
            '    oINI.WriteString("MINERALCACHE", "AdaptableR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "AdaptableG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "AdaptableB", .B.ToString)
            'End With
            'With AdaptableMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "AdaptableR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "AdaptableG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "AdaptableB", .B.ToString)
            'End With

            'With BarrenBuildGhost
            '    oINI.WriteString("BUILDGHOST", "BarrenR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "BarrenG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "BarrenB", .B.ToString)
            'End With
            'With BarrenMineralCache
            '    oINI.WriteString("MINERALCACHE", "BarrenR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "BarrenG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "BarrenB", .B.ToString)
            'End With
            'With BarrenMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "BarrenR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "BarrenG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "BarrenB", .B.ToString)
            'End With

            'With DesertBuildGhost
            '    oINI.WriteString("BUILDGHOST", "DesertR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "DesertG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "DesertB", .B.ToString)
            'End With
            'With DesertMineralCache
            '    oINI.WriteString("MINERALCACHE", "DesertR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "DesertG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "DesertB", .B.ToString)
            'End With
            'With DesertMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "DesertR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "DesertG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "DesertB", .B.ToString)
            'End With

            'With IceBuildGhost
            '    oINI.WriteString("BUILDGHOST", "IceR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "IceG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "IceB", .B.ToString)
            'End With
            'With IceMineralCache
            '    oINI.WriteString("MINERALCACHE", "IceR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "IceG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "IceB", .B.ToString)
            'End With
            'With IceMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "IceR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "IceG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "IceB", .B.ToString)
            'End With

            'With LavaBuildGhost
            '    oINI.WriteString("BUILDGHOST", "LavaR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "LavaG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "LavaB", .B.ToString)
            'End With
            'With LavaMineralCache
            '    oINI.WriteString("MINERALCACHE", "LavaR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "LavaG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "LavaB", .B.ToString)
            'End With
            'With LavaMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "LavaR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "LavaG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "LavaB", .B.ToString)
            'End With

            'With TerranBuildGhost
            '    oINI.WriteString("BUILDGHOST", "TerranR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "TerranG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "TerranB", .B.ToString)
            'End With
            'With TerranMineralCache
            '    oINI.WriteString("MINERALCACHE", "TerranR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "TerranG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "TerranB", .B.ToString)
            'End With
            'With TerranMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "TerranR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "TerranG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "TerranB", .B.ToString)
            'End With

            'With WaterworldBuildGhost
            '    oINI.WriteString("BUILDGHOST", "WaterworldR", .R.ToString)
            '    oINI.WriteString("BUILDGHOST", "WaterworldG", .G.ToString)
            '    oINI.WriteString("BUILDGHOST", "WaterworldB", .B.ToString)
            'End With
            'With WaterworldMineralCache
            '    oINI.WriteString("MINERALCACHE", "WaterworldR", .R.ToString)
            '    oINI.WriteString("MINERALCACHE", "WaterworldG", .G.ToString)
            '    oINI.WriteString("MINERALCACHE", "WaterworldB", .B.ToString)
            'End With
            'With WaterworldMinimapAngle
            '    oINI.WriteString("MINIMAPANGLE", "WaterworldR", .R.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "WaterworldG", .G.ToString)
            '    oINI.WriteString("MINIMAPANGLE", "WaterworldB", .B.ToString)
            'End With

            If ShowIntro = True Then
                oINI.WriteString("SETTINGS", "ShowIntroScreen", "1")
            Else : oINI.WriteString("SETTINGS", "ShowIntroScreen", "0")
            End If
            If ShowTargetBoxes = True Then
                oINI.WriteString("SETTINGS", "ShowTargetBoxes", "1")
            Else : oINI.WriteString("SETTINGS", "ShowTargetBoxes", "0")
            End If

            oINI.WriteString("GRAPHICS", "WaterRenderMethod", WaterRenderMethod.ToString)
            oINI.WriteString("GRAPHICS", "WaterTextureRes", WaterTextureRes.ToString)
            oINI.WriteString("GRAPHICS", "SmoothFOW", CInt(SmoothFOW).ToString)
            'oINI.WriteString("GRAPHICS", "MaxLights", MaxLights.ToString)
            'oINI.WriteString("GRAPHICS", "FarClippingPlane", FarClippingPlane.ToString)
            'oINI.WriteString("GRAPHICS", "NearClippingPlane", NearClippingPlane.ToString)
            'oINI.WriteString("GRAPHICS", "TerrainNearClippingPlane", TerrainNearClippingPlane.ToString)
            'oINI.WriteString("GRAPHICS", "TerrainFarClippingPlane", TerrainFarClippingPlane.ToString)
            'oINI.WriteString("PLANET_VIEW", "ShowMiniMap", CInt(ShowMiniMap).ToString)
            oINI.WriteString("GRAPHICS", "EntityClipPlane", EntityClipPlane.ToString)
            oINI.WriteString("GRAPHICS", "ModelTextureResolution", CInt(ModelTextureResolution).ToString)
            oINI.WriteString("GRAPHICS", "DrawGrid", CInt(DrawGrid).ToString)
            oINI.WriteString("GRAPHICS", "ShowFOWTerrainShading", CInt(ShowFOWTerrainShading).ToString)
            oINI.WriteString("GRAPHICS", "FOWTextureResolution", FOWTextureResolution.ToString)
            'oINI.WriteString("HUD", "MiniMapLocX", MiniMapLocX.ToString)
            'oINI.WriteString("HUD", "MiniMapLocY", MiniMapLocY.ToString)
            'oINI.WriteString("HUD", "MiniMapWidthHeight", MiniMapWidthHeight.ToString)
            oINI.WriteString("GRAPHICS", "RenderMineralCaches", CInt(RenderMineralCaches).ToString)
            'oINI.WriteString("GRAPHICS", "DITHER", CInt(Dither).ToString)
            'oINI.WriteString("GRAPHICS", "SpecularEnabled", CInt(SpecularEnabled).ToString)
            'oINI.WriteString("SETTINGS", "NotificationDisplayTime", NotificationDisplayTime.ToString)
            oINI.WriteString("AUDIO", "AudioEnabled", CInt(AudioOn).ToString)
            'oINI.WriteString("SETTINGS", "ScreenShakeEnable", CInt(ScreenShakeEnabled).ToString)
            oINI.WriteString("SETTINGS", "PlanetMinimapZoomLevel", CInt(PlanetMinimapZoomLevel).ToString)
            oINI.WriteString("GRAPHICS", "TerrainTextureResolution", TerrainTextureResolution.ToString)

            oINI.WriteString("GRAPHICS", "EngineFXParticles", EngineFXParticles.ToString)
            oINI.WriteString("GRAPHICS", "BurnFXParticles", BurnFXParticles.ToString)
            oINI.WriteString("GRAPHICS", "RenderShieldFX", CInt(RenderShieldFX).ToString)
            oINI.WriteString("GRAPHICS", "StarfieldParticles", StarfieldParticlesSpace.ToString)
            oINI.WriteString("GRAPHICS", "PlanetFXParticles", PlanetFXParticles.ToString)
            oINI.WriteString("GRAPHICS", "PostGlowAmount", PostGlowAmt.ToString("0.#"))
            oINI.WriteString("GRAPHICS", "HiResPlanetTexture", CInt(HiResPlanetTexture).ToString)
            oINI.WriteString("GRAPHICS", "PlanetModelTextureResolution", PlanetModelTextureWH.ToString)

            'oINI.WriteString("HUD", "BuildWindowLeft", BuildWindowLeft.ToString)
            'oINI.WriteString("HUD", "BuildWindowTop", BuildWindowTop.ToString)
            'oINI.WriteString("HUD", "ResearchWindowLeft", ResearchWindowLeft.ToString)
            'oINI.WriteString("HUD", "ResearchWindowTop", ResearchWindowTop.ToString)

            'oINI.WriteString("SETTINGS", "CtrlQExits", CInt(CtrlQExits).ToString)
            'oINI.WriteString("SETTINGS", "RenderHPBars", CInt(RenderHPBars).ToString)
            'oINI.WriteString("INTERFACE", "DoNotShowEngineerCancelAlert", CInt(bDoNotShowEngineerCancelAlert).ToString)

            'If bRanBefore = True Then oINI.WriteString("SETTINGS", "RanBefore", "1")

            oINI.WriteString("INTERFACE", "BorderColor_A", InterfaceBorderColor.A.ToString)
            oINI.WriteString("INTERFACE", "BorderColor_R", InterfaceBorderColor.R.ToString)
            oINI.WriteString("INTERFACE", "BorderColor_G", InterfaceBorderColor.G.ToString)
            oINI.WriteString("INTERFACE", "BorderColor_B", InterfaceBorderColor.B.ToString)

            oINI.WriteString("INTERFACE", "FillColor_A", InterfaceFillColor.A.ToString)
            oINI.WriteString("INTERFACE", "FillColor_R", InterfaceFillColor.R.ToString)
            oINI.WriteString("INTERFACE", "FillColor_G", InterfaceFillColor.G.ToString)
            oINI.WriteString("INTERFACE", "FillColor_B", InterfaceFillColor.B.ToString)

            oINI.WriteString("INTERFACE", "TextBoxFillColor_A", InterfaceTextBoxFillColor.A.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_R", InterfaceTextBoxFillColor.R.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_G", InterfaceTextBoxFillColor.G.ToString)
            oINI.WriteString("INTERFACE", "TextBoxFillColor_B", InterfaceTextBoxFillColor.B.ToString)

            oINI.WriteString("INTERFACE", "TextBoxForeColor_A", InterfaceTextBoxForeColor.A.ToString)
            oINI.WriteString("INTERFACE", "TextBoxForeColor_R", InterfaceTextBoxForeColor.R.ToString)
            oINI.WriteString("INTERFACE", "TextBoxForeColor_G", InterfaceTextBoxForeColor.G.ToString)
            oINI.WriteString("INTERFACE", "TextBoxForeColor_B", InterfaceTextBoxForeColor.B.ToString)

            oINI.WriteString("INTERFACE", "ButtonColor_A", InterfaceButtonColor.A.ToString)
            oINI.WriteString("INTERFACE", "ButtonColor_R", InterfaceButtonColor.R.ToString)
            oINI.WriteString("INTERFACE", "ButtonColor_G", InterfaceButtonColor.G.ToString)
            oINI.WriteString("INTERFACE", "ButtonColor_B", InterfaceButtonColor.B.ToString)

            'Hud Position Settings
            'If AdvanceDisplayLocX <> -1 Then oINI.WriteString("INTERFACE", "AdvanceDisplayLocX", AdvanceDisplayLocX.ToString)
            'If AdvanceDisplayLocY <> -1 Then oINI.WriteString("INTERFACE", "AdvanceDisplayLocY", AdvanceDisplayLocY.ToString)
            'If BehaviorLocX <> -1 Then oINI.WriteString("INTERFACE", "BehaviorLocX", BehaviorLocX.ToString)
            'If BehaviorLocY <> -1 Then oINI.WriteString("INTERFACE", "BehaviorLocY", BehaviorLocY.ToString)
            'If ContentsLocX <> -1 Then oINI.WriteString("INTERFACE", "ContentsLocX", ContentsLocX.ToString)
            'If ContentsLocY <> -1 Then oINI.WriteString("INTERFACE", "ContentsLocY", ContentsLocY.ToString)
            'If EnvirDisplayLocX <> -1 Then oINI.WriteString("INTERFACE", "EnvirDisplayLocX", EnvirDisplayLocX.ToString)
            'If EnvirDisplayLocY <> -1 Then oINI.WriteString("INTERFACE", "EnvirDisplayLocY", EnvirDisplayLocY.ToString)
            'If SelectLocX <> -1 Then oINI.WriteString("INTERFACE", "SelectLocX", SelectLocX.ToString)
            'If SelectLocY <> -1 Then oINI.WriteString("INTERFACE", "SelectLocY", SelectLocY.ToString)
            'If ProdStatusLocX <> -1 Then oINI.WriteString("INTERFACE", "ProdStatusLocX", ProdStatusLocX.ToString)
            'If ProdStatusLocY <> -1 Then oINI.WriteString("INTERFACE", "ProdStatusLocY", ProdStatusLocY.ToString)
            'If ChatWindowLocX <> -1 Then oINI.WriteString("INTERFACE", "ChatWindowLocX", ChatWindowLocX.ToString)
            'If ChatWindowLocY <> -1 Then oINI.WriteString("INTERFACE", "ChatWindowLocY", ChatWindowLocY.ToString)
            'End of HUD Position settings
            'If FilterBadWords = True Then oINI.WriteString("INTERFACE", "FilterBadWords", "0") Else oINI.WriteString("INTERFACE", "FilterBadWords", "1")

            oINI = Nothing
        End Sub
    End Structure
    Public muSettings As EngineSettings
#End Region

End Module

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
'            Dim fZOffset As Single = -9000

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

'                fXS = (Rnd() * 6) - 3
'                fYS = 0

'                'Ok, set up its extent...
'                mfMinExt(lIndex) = vecEmitter.X - 480
'                mfMaxExt(lIndex) = vecEmitter.X - 352
'                mbUseHorizontal(lIndex) = True

'                moParticles(lIndex).Reset(vecEmitter.X - 480 + (3.2F * lTemp), lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 100 Then
'                'C Group -15
'                lTemp = lIndex - 80
'                fXS = (Rnd() * 6) - 3
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
'                fYS = (Rnd() * 6) - 3

'                'ok, set up its extent
'                mfMinExt(lIndex) = vecEmitter.Y - 64
'                mfMaxExt(lIndex) = vecEmitter.Y + 64
'                mbUseHorizontal(lIndex) = False

'                moParticles(lIndex).Reset(vecEmitter.X - 480, vecEmitter.Y - 64 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 210 Then
'                'E and F Group
'                lTemp = lIndex - 150
'                fXS = Rnd() * 6 - 3
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
'                fXS = 0 : fYS = Rnd() * 6 - 3

'                mfMinExt(lIndex) = vecEmitter.Y : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False
'                moParticles(lIndex).Reset(vecEmitter.X - 160, vecEmitter.Y + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 270 Then
'                'H Group
'                lTemp = lIndex - 220
'                fXS = 0 : fYS = Rnd() * 6 - 3
'                mfMinExt(lIndex) = vecEmitter.Y - 64 : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False
'                moParticles(lIndex).Reset(vecEmitter.X - 256, vecEmitter.Y - 64 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 370 Then
'                'I or J group
'                lTemp = lIndex - 270
'                fXS = Rnd() * 6 - 3 : fYS = 0
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
'                fXS = 0 : fYS = Rnd() * 6 - 3
'                mfMinExt(lIndex) = vecEmitter.Y - 64 : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False '64
'                moParticles(lIndex).Reset(vecEmitter.X + 16, vecEmitter.Y - 32 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 490 Then
'                'L and M group
'                lTemp = lIndex - 410
'                fXS = Rnd() * 6 - 3 : fYS = 0
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
'                fXS = 0 : fYS = Rnd() * 6 - 3
'                mfMinExt(lIndex) = vecEmitter.Y - 64 : mfMaxExt(lIndex) = vecEmitter.Y + 64 : mbUseHorizontal(lIndex) = False
'                moParticles(lIndex).Reset(vecEmitter.X + 192, vecEmitter.Y - 64 + (3.2F * lTemp), 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            ElseIf lIndex < 600 Then
'                'O and P Group
'                lTemp = lIndex - 540
'                fXS = Rnd() * 6 - 3 : fYS = 0
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
'                fXS = 0 : fYS = Rnd() * 6 - 3
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
'                fXS = Rnd() * 6 - 3 : fYS = 0
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
'                fXS = Rnd() * 6 - 3 : fYS = 0
'                If lTemp > 220 Then
'                    lTemp -= 110
'                    lYLoc = CInt(vecEmitter.Y + 100)
'                Else : lYLoc = CInt(vecEmitter.Y - 100)
'                End If
'                lXLoc = CInt((Rnd() * 1024) - 512) '+ vecEmitter.X
'                If lXLoc < 0 Then fXS = (Rnd() * 5) + 1 Else fXS = -((Rnd() * 5) + 1)
'                mfMinExt(lIndex) = vecEmitter.X - 512 : mfMaxExt(lIndex) = vecEmitter.X + 512 : mbUseHorizontal(lIndex) = True
'                moParticles(lIndex).Reset(lXLoc, lYLoc, 0, fXS, fYS, 0, 0, 0, 0, 255, 255, 255, 255)
'            Else
'                lTemp = CInt(((lIndex - 1200) / (mlParticleUB - 1200)) * 1200)
'                lXLoc = CInt(moParticles(lTemp).vecLoc.X)
'                lYLoc = CInt(moParticles(lTemp).vecLoc.Y)

'                Dim fXA As Single
'                Dim fYA As Single

'                If Rnd() * 100 < 50 Then
'                    fXS = -((Rnd() * 0.25F) + 0.3F) : fXA = Rnd() * 0.01F
'                Else : fXS = Rnd() * 0.25F + 0.3F : fXA = Rnd() * -0.01F
'                End If
'                If Rnd() * 100 < 50 Then
'                    fYS = -((Rnd() * 0.25F) + 0.3F) : fYA = Rnd() * 0.001F
'                Else : fYS = Rnd() * 0.25F + 0.3F : fYA = Rnd() * -0.001F
'                End If

'                moParticles(lIndex).Reset(lXLoc, lYLoc, 0, fXS, fYS, 0, fXA, fYA, 0, 255, 255, 255, 255)
'                moParticles(lIndex).fAChg = -((Rnd() * 3) + 2)

'                If mbFadedIn = False Then
'                    moParticles(lIndex).mfA = 5
'                End If

'                'moParticles(lIndex).fRChg = -((Rnd() * 5) + 25)
'                moParticles(lIndex).fGChg = -((Rnd() * 5) + 1)
'                moParticles(lIndex).fBChg = -((Rnd() * 1) + 1)
'            End If

'            If lIndex < 1200 Then       '680
'                moParticles(lIndex).mfA = 0
'                moParticles(lIndex).fAChg = 1
'                moParticles(lIndex).fRChg = 0
'                moParticles(lIndex).fGChg = Rnd() * -5
'                moParticles(lIndex).fBChg = Rnd() * -5
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
'                    'If .mfR <= 64 Then .fRChg = Rnd() * 5 : .mfR = 64
'                    If .mfG <= 64 Then .fGChg = Rnd() * 5 : .mfG = 64
'                    If .mfB <= 64 Then .fBChg = Rnd() * 5 : .mfB = 64

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

'        moDevice.Transform.World = Matrix.Identity
'        moDevice.Transform.View = Matrix.LookAtLH(New Vector3(0, 0, -10000), New Vector3(0, 0, 0), New Vector3(0.0F, 1.0F, 0.0F))    'up is always this
'        moDevice.Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, CSng(moDevice.PresentationParameters.BackBufferWidth / moDevice.PresentationParameters.BackBufferHeight), muSettings.TerrainNearClippingPlane, muSettings.FarClippingPlane)

'        'now, render
'        Device.IsUsingEventHandlers = False
'        If moParticle Is Nothing OrElse moParticle.Disposed = True Then
'            'moParticle = goResMgr.GetTexture("Particle.dds", EpicaResourceManager.eGetTextureType.NoSpecifics)
'            moParticle = moResMgr.GetTexture("Particle.dds", EpicaResourceManager.eGetTextureType.NoSpecifics)
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

Public Class PostShader
    Private moEffect As Effect = Nothing
    Private moDevice As Device

    Private moSceneMap As Texture
    Private moDownsampleMap As Texture
    Private moBlurMap1 As Texture
    Private moBlurMap2 As Texture

    Private mhdlWindowSize As EffectHandle
    Private mhdlSceneMap As EffectHandle
    Private mhdlDownsample As EffectHandle
    Private mhdlBlurMap1 As EffectHandle
    Private mhdlBlurMap2 As EffectHandle
    Private mhdlGlowIntensity As EffectHandle

    Private mfPreviousGlow As Single = -1.0F
    Private mfPreviousfWS0 As Single = 0
    Private mfPreviousfWS1 As Single = 0

    Private mfWS_Values() As Single

    Public Sub ExecutePostProcess()

        If moEffect Is Nothing OrElse moEffect.Disposed = True Then
            ReinitializeEffect()
        End If
        'Validate our parameters
        If mhdlSceneMap Is Nothing = False Then moEffect.SetValue(mhdlSceneMap, moSceneMap)
        If mfPreviousGlow <> muSettings.PostGlowAmt Then
            mfPreviousGlow = muSettings.PostGlowAmt
            If mhdlGlowIntensity Is Nothing = False Then moEffect.SetValue(mhdlGlowIntensity, muSettings.PostGlowAmt)
        End If
        ValidateTextures()

        Dim oOriginal As Surface = moDevice.GetBackBuffer(0, 0, BackBufferType.Mono)
        Dim oScene As Surface = moSceneMap.GetSurfaceLevel(0)

        moDevice.StretchRectangle(oOriginal, New Rectangle(0, 0, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight), _
          oScene, New System.Drawing.Rectangle(0, 0, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight), TextureFilter.None)

        If oScene Is Nothing = False Then oScene.Dispose()
        oScene = Nothing

        'Turn off the Z buffer
        moDevice.RenderState.ZBufferEnable = False
        moDevice.RenderState.ZBufferWriteEnable = False
        'And alpha blending/light
        moDevice.RenderState.AlphaBlendEnable = False
        moDevice.RenderState.Lighting = False

        moDevice.Indices = Nothing
        moDevice.SetTexture(0, Nothing)
        moDevice.SetTexture(1, Nothing)
        moDevice.RenderState.CullMode = Cull.None
        moDevice.RenderState.FogEnable = False

        Dim lPasses As Int32 = moEffect.Begin(FX.None)
        For X As Int32 = 0 To lPasses - 1
            If oScene Is Nothing = False Then oScene.Dispose()
            If X = 0 Then
                'downsample
                oScene = moDownsampleMap.GetSurfaceLevel(0)
                moDevice.SetRenderTarget(0, oScene)
            ElseIf X = 1 Then
                'blurmap1
                oScene = moBlurMap1.GetSurfaceLevel(0)
                moDevice.SetRenderTarget(0, oScene)
            ElseIf X = 2 Then
                'blurmap2
                oScene = moBlurMap2.GetSurfaceLevel(0)
                moDevice.SetRenderTarget(0, oScene)
            Else
                'reset render target
                moDevice.SetRenderTarget(0, oOriginal)
            End If

            moEffect.BeginPass(X)
            RenderQuad(moDevice)
            moEffect.EndPass()

        Next X
        moEffect.End()
        If oScene Is Nothing = False Then oScene.Dispose()
        oScene = Nothing

        'Turn on the Z buffer
        moDevice.RenderState.ZBufferEnable = True
        moDevice.RenderState.ZBufferWriteEnable = True
        'And alpha blending/light
        moDevice.RenderState.AlphaBlendEnable = True
        moDevice.RenderState.Lighting = True
        oOriginal.Dispose()
        oOriginal = Nothing
    End Sub

    Private Sub ValidateTextures()

        If moSceneMap Is Nothing OrElse moSceneMap.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moSceneMap = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            Device.IsUsingEventHandlers = True
        End If
        If moDownsampleMap Is Nothing OrElse moDownsampleMap.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moDownsampleMap = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth \ 4, moDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            If mhdlDownsample Is Nothing = False Then moEffect.SetValue(mhdlDownsample, moDownsampleMap)
            Device.IsUsingEventHandlers = True
        End If
        If moBlurMap1 Is Nothing OrElse moBlurMap1.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moBlurMap1 = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth \ 4, moDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            If mhdlBlurMap1 Is Nothing = False Then moEffect.SetValue(mhdlBlurMap1, moBlurMap1)
            Device.IsUsingEventHandlers = True
        End If
        If moBlurMap2 Is Nothing OrElse moBlurMap2.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moBlurMap2 = New Texture(moDevice, moDevice.PresentationParameters.BackBufferWidth \ 4, moDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, moDevice.PresentationParameters.BackBufferFormat, Pool.Default)
            If mhdlBlurMap2 Is Nothing = False Then moEffect.SetValue(mhdlBlurMap2, moBlurMap2)
            Device.IsUsingEventHandlers = True
        End If


        mfWS_Values(0) = moDevice.PresentationParameters.BackBufferWidth
        mfWS_Values(1) = moDevice.PresentationParameters.BackBufferHeight
        If mfPreviousfWS0 <> mfWS_Values(0) OrElse mfPreviousfWS1 <> mfWS_Values(1) Then
            mfPreviousfWS0 = mfWS_Values(0) : mfPreviousfWS1 = mfWS_Values(1)
            If mhdlWindowSize Is Nothing = False Then moEffect.SetValue(mhdlWindowSize, mfWS_Values)
        End If

    End Sub

    Public Sub New(ByRef oDevice As Device)
        moDevice = oDevice

        ReDim mfWS_Values(1)

        ReinitializeEffect()
    End Sub

    Private Sub ReinitializeEffect()
        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("ClientConfig.PostScreenGlow.fx")
        Device.IsUsingEventHandlers = False
		moEffect = Effect.FromStream(moDevice, oStream, Nothing, ShaderFlags.None, Nothing)
		'moEffect = Effect.FromStream(moDevice, oStream, Nothing, Nothing, ShaderFlags.None, Nothing)
        Device.IsUsingEventHandlers = True

        mhdlWindowSize = moEffect.GetParameter(Nothing, "windowSize")
        mhdlSceneMap = moEffect.GetParameter(Nothing, "sceneMap")
        mhdlDownsample = moEffect.GetParameter(Nothing, "downsampleMap")
        mhdlBlurMap1 = moEffect.GetParameter(Nothing, "blurMap1")
        mhdlBlurMap2 = moEffect.GetParameter(Nothing, "blurMap2")
        mhdlGlowIntensity = moEffect.GetParameter(Nothing, "GlowIntensity")

        moEffect.Technique = "ScreenGlow"
    End Sub

    Private Shared mvbQuad As VertexBuffer = Nothing
    Private Shared Sub RenderQuad(ByRef oDevice As Device)
        If mvbQuad Is Nothing OrElse mvbQuad.Disposed = True Then
            If InitVBQuad(oDevice) = False Then Return
        End If
        oDevice.VertexFormat = CustomVertex.PositionTextured.Format
        oDevice.SetStreamSource(0, mvbQuad, 0, CustomVertex.PositionTextured.StrideSize)
        oDevice.DrawPrimitives(PrimitiveType.TriangleStrip, 0, 2)
    End Sub
    Private Shared Function InitVBQuad(ByRef oDevice As Device) As Boolean
        Dim uVert As CustomVertex.PositionTextured
        If mvbQuad Is Nothing OrElse mvbQuad.Disposed = True Then mvbQuad = New VertexBuffer(uVert.GetType, 4, oDevice, Usage.WriteOnly, CustomVertex.PositionTextured.Format, Pool.Managed)
        Dim bResult As Boolean = False

        Try
            Dim gfxStream As GraphicsStream = mvbQuad.Lock(0, 0, 0)
            Dim uVerts(3) As CustomVertex.PositionTextured
            uVerts(0) = New CustomVertex.PositionTextured(New Vector3(-1.0F, -1.0F, 0.5F), 0, 1)
            uVerts(1) = New CustomVertex.PositionTextured(New Vector3(-1.0F, 1.0F, 0.5F), 0, 0)
            uVerts(2) = New CustomVertex.PositionTextured(New Vector3(1.0F, -1.0F, 0.5F), 1, 1)
            uVerts(3) = New CustomVertex.PositionTextured(New Vector3(1.0F, 1.0F, 0.5F), 1, 0)
            gfxStream.Write(uVerts)
            mvbQuad.Unlock()
            bResult = True
        Catch ex As Exception
            bResult = False
        End Try
        Return bResult
    End Function

    Public Sub DisposeMe()
        mhdlWindowSize = Nothing
        mhdlSceneMap = Nothing
        mhdlDownsample = Nothing
        mhdlBlurMap1 = Nothing
        mhdlBlurMap2 = Nothing
        mhdlGlowIntensity = Nothing
        Try
            If moSceneMap Is Nothing = False Then moSceneMap.Dispose()
        Catch
        End Try
        moSceneMap = Nothing
        Try
            If moDownsampleMap Is Nothing = False Then moDownsampleMap.Dispose()
        Catch
        End Try
        moDownsampleMap = Nothing
        Try
            If moBlurMap1 Is Nothing = False Then moBlurMap1.Dispose()
        Catch
        End Try
        moBlurMap1 = Nothing
        Try
            If moBlurMap2 Is Nothing = False Then moBlurMap2.Dispose()
        Catch
        End Try
        moBlurMap2 = Nothing
        Try
            If moEffect Is Nothing = False Then moEffect.Dispose()
        Catch
        End Try
        moEffect = Nothing
        moDevice = Nothing
    End Sub
End Class

Public Class Camera
    Public mlCameraX As Int32 = 8300
    Public mlCameraY As Int32 = 875
    Public mlCameraZ As Int32 = 5500
    Public mlCameraAtX As Int32 = 0
    Public mlCameraAtY As Int32 = 0
    Public mlCameraAtZ As Int32 = 0

    Public bScrollLeft As Boolean = False
    Public bScrollRight As Boolean = False
    Public bScrollUp As Boolean = False
    Public bScrollDown As Boolean = False

    Public Function GetCameraAngleDegrees() As Single
        Return LineAngleDegrees(mlCameraX, mlCameraZ, mlCameraAtX, mlCameraAtZ)
    End Function

    Public Sub SetupMatrices(ByRef oDevice As Device)
        oDevice.Transform.View = Matrix.LookAtLH(New Vector3(mlCameraX, mlCameraY, mlCameraZ), New Vector3(mlCameraAtX, mlCameraAtY, mlCameraAtZ), New Vector3(0.0F, 1.0F, 0.0F))    'up is always this

        oDevice.Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, CSng(oDevice.PresentationParameters.BackBufferWidth / oDevice.PresentationParameters.BackBufferHeight), muSettings.TerrainNearClippingPlane, muSettings.FarClippingPlane)
        'oDevice.Transform.Projection = Matrix.PerspectiveFovLH(0.7853982F, CSng(oDevice.PresentationParameters.BackBufferWidth / oDevice.PresentationParameters.BackBufferHeight), muSettings.NearClippingPlane, muSettings.FarClippingPlane)

    End Sub


    'This sub is called every frame
    Public Sub ScrollCamera(ByVal CurrentEnvirView As Int32)

        'handle view scrolling... however, this comes with a twist... we need to determine our angle to the target
        ' deltaXx is the delta when scrolling camera's X along the X world axis
        ' deltaXz is the delta when scrolling camera's X along the Z world axis
        ' deltaZx is the delta when scrolling camera's Z along the X world axis
        ' deltaZz is the delta when scrolling camera's Z along the Z world axis
        ' which is to say, when changing X (scrolling horizontally), deltaXx is change to X and deltaXz is change to Z
        ' and when changing Z (scrolling vertically), deltaZx is change to X and deltaZz is change to Z
        Dim vecCameraLoc As Vector3 = New Vector3(mlCameraX, mlCameraY, mlCameraZ)
        Dim vecCameraAt As Vector3 = New Vector3(mlCameraAtX, mlCameraAtY, mlCameraAtZ)
        Dim vecTemp As Vector3 = Vector3.Subtract(vecCameraLoc, vecCameraAt)

        Dim deltaXx As Int32
        Dim deltaXz As Int32
        Dim deltaZx As Int32
        Dim deltaZz As Int32

        Dim lScrollRate As Int32

        vecTemp.Normalize()
        Dim vecDot As Vector3 = Vector3.Cross(vecTemp, New Vector3(0, 1, 0))


        lScrollRate = 500

        deltaXx = CInt(vecDot.X * lScrollRate)
        deltaXz = CInt(vecDot.Z * lScrollRate)
        deltaZx = CInt(vecTemp.X * lScrollRate)
        deltaZz = CInt(vecTemp.Z * lScrollRate)

        If bScrollLeft Then
            mlCameraX -= deltaXx
            mlCameraAtX -= deltaXx
            mlCameraZ -= deltaXz
            mlCameraAtZ -= deltaXz
            'Clear our tracking index 
        ElseIf bScrollRight Then
            mlCameraX += deltaXx
            mlCameraAtX += deltaXx
            mlCameraZ += deltaXz
            mlCameraAtZ += deltaXz
            'Clear our tracking index 
        End If
        If bScrollUp Then
            mlCameraZ -= deltaZz
            mlCameraAtZ -= deltaZz
            mlCameraX -= deltaZx
            mlCameraAtX -= deltaZx
            'Clear our tracking index 
        ElseIf bScrollDown Then
            mlCameraZ += deltaZz
            mlCameraAtZ += deltaZz
            mlCameraX += deltaZx
            mlCameraAtX += deltaZx 
        End If
  

        vecTemp = Nothing
        vecCameraAt = Nothing
        vecCameraLoc = Nothing
        vecDot = Nothing
    End Sub

    Public Sub ModifyZoom(ByVal lAmt As Int32) 
        Dim oVec3 As Vector3 = New Vector3(mlCameraX - mlCameraAtX, mlCameraY - mlCameraAtY, mlCameraZ - mlCameraAtZ)
        oVec3.Normalize()
        mlCameraX += CInt(lAmt * oVec3.X)
        mlCameraY += CInt(lAmt * oVec3.Y)
        mlCameraZ += CInt(lAmt * oVec3.Z)

        oVec3 = Nothing
    End Sub

    Public Sub RotateCamera(ByVal deltaX As Int32, ByVal deltaY As Int32)
        Dim fXYaw As Single
        Dim fZYaw As Single

        'X is easy
        RotatePoint(mlCameraAtX, mlCameraAtZ, mlCameraX, mlCameraZ, deltaX)

        'Ok, Y is the real pain...
        Dim vecAngle As Vector3 = New Vector3(mlCameraAtX - mlCameraX, mlCameraAtY - mlCameraY, mlCameraAtZ - mlCameraZ)
        vecAngle.Normalize()

        fXYaw = vecAngle.X * deltaY
        fZYaw = vecAngle.Z * deltaY

        RotatePoint(mlCameraAtX, mlCameraAtY, mlCameraX, mlCameraY, fXYaw)
        RotatePoint(mlCameraAtZ, mlCameraAtY, mlCameraZ, mlCameraY, fZYaw)
 
    End Sub


#Region " Frustrum Culling "
    Private mp_Frustrum() As Plane

    Public Sub CalculateFrustrum(ByRef oDevice As Device)

        Dim lClipPlane As Int32 = muSettings.EntityClipPlane
 

        Dim matProj As Matrix = Matrix.PerspectiveFovLH(0.7853982F, CSng(oDevice.PresentationParameters.BackBufferWidth / oDevice.PresentationParameters.BackBufferHeight), 1, lClipPlane)
        Dim matComb As Matrix = Matrix.Multiply(oDevice.Transform.View, matProj)

        If mp_Frustrum Is Nothing Then ReDim mp_Frustrum(5)

        'Left clipping plane
        With mp_Frustrum(0)
            .A = matComb.M14 + matComb.M11
            .B = matComb.M24 + matComb.M21
            .C = matComb.M34 + matComb.M31
            .D = matComb.M44 + matComb.M41
        End With

        'Right clipping plane
        With mp_Frustrum(1)
            .A = matComb.M14 - matComb.M11
            .B = matComb.M24 - matComb.M21
            .C = matComb.M34 - matComb.M31
            .D = matComb.M44 - matComb.M41
        End With

        'Top Clipping Plane
        With mp_Frustrum(2)
            .A = matComb.M14 - matComb.M12
            .B = matComb.M24 - matComb.M22
            .C = matComb.M34 - matComb.M32
            .D = matComb.M44 - matComb.M42
        End With

        'Bottom clipping plane
        With mp_Frustrum(3)
            .A = matComb.M14 + matComb.M12
            .B = matComb.M24 + matComb.M22
            .C = matComb.M34 + matComb.M32
            .D = matComb.M44 + matComb.M42
        End With

        'Near clipping plane
        With mp_Frustrum(4)
            .A = matComb.M13
            .B = matComb.M23
            .C = matComb.M33
            .D = matComb.M43
        End With

        'Far clipping Plane
        With mp_Frustrum(5)
            .A = matComb.M14 - matComb.M13
            .B = matComb.M24 - matComb.M23
            .C = matComb.M34 - matComb.M33
            .D = matComb.M44 - matComb.M43
        End With

        For X As Int32 = 0 To 5
            mp_Frustrum(X).Normalize()
        Next X
    End Sub

    Public Function CullObject(ByRef bBox As CullBox) As Boolean
        'for each plane in the frustrum
        Dim c1 As Vector3
        Dim c2 As Vector3

        For X As Int32 = 0 To 5
            With mp_Frustrum(X)
                'find furthest and nears point to the plane
                If .A > 0.0F Then
                    c1.X = bBox.vecMax.X : c2.X = bBox.vecMin.X
                Else
                    c1.X = bBox.vecMin.X : c2.X = bBox.vecMax.X
                End If
                If .B > 0.0F Then
                    c1.Y = bBox.vecMax.Y : c2.Y = bBox.vecMin.Y
                Else
                    c1.Y = bBox.vecMin.Y : c2.Y = bBox.vecMax.Y
                End If
                If .C > 0.0F Then
                    c1.Z = bBox.vecMax.Z : c2.Z = bBox.vecMin.Z
                Else
                    c1.Z = bBox.vecMin.Z : c2.Z = bBox.vecMax.Z
                End If

                Dim fD1 As Single = .A * c1.X + .B * c1.Y + .C * c1.Z + .D
                If fD1 < 0.0F Then
                    Dim fD2 As Single = .A * c2.X + .B * c2.Y + .C * c2.Z + .D
                    If fD2 < 0.0F Then Return True
                End If
            End With
        Next X

        Return False
    End Function

    Public Function CullObject_NoMaxDistance(ByRef bBox As CullBox) As Boolean
        'for each plane in the frustrum
        Dim c1 As Vector3
        Dim c2 As Vector3

        For X As Int32 = 0 To 4
            With mp_Frustrum(X)
                'find furthest and nears point to the plane
                If .A > 0.0F Then
                    c1.X = bBox.vecMax.X : c2.X = bBox.vecMin.X
                Else
                    c1.X = bBox.vecMin.X : c2.X = bBox.vecMax.X
                End If
                If .B > 0.0F Then
                    c1.Y = bBox.vecMax.Y : c2.Y = bBox.vecMin.Y
                Else
                    c1.Y = bBox.vecMin.Y : c2.Y = bBox.vecMax.Y
                End If
                If .C > 0.0F Then
                    c1.Z = bBox.vecMax.Z : c2.Z = bBox.vecMin.Z
                Else
                    c1.Z = bBox.vecMin.Z : c2.Z = bBox.vecMax.Z
                End If

                Dim fD1 As Single = .A * c1.X + .B * c1.Y + .C * c1.Z + .D
                If fD1 < 0.0F Then
                    Dim fD2 As Single = .A * c2.X + .B * c2.Y + .C * c2.Z + .D
                    If fD2 < 0.0F Then Return True
                End If
            End With
        Next X

        Return False
    End Function
#End Region


End Class


Public Class BPLogo

	Private muBeyondVerts() As CustomVertex.TransformedColoredTextured
	Private muProtocolVerts() As CustomVertex.TransformedColoredTextured

	Private mlBeyondAlpha() As Int32
	Private mlProtocolAlpha() As Int32
	Private mlSpriteAlpha As Int32 = 0
	Private mbFadedIn As Boolean = False

	Private moBeyondTex As Texture = Nothing
	Private moProtocolTex As Texture = Nothing
	Private moBPTex As Texture = Nothing
	Private moSprite As Sprite = Nothing

	Public Sub New(ByRef oDevice As Device, ByVal lVertUB As Int32)		'vertub should be 19
		ReDim muBeyondVerts(lVertUB)
		ReDim muProtocolVerts(lVertUB)
		ReDim mlBeyondAlpha(lVertUB)
		ReDim mlProtocolAlpha(lVertUB)

		Dim lTotalWidth As Int32 = CInt(oDevice.PresentationParameters.BackBufferWidth * 0.9F)

		'both images (beyond and protocol) are 512x64

		'we'll do the Beyond Verts first...
		Dim lOffset As Int32 = (oDevice.PresentationParameters.BackBufferWidth - lTotalWidth) \ 2
		'now, get our width of beyond...... BP icon is 164 x 256
		Dim lBeyondWidth As Int32 = oDevice.PresentationParameters.BackBufferWidth \ 2 - lOffset - 42		'42 for the BP Icon
		Dim lVertWidth As Int32 = lBeyondWidth \ (muBeyondVerts.GetUpperBound(0) \ 2)
		Dim lVertHeight As Int32 = CInt((lBeyondWidth / 512.0F) * 64)

		Dim lTop As Int32 = 80 - (lVertHeight \ 2)
		Dim lBottom As Int32 = lTop + lVertHeight

		For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
			With muBeyondVerts(X)
				.X = (X \ 2) * lVertWidth
				If X Mod 2 = 0 Then
					.Y = lTop
					.Tv = 0
				Else
					.Y = lBottom
					.Tv = 1
				End If
				.Z = 0
				.Rhw = 0
				.Tu = .X / lBeyondWidth
				.X += lOffset
				.Color = System.Drawing.Color.FromArgb(0, 255, 255, 255).ToArgb
			End With
			mlBeyondAlpha(X) = 0
		Next X

		'Now, for the Protocol Verts...
		lOffset = oDevice.PresentationParameters.BackBufferWidth \ 2
		lOffset += 42
		Dim lProtocolWidth As Int32 = oDevice.PresentationParameters.BackBufferWidth - lOffset
		lVertWidth = lProtocolWidth \ (muProtocolVerts.GetUpperBound(0) \ 2)
		lVertHeight = CInt((lProtocolWidth / 512.0F) * 64)

		lTop = 80 - (lVertHeight \ 2)
		lBottom = lTop + lVertHeight
		For X As Int32 = 0 To muProtocolVerts.GetUpperBound(0)
			With muProtocolVerts(X)
				.X = ((X \ 2) * lVertWidth)
				If X Mod 2 = 0 Then
					.Y = lTop
					.Tv = 0
				Else
					.Y = lBottom
					.Tv = 1
				End If
				.Z = 0
				.Rhw = 0
				.Tu = .X / lProtocolWidth
				.X += lOffset
				.Color = System.Drawing.Color.FromArgb(0, 255, 255, 255).ToArgb
			End With
			mlProtocolAlpha(X) = 0
		Next X
	End Sub

	Public Function Render(ByVal bFadeout As Boolean, ByRef oDevice As Device) As Boolean
		If bFadeout = True Then
			Dim bDone As Boolean = True
			For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
				Dim lValue As Int32 = mlBeyondAlpha(X)
				lValue -= 10
				If lValue < 0 Then lValue = 0
				If lValue <> 0 Then bDone = False
				mlBeyondAlpha(X) = lValue
				muBeyondVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
				lValue = mlProtocolAlpha(X)
				lValue -= 10
				If lValue < 0 Then lValue = 0
				If lValue <> 0 Then bDone = False
				mlProtocolAlpha(X) = lValue
				muProtocolVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
			Next X
			mlSpriteAlpha -= 15
			If mlSpriteAlpha < 0 Then mlSpriteAlpha = 0
			If bDone = True Then Return True
		ElseIf mbFadedIn = False Then
			Dim bDone As Boolean = True
			For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
				Dim lValue As Int32 = mlBeyondAlpha(X)
				lValue += 1
				If lValue > 255 Then lValue = 255
				If lValue <> 255 Then bDone = False
				mlBeyondAlpha(X) = lValue
				muBeyondVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
				lValue = mlProtocolAlpha(X)
				lValue += 1
				If lValue > 255 Then lValue = 255
				If lValue <> 255 Then bDone = False
				mlProtocolAlpha(X) = lValue
				muProtocolVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
			Next
			mlSpriteAlpha += 5
			If mlSpriteAlpha > 255 Then mlSpriteAlpha = 255
			If bDone = True Then mbFadedIn = True
		Else
			For X As Int32 = 0 To muBeyondVerts.GetUpperBound(0)
				Dim lValue As Int32 = mlBeyondAlpha(X)
				lValue += CInt(Rnd() * 32) - 16
				If lValue < 32 Then lValue = 32
				If lValue > 255 Then lValue = 255
				lValue = 255
				mlBeyondAlpha(X) = lValue
				muBeyondVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb

				lValue = mlProtocolAlpha(X)
				lValue += CInt(Rnd() * 32) - 16
				If lValue < 32 Then lValue = 32
				If lValue > 255 Then lValue = 255
				lValue = 255
				mlProtocolAlpha(X) = lValue
				muProtocolVerts(X).Color = System.Drawing.Color.FromArgb(lValue, 255, 255, 255).ToArgb
			Next X
		End If

		If moBeyondTex Is Nothing OrElse moBeyondTex.Disposed = True Then
			Device.IsUsingEventHandlers = False
			moBeyondTex = moResMgr.GetTexture("Beyond.dds", EpicaResourceManager.eGetTextureType.NoSpecifics, "Misc.pak")
			Device.IsUsingEventHandlers = True
		End If
		If moProtocolTex Is Nothing OrElse moProtocolTex.Disposed = True Then
			Device.IsUsingEventHandlers = False
			moProtocolTex = moResMgr.GetTexture("Protocol.dds", EpicaResourceManager.eGetTextureType.NoSpecifics, "Misc.pak")
			Device.IsUsingEventHandlers = True
		End If
		If moBPTex Is Nothing OrElse moBPTex.Disposed = True Then
			Device.IsUsingEventHandlers = False
			moBPTex = moResMgr.GetTexture("BP.dds", EpicaResourceManager.eGetTextureType.NoSpecifics, "Misc.pak")
			Device.IsUsingEventHandlers = True
		End If

		With oDevice
			.RenderState.CullMode = Cull.None
			.Transform.World = Matrix.Identity

			.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
			.SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
			.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

			.RenderState.SourceBlend = Blend.SourceAlpha
			.RenderState.DestinationBlend = Blend.One
			.RenderState.AlphaBlendEnable = True
			'.RenderState.ZBufferWriteEnable = False
			'.RenderState.Lighting = False

			.VertexFormat = CustomVertex.TransformedColoredTextured.Format
			.SetTexture(0, moBeyondTex)
			.DrawUserPrimitives(PrimitiveType.TriangleStrip, muBeyondVerts.GetUpperBound(0) - 1, muBeyondVerts)
			.SetTexture(0, moProtocolTex)
			.DrawUserPrimitives(PrimitiveType.TriangleStrip, muProtocolVerts.GetUpperBound(0) - 1, muProtocolVerts)

			'Then, reset our device...
			'.RenderState.ZBufferWriteEnable = True
			'.RenderState.Lighting = True
			'.RenderState.PointSpriteEnable = False
			.RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
			.RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
			.RenderState.AlphaBlendEnable = True

			.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
			.SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
			.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
		End With

		If moSprite Is Nothing OrElse moSprite.Disposed = True Then
			Device.IsUsingEventHandlers = False
			moSprite = New Sprite(oDevice)
			Device.IsUsingEventHandlers = True
		End If

		Dim rcDest As Rectangle
		rcDest.X = oDevice.PresentationParameters.BackBufferWidth \ 2 - 64
		rcDest.Y = 16
		rcDest.Width = 128
		rcDest.Height = 128
		moSprite.Begin(SpriteFlags.AlphaBlend)
		moSprite.Draw2D(moBPTex, Rectangle.Empty, rcDest, Point.Empty, 0, rcDest.Location, System.Drawing.Color.FromArgb(mlSpriteAlpha, 255, 255, 255))
		moSprite.End()

		Return False
	End Function

	Protected Overrides Sub Finalize()
		If moBeyondTex Is Nothing = False Then moBeyondTex.Dispose()
		moBeyondTex = Nothing
		If moProtocolTex Is Nothing = False Then moProtocolTex.Dispose()
		moProtocolTex = Nothing
		If moBPTex Is Nothing = False Then moBPTex.Dispose()
		moBPTex = Nothing
		If moSprite Is Nothing = False Then moSprite.Dispose()
		moSprite = Nothing
		MyBase.Finalize()
	End Sub
End Class
