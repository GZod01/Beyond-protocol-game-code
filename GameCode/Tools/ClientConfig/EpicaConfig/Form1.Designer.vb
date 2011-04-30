<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.picDisplay = New System.Windows.Forms.PictureBox
        Me.chkSmoothFOW = New System.Windows.Forms.CheckBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnReset = New System.Windows.Forms.Button
        Me.cboDeviceAdapter = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.chkHiResPlanet = New System.Windows.Forms.CheckBox
        Me.cboGlowFX = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.cboTerrTexRes = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboWaterRes = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboTextureResolution = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboFOWRes = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboVertProc = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboFullScreenRes = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkFullScreen = New System.Windows.Forms.CheckBox
        Me.tmrRedraw = New System.Windows.Forms.Timer(Me.components)
        Me.btnToggleDisplay = New System.Windows.Forms.Button
        Me.btnLaunch = New System.Windows.Forms.Button
        Me.picPreRendered = New System.Windows.Forms.PictureBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tbpGeneral = New System.Windows.Forms.TabPage
        Me.btnResetInterface = New System.Windows.Forms.Button
        Me.hscrClipPlane = New System.Windows.Forms.HScrollBar
        Me.Label9 = New System.Windows.Forms.Label
        Me.chkTripleBuffer = New System.Windows.Forms.CheckBox
        Me.chkVerticalSync = New System.Windows.Forms.CheckBox
        Me.tbpTextures = New System.Windows.Forms.TabPage
        Me.chkDeepCosmos = New System.Windows.Forms.CheckBox
        Me.tbpLighting = New System.Windows.Forms.TabPage
        Me.chkIllumPlanet = New System.Windows.Forms.CheckBox
        Me.chkBumpPlanet = New System.Windows.Forms.CheckBox
        Me.chkBumpTerrain = New System.Windows.Forms.CheckBox
        Me.chkPlayerCustom = New System.Windows.Forms.CheckBox
        Me.cboIllumMaps = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.cboLightQuality = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.tbpParticles = New System.Windows.Forms.TabPage
        Me.hscrStarPlanet = New System.Windows.Forms.HScrollBar
        Me.Label16 = New System.Windows.Forms.Label
        Me.hscrStarSpace = New System.Windows.Forms.HScrollBar
        Me.Label15 = New System.Windows.Forms.Label
        Me.chkShieldFX = New System.Windows.Forms.CheckBox
        Me.cboPlanetFX = New System.Windows.Forms.ComboBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.cboBurnFX = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.cboEngineFX = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.btnDefaults = New System.Windows.Forms.Button
        CType(Me.picDisplay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPreRendered, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.tbpGeneral.SuspendLayout()
        Me.tbpTextures.SuspendLayout()
        Me.tbpLighting.SuspendLayout()
        Me.tbpParticles.SuspendLayout()
        Me.SuspendLayout()
        '
        'picDisplay
        '
        Me.picDisplay.Location = New System.Drawing.Point(12, 12)
        Me.picDisplay.Name = "picDisplay"
        Me.picDisplay.Size = New System.Drawing.Size(512, 384)
        Me.picDisplay.TabIndex = 0
        Me.picDisplay.TabStop = False
        '
        'chkSmoothFOW
        '
        Me.chkSmoothFOW.AutoSize = True
        Me.chkSmoothFOW.Location = New System.Drawing.Point(11, 114)
        Me.chkSmoothFOW.Name = "chkSmoothFOW"
        Me.chkSmoothFOW.Size = New System.Drawing.Size(118, 17)
        Me.chkSmoothFOW.TabIndex = 19
        Me.chkSmoothFOW.Text = "Smooth Fog-of-War"
        Me.chkSmoothFOW.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(714, 367)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 18
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(530, 367)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(75, 23)
        Me.btnReset.TabIndex = 4
        Me.btnReset.Text = "Reset"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'cboDeviceAdapter
        '
        Me.cboDeviceAdapter.FormattingEnabled = True
        Me.cboDeviceAdapter.Location = New System.Drawing.Point(11, 29)
        Me.cboDeviceAdapter.Name = "cboDeviceAdapter"
        Me.cboDeviceAdapter.Size = New System.Drawing.Size(235, 21)
        Me.cboDeviceAdapter.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 13)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(84, 13)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Device Adapter:"
        '
        'chkHiResPlanet
        '
        Me.chkHiResPlanet.AutoSize = True
        Me.chkHiResPlanet.Location = New System.Drawing.Point(11, 91)
        Me.chkHiResPlanet.Name = "chkHiResPlanet"
        Me.chkHiResPlanet.Size = New System.Drawing.Size(178, 17)
        Me.chkHiResPlanet.TabIndex = 16
        Me.chkHiResPlanet.Text = "High Resolution Planet Textures"
        Me.chkHiResPlanet.UseVisualStyleBackColor = True
        '
        'cboGlowFX
        '
        Me.cboGlowFX.FormattingEnabled = True
        Me.cboGlowFX.Location = New System.Drawing.Point(125, 68)
        Me.cboGlowFX.Name = "cboGlowFX"
        Me.cboGlowFX.Size = New System.Drawing.Size(121, 21)
        Me.cboGlowFX.TabIndex = 14
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(8, 71)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(92, 13)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Glow FX Intensity:"
        '
        'cboTerrTexRes
        '
        Me.cboTerrTexRes.FormattingEnabled = True
        Me.cboTerrTexRes.Location = New System.Drawing.Point(125, 37)
        Me.cboTerrTexRes.Name = "cboTerrTexRes"
        Me.cboTerrTexRes.Size = New System.Drawing.Size(121, 21)
        Me.cboTerrTexRes.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(90, 13)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Terrain Texturing:"
        '
        'cboWaterRes
        '
        Me.cboWaterRes.FormattingEnabled = True
        Me.cboWaterRes.Location = New System.Drawing.Point(125, 127)
        Me.cboWaterRes.Name = "cboWaterRes"
        Me.cboWaterRes.Size = New System.Drawing.Size(121, 21)
        Me.cboWaterRes.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 130)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(92, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Water Resolution:"
        '
        'cboTextureResolution
        '
        Me.cboTextureResolution.FormattingEnabled = True
        Me.cboTextureResolution.Location = New System.Drawing.Point(125, 10)
        Me.cboTextureResolution.Name = "cboTextureResolution"
        Me.cboTextureResolution.Size = New System.Drawing.Size(121, 21)
        Me.cboTextureResolution.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(99, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Texture Resolution:"
        '
        'cboFOWRes
        '
        Me.cboFOWRes.FormattingEnabled = True
        Me.cboFOWRes.Location = New System.Drawing.Point(125, 64)
        Me.cboFOWRes.Name = "cboFOWRes"
        Me.cboFOWRes.Size = New System.Drawing.Size(121, 21)
        Me.cboFOWRes.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(116, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Fog of War Resolution:"
        '
        'cboVertProc
        '
        Me.cboVertProc.FormattingEnabled = True
        Me.cboVertProc.Location = New System.Drawing.Point(125, 100)
        Me.cboVertProc.Name = "cboVertProc"
        Me.cboVertProc.Size = New System.Drawing.Size(121, 21)
        Me.cboVertProc.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 103)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "T&&L Mode:"
        '
        'cboFullScreenRes
        '
        Me.cboFullScreenRes.FormattingEnabled = True
        Me.cboFullScreenRes.Location = New System.Drawing.Point(125, 73)
        Me.cboFullScreenRes.Name = "cboFullScreenRes"
        Me.cboFullScreenRes.Size = New System.Drawing.Size(121, 21)
        Me.cboFullScreenRes.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 76)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(111, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Fullscreen Resolution:"
        '
        'chkFullScreen
        '
        Me.chkFullScreen.AutoSize = True
        Me.chkFullScreen.Location = New System.Drawing.Point(11, 56)
        Me.chkFullScreen.Name = "chkFullScreen"
        Me.chkFullScreen.Size = New System.Drawing.Size(74, 17)
        Me.chkFullScreen.TabIndex = 0
        Me.chkFullScreen.Text = "Fullscreen"
        Me.chkFullScreen.UseVisualStyleBackColor = True
        '
        'tmrRedraw
        '
        Me.tmrRedraw.Interval = 5
        '
        'btnToggleDisplay
        '
        Me.btnToggleDisplay.Location = New System.Drawing.Point(175, 402)
        Me.btnToggleDisplay.Name = "btnToggleDisplay"
        Me.btnToggleDisplay.Size = New System.Drawing.Size(193, 23)
        Me.btnToggleDisplay.TabIndex = 3
        Me.btnToggleDisplay.Text = "Show Compare Image"
        Me.btnToggleDisplay.UseVisualStyleBackColor = True
        '
        'btnLaunch
        '
        Me.btnLaunch.Location = New System.Drawing.Point(567, 402)
        Me.btnLaunch.Name = "btnLaunch"
        Me.btnLaunch.Size = New System.Drawing.Size(193, 23)
        Me.btnLaunch.TabIndex = 4
        Me.btnLaunch.Text = "Launch Game"
        Me.btnLaunch.UseVisualStyleBackColor = True
        Me.btnLaunch.Visible = False
        '
        'picPreRendered
        '
        Me.picPreRendered.Image = Global.ClientConfig.My.Resources.Resources.EpicaCompare
        Me.picPreRendered.Location = New System.Drawing.Point(12, 12)
        Me.picPreRendered.Name = "picPreRendered"
        Me.picPreRendered.Size = New System.Drawing.Size(512, 384)
        Me.picPreRendered.TabIndex = 5
        Me.picPreRendered.TabStop = False
        Me.picPreRendered.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tbpGeneral)
        Me.TabControl1.Controls.Add(Me.tbpTextures)
        Me.TabControl1.Controls.Add(Me.tbpLighting)
        Me.TabControl1.Controls.Add(Me.tbpParticles)
        Me.TabControl1.Location = New System.Drawing.Point(530, 12)
        Me.TabControl1.Multiline = True
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(259, 353)
        Me.TabControl1.TabIndex = 6
        '
        'tbpGeneral
        '
        Me.tbpGeneral.Controls.Add(Me.btnResetInterface)
        Me.tbpGeneral.Controls.Add(Me.hscrClipPlane)
        Me.tbpGeneral.Controls.Add(Me.Label9)
        Me.tbpGeneral.Controls.Add(Me.chkTripleBuffer)
        Me.tbpGeneral.Controls.Add(Me.chkVerticalSync)
        Me.tbpGeneral.Controls.Add(Me.cboWaterRes)
        Me.tbpGeneral.Controls.Add(Me.chkFullScreen)
        Me.tbpGeneral.Controls.Add(Me.Label1)
        Me.tbpGeneral.Controls.Add(Me.cboDeviceAdapter)
        Me.tbpGeneral.Controls.Add(Me.cboFullScreenRes)
        Me.tbpGeneral.Controls.Add(Me.Label8)
        Me.tbpGeneral.Controls.Add(Me.Label2)
        Me.tbpGeneral.Controls.Add(Me.cboVertProc)
        Me.tbpGeneral.Controls.Add(Me.Label5)
        Me.tbpGeneral.Location = New System.Drawing.Point(4, 22)
        Me.tbpGeneral.Name = "tbpGeneral"
        Me.tbpGeneral.Padding = New System.Windows.Forms.Padding(3)
        Me.tbpGeneral.Size = New System.Drawing.Size(251, 327)
        Me.tbpGeneral.TabIndex = 0
        Me.tbpGeneral.Text = "General"
        Me.tbpGeneral.UseVisualStyleBackColor = True
        '
        'btnResetInterface
        '
        Me.btnResetInterface.Location = New System.Drawing.Point(11, 217)
        Me.btnResetInterface.Name = "btnResetInterface"
        Me.btnResetInterface.Size = New System.Drawing.Size(131, 23)
        Me.btnResetInterface.TabIndex = 20
        Me.btnResetInterface.Text = "Reset Interface Colors"
        Me.btnResetInterface.UseVisualStyleBackColor = True
        '
        'hscrClipPlane
        '
        Me.hscrClipPlane.Location = New System.Drawing.Point(125, 198)
        Me.hscrClipPlane.Maximum = 400
        Me.hscrClipPlane.Minimum = 100
        Me.hscrClipPlane.Name = "hscrClipPlane"
        Me.hscrClipPlane.Size = New System.Drawing.Size(121, 16)
        Me.hscrClipPlane.TabIndex = 21
        Me.hscrClipPlane.Value = 300
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 198)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 13)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "Entity Clip Plane:"
        '
        'chkTripleBuffer
        '
        Me.chkTripleBuffer.AutoSize = True
        Me.chkTripleBuffer.Location = New System.Drawing.Point(11, 178)
        Me.chkTripleBuffer.Name = "chkTripleBuffer"
        Me.chkTripleBuffer.Size = New System.Drawing.Size(83, 17)
        Me.chkTripleBuffer.TabIndex = 19
        Me.chkTripleBuffer.Text = "Triple Buffer"
        Me.chkTripleBuffer.UseVisualStyleBackColor = True
        '
        'chkVerticalSync
        '
        Me.chkVerticalSync.AutoSize = True
        Me.chkVerticalSync.Location = New System.Drawing.Point(11, 155)
        Me.chkVerticalSync.Name = "chkVerticalSync"
        Me.chkVerticalSync.Size = New System.Drawing.Size(88, 17)
        Me.chkVerticalSync.TabIndex = 18
        Me.chkVerticalSync.Text = "Vertical Sync"
        Me.chkVerticalSync.UseVisualStyleBackColor = True
        '
        'tbpTextures
        '
        Me.tbpTextures.Controls.Add(Me.chkDeepCosmos)
        Me.tbpTextures.Controls.Add(Me.chkSmoothFOW)
        Me.tbpTextures.Controls.Add(Me.cboTextureResolution)
        Me.tbpTextures.Controls.Add(Me.Label4)
        Me.tbpTextures.Controls.Add(Me.Label6)
        Me.tbpTextures.Controls.Add(Me.chkHiResPlanet)
        Me.tbpTextures.Controls.Add(Me.cboTerrTexRes)
        Me.tbpTextures.Controls.Add(Me.cboFOWRes)
        Me.tbpTextures.Controls.Add(Me.Label3)
        Me.tbpTextures.Location = New System.Drawing.Point(4, 22)
        Me.tbpTextures.Name = "tbpTextures"
        Me.tbpTextures.Padding = New System.Windows.Forms.Padding(3)
        Me.tbpTextures.Size = New System.Drawing.Size(251, 327)
        Me.tbpTextures.TabIndex = 1
        Me.tbpTextures.Text = "Textures"
        Me.tbpTextures.UseVisualStyleBackColor = True
        '
        'chkDeepCosmos
        '
        Me.chkDeepCosmos.AutoSize = True
        Me.chkDeepCosmos.Location = New System.Drawing.Point(11, 137)
        Me.chkDeepCosmos.Name = "chkDeepCosmos"
        Me.chkDeepCosmos.Size = New System.Drawing.Size(130, 17)
        Me.chkDeepCosmos.TabIndex = 20
        Me.chkDeepCosmos.Text = "Render Deep Cosmos"
        Me.chkDeepCosmos.UseVisualStyleBackColor = True
        '
        'tbpLighting
        '
        Me.tbpLighting.Controls.Add(Me.chkIllumPlanet)
        Me.tbpLighting.Controls.Add(Me.chkBumpPlanet)
        Me.tbpLighting.Controls.Add(Me.chkBumpTerrain)
        Me.tbpLighting.Controls.Add(Me.chkPlayerCustom)
        Me.tbpLighting.Controls.Add(Me.cboIllumMaps)
        Me.tbpLighting.Controls.Add(Me.Label11)
        Me.tbpLighting.Controls.Add(Me.cboGlowFX)
        Me.tbpLighting.Controls.Add(Me.Label7)
        Me.tbpLighting.Controls.Add(Me.cboLightQuality)
        Me.tbpLighting.Controls.Add(Me.Label10)
        Me.tbpLighting.Location = New System.Drawing.Point(4, 22)
        Me.tbpLighting.Name = "tbpLighting"
        Me.tbpLighting.Size = New System.Drawing.Size(251, 327)
        Me.tbpLighting.TabIndex = 2
        Me.tbpLighting.Text = "Lighting"
        Me.tbpLighting.UseVisualStyleBackColor = True
        '
        'chkIllumPlanet
        '
        Me.chkIllumPlanet.AutoSize = True
        Me.chkIllumPlanet.Location = New System.Drawing.Point(11, 164)
        Me.chkIllumPlanet.Name = "chkIllumPlanet"
        Me.chkIllumPlanet.Size = New System.Drawing.Size(139, 17)
        Me.chkIllumPlanet.TabIndex = 20
        Me.chkIllumPlanet.Text = "Illuminate Planet Terrain"
        Me.chkIllumPlanet.UseVisualStyleBackColor = True
        '
        'chkBumpPlanet
        '
        Me.chkBumpPlanet.AutoSize = True
        Me.chkBumpPlanet.Location = New System.Drawing.Point(11, 141)
        Me.chkBumpPlanet.Name = "chkBumpPlanet"
        Me.chkBumpPlanet.Size = New System.Drawing.Size(147, 17)
        Me.chkBumpPlanet.TabIndex = 19
        Me.chkBumpPlanet.Text = "Bump Map Planet Models"
        Me.chkBumpPlanet.UseVisualStyleBackColor = True
        '
        'chkBumpTerrain
        '
        Me.chkBumpTerrain.AutoSize = True
        Me.chkBumpTerrain.Location = New System.Drawing.Point(11, 118)
        Me.chkBumpTerrain.Name = "chkBumpTerrain"
        Me.chkBumpTerrain.Size = New System.Drawing.Size(113, 17)
        Me.chkBumpTerrain.TabIndex = 18
        Me.chkBumpTerrain.Text = "Bump Map Terrain"
        Me.chkBumpTerrain.UseVisualStyleBackColor = True
        '
        'chkPlayerCustom
        '
        Me.chkPlayerCustom.AutoSize = True
        Me.chkPlayerCustom.Location = New System.Drawing.Point(11, 95)
        Me.chkPlayerCustom.Name = "chkPlayerCustom"
        Me.chkPlayerCustom.Size = New System.Drawing.Size(163, 17)
        Me.chkPlayerCustom.TabIndex = 17
        Me.chkPlayerCustom.Text = "Render Player Custom Colors"
        Me.chkPlayerCustom.UseVisualStyleBackColor = True
        '
        'cboIllumMaps
        '
        Me.cboIllumMaps.FormattingEnabled = True
        Me.cboIllumMaps.Location = New System.Drawing.Point(125, 41)
        Me.cboIllumMaps.Name = "cboIllumMaps"
        Me.cboIllumMaps.Size = New System.Drawing.Size(121, 21)
        Me.cboIllumMaps.TabIndex = 12
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(8, 44)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(91, 13)
        Me.Label11.TabIndex = 11
        Me.Label11.Text = "Illumination Maps:"
        '
        'cboLightQuality
        '
        Me.cboLightQuality.FormattingEnabled = True
        Me.cboLightQuality.Location = New System.Drawing.Point(125, 14)
        Me.cboLightQuality.Name = "cboLightQuality"
        Me.cboLightQuality.Size = New System.Drawing.Size(121, 21)
        Me.cboLightQuality.TabIndex = 10
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 17)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(68, 13)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "Light Quality:"
        '
        'tbpParticles
        '
        Me.tbpParticles.Controls.Add(Me.hscrStarPlanet)
        Me.tbpParticles.Controls.Add(Me.Label16)
        Me.tbpParticles.Controls.Add(Me.hscrStarSpace)
        Me.tbpParticles.Controls.Add(Me.Label15)
        Me.tbpParticles.Controls.Add(Me.chkShieldFX)
        Me.tbpParticles.Controls.Add(Me.cboPlanetFX)
        Me.tbpParticles.Controls.Add(Me.Label14)
        Me.tbpParticles.Controls.Add(Me.cboBurnFX)
        Me.tbpParticles.Controls.Add(Me.Label13)
        Me.tbpParticles.Controls.Add(Me.cboEngineFX)
        Me.tbpParticles.Controls.Add(Me.Label12)
        Me.tbpParticles.Location = New System.Drawing.Point(4, 22)
        Me.tbpParticles.Name = "tbpParticles"
        Me.tbpParticles.Size = New System.Drawing.Size(251, 327)
        Me.tbpParticles.TabIndex = 3
        Me.tbpParticles.Text = "Particles"
        Me.tbpParticles.UseVisualStyleBackColor = True
        '
        'hscrStarPlanet
        '
        Me.hscrStarPlanet.Location = New System.Drawing.Point(125, 142)
        Me.hscrStarPlanet.Maximum = 400
        Me.hscrStarPlanet.Minimum = 10
        Me.hscrStarPlanet.Name = "hscrStarPlanet"
        Me.hscrStarPlanet.Size = New System.Drawing.Size(121, 16)
        Me.hscrStarPlanet.TabIndex = 25
        Me.hscrStarPlanet.Value = 10
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(8, 142)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(103, 13)
        Me.Label16.TabIndex = 24
        Me.Label16.Text = "Star Density (Planet)"
        '
        'hscrStarSpace
        '
        Me.hscrStarSpace.Location = New System.Drawing.Point(125, 116)
        Me.hscrStarSpace.Maximum = 400
        Me.hscrStarSpace.Minimum = 10
        Me.hscrStarSpace.Name = "hscrStarSpace"
        Me.hscrStarSpace.Size = New System.Drawing.Size(121, 16)
        Me.hscrStarSpace.TabIndex = 23
        Me.hscrStarSpace.Value = 10
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(8, 119)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(104, 13)
        Me.Label15.TabIndex = 22
        Me.Label15.Text = "Star Density (Space)"
        '
        'chkShieldFX
        '
        Me.chkShieldFX.AutoSize = True
        Me.chkShieldFX.Location = New System.Drawing.Point(11, 96)
        Me.chkShieldFX.Name = "chkShieldFX"
        Me.chkShieldFX.Size = New System.Drawing.Size(109, 17)
        Me.chkShieldFX.TabIndex = 19
        Me.chkShieldFX.Text = "Render Shield FX"
        Me.chkShieldFX.UseVisualStyleBackColor = True
        '
        'cboPlanetFX
        '
        Me.cboPlanetFX.FormattingEnabled = True
        Me.cboPlanetFX.Location = New System.Drawing.Point(125, 69)
        Me.cboPlanetFX.Name = "cboPlanetFX"
        Me.cboPlanetFX.Size = New System.Drawing.Size(121, 21)
        Me.cboPlanetFX.TabIndex = 16
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(8, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(56, 13)
        Me.Label14.TabIndex = 15
        Me.Label14.Text = "Planet FX:"
        '
        'cboBurnFX
        '
        Me.cboBurnFX.FormattingEnabled = True
        Me.cboBurnFX.Location = New System.Drawing.Point(125, 42)
        Me.cboBurnFX.Name = "cboBurnFX"
        Me.cboBurnFX.Size = New System.Drawing.Size(121, 21)
        Me.cboBurnFX.TabIndex = 14
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(8, 45)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(48, 13)
        Me.Label13.TabIndex = 13
        Me.Label13.Text = "Burn FX:"
        '
        'cboEngineFX
        '
        Me.cboEngineFX.FormattingEnabled = True
        Me.cboEngineFX.Location = New System.Drawing.Point(125, 15)
        Me.cboEngineFX.Name = "cboEngineFX"
        Me.cboEngineFX.Size = New System.Drawing.Size(121, 21)
        Me.cboEngineFX.TabIndex = 12
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(8, 18)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(59, 13)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "Engine FX:"
        '
        'btnDefaults
        '
        Me.btnDefaults.Location = New System.Drawing.Point(622, 367)
        Me.btnDefaults.Name = "btnDefaults"
        Me.btnDefaults.Size = New System.Drawing.Size(75, 23)
        Me.btnDefaults.TabIndex = 19
        Me.btnDefaults.Text = "Defaults"
        Me.btnDefaults.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 435)
        Me.Controls.Add(Me.btnDefaults)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.picPreRendered)
        Me.Controls.Add(Me.btnLaunch)
        Me.Controls.Add(Me.btnToggleDisplay)
        Me.Controls.Add(Me.picDisplay)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Client Configuration Utility"
        CType(Me.picDisplay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPreRendered, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.tbpGeneral.ResumeLayout(False)
        Me.tbpGeneral.PerformLayout()
        Me.tbpTextures.ResumeLayout(False)
        Me.tbpTextures.PerformLayout()
        Me.tbpLighting.ResumeLayout(False)
        Me.tbpLighting.PerformLayout()
        Me.tbpParticles.ResumeLayout(False)
        Me.tbpParticles.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents picDisplay As System.Windows.Forms.PictureBox
    Friend WithEvents cboFullScreenRes As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkFullScreen As System.Windows.Forms.CheckBox
    Friend WithEvents cboGlowFX As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cboTerrTexRes As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboWaterRes As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboTextureResolution As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboFOWRes As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboVertProc As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents chkHiResPlanet As System.Windows.Forms.CheckBox
    Friend WithEvents cboDeviceAdapter As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents tmrRedraw As System.Windows.Forms.Timer
    Friend WithEvents btnToggleDisplay As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents btnLaunch As System.Windows.Forms.Button
	Friend WithEvents picPreRendered As System.Windows.Forms.PictureBox
    Friend WithEvents chkSmoothFOW As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tbpGeneral As System.Windows.Forms.TabPage
    Friend WithEvents tbpTextures As System.Windows.Forms.TabPage
    Friend WithEvents tbpLighting As System.Windows.Forms.TabPage
    Friend WithEvents tbpParticles As System.Windows.Forms.TabPage
    Friend WithEvents hscrClipPlane As System.Windows.Forms.HScrollBar
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents chkTripleBuffer As System.Windows.Forms.CheckBox
    Friend WithEvents chkVerticalSync As System.Windows.Forms.CheckBox
    Friend WithEvents chkDeepCosmos As System.Windows.Forms.CheckBox
    Friend WithEvents cboLightQuality As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents chkIllumPlanet As System.Windows.Forms.CheckBox
    Friend WithEvents chkBumpPlanet As System.Windows.Forms.CheckBox
    Friend WithEvents chkBumpTerrain As System.Windows.Forms.CheckBox
    Friend WithEvents chkPlayerCustom As System.Windows.Forms.CheckBox
    Friend WithEvents cboIllumMaps As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents hscrStarPlanet As System.Windows.Forms.HScrollBar
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents hscrStarSpace As System.Windows.Forms.HScrollBar
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents chkShieldFX As System.Windows.Forms.CheckBox
    Friend WithEvents cboPlanetFX As System.Windows.Forms.ComboBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents cboBurnFX As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cboEngineFX As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents btnResetInterface As System.Windows.Forms.Button
    Friend WithEvents btnDefaults As System.Windows.Forms.Button

End Class
