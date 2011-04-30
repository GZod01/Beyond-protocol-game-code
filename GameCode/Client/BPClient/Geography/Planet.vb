
Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class Planet
    Inherits Base_GUID

    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Public PlanetName As String
    Public MapTypeID As Byte        'the planet typeid
    Private myPlanetSizeID As Byte     'used for determining map size, 0-tiny, 1-small, 2-medium, 3-large, 4-huge. Maybe able to remove and base off of Radius
    Private miPlanetRadius As Int16    'might be able to remove sizeID and base the map size off of radius
    Public LocX As Int32
    Public LocY As Int32
    Public LocZ As Int32
    Public Vegetation As Byte
    Public Atmosphere As Byte
    Public Hydrosphere As Byte
    Public Gravity As Byte
    Public SurfaceTemperature As Byte
    Public RotationDelay As Int16   'cycles between incrementing rotation angle

    Public ParentSystem As SolarSystem

    Public Shared moSprite As Sprite
    Private Shared moPGlowSpr As Sprite
    Private Shared moStarSprite As Sprite
    Private Shared moBillboardMesh As Mesh
    Public Shared Sub EnsureSpritesValid(ByVal oDevice As Device)
        If moSprite Is Nothing OrElse moSprite.Disposed = True Then
            Device.IsUsingEventHandlers = False
            moSprite = New Sprite(oDevice)
            'AddHandler moSprite.Disposing, AddressOf SpriteDispose
            'AddHandler moSprite.Lost, AddressOf SpriteLost
            'AddHandler moSprite.Reset, AddressOf SpriteLost
            Device.IsUsingEventHandlers = True
        End If
    End Sub
    Public Shared Sub ReleaseSprites()
        Try
            If moSprite Is Nothing = False Then moSprite.Dispose()
            If moPGlowSpr Is Nothing = False Then moPGlowSpr.Dispose()
            If moStarSprite Is Nothing = False Then moStarSprite.Dispose()
            If moBumpPlanet Is Nothing = False Then moBumpPlanet.DisposeMe()
            If moBillboardMesh Is Nothing = False Then moBillboardMesh.Dispose()
            If moPlanetGlowTex Is Nothing = False Then moPlanetGlowTex.Dispose()
            If moMapTex Is Nothing = False Then moMapTex.Dispose()
            If moColorSphere Is Nothing = False Then moColorSphere.Dispose()

            If moCloudTex Is Nothing = False Then moCloudTex.Dispose()
            If mvbClouds Is Nothing = False Then mvbClouds.Dispose()
            If mibClouds Is Nothing = False Then mibClouds.Dispose()
            If moLensFlare Is Nothing = False Then moLensFlare.Dispose()
        Catch
        End Try

        moSprite = Nothing
        moPGlowSpr = Nothing
        moStarSprite = Nothing
        moBumpPlanet = Nothing
        moBillboardMesh = Nothing
        mbSkyboxReady = False
        moPlanetGlowTex = Nothing
        moMapTex = Nothing
        moColorSphere = Nothing
        moCloudTex = Nothing
        mvbClouds = Nothing
        mibClouds = Nothing
        moLensFlare = Nothing

    End Sub

    Private moTerrain As TerrainClass
    Private moEnvirColor As System.Drawing.Color
    Private miSpeedMod(,,) As Int16
    Private mlFullSizeWH As Int32
    Private mlDoubleWH As Int32
    Private mbSpeedModRdy As Boolean = False

	Public Shared fDayTimeRatio As Single

    Public Shared mbShowLightMap As Boolean = True

    Public lLastBBRequest As Int32 = 0
    Public Structure GuildBillboardSlot
        Public lGuildID As Int32
        Public iRecruitFlags As Int16
        Public lIcon As Int32
        Public BidAmount As Int32
        Public BillboardText As String
        Public oBBTex As Texture
        Private msOrigGuildName As String
        Public Sub CheckBBTex()
            If msOrigGuildName <> GetCacheObjectValue(lGuildID, ObjectType.eGuild) Then
                If oBBTex Is Nothing = False Then oBBTex.Dispose()
                oBBTex = Nothing
            End If
        End Sub
        Public Sub GenerateBBTex(ByRef oDev As Device)
            msOrigGuildName = GetCacheObjectValue(lGuildID, ObjectType.eGuild)

            Device.IsUsingEventHandlers = False
            'If ctlDiplomacy.moIconBack Is Nothing OrElse ctlDiplomacy.moIconBack.Disposed = True Then ctlDiplomacy.moIconBack = goResMgr.GetTexture("DipPlayerBack.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
            'If ctlDiplomacy.moIconFore Is Nothing OrElse ctlDiplomacy.moIconFore.Disposed = True Then ctlDiplomacy.moIconFore = goResMgr.GetTexture("DipPlayerFore.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
            'If ctlDiplomacy.moSprite Is Nothing OrElse ctlDiplomacy.moSprite.Disposed = True Then
            '    ctlDiplomacy.moSprite = New Sprite(oDev)
            'End If
            oBBTex = New Texture(oDev, 128, 128, 1, Usage.RenderTarget, Format.A8R8G8B8, Pool.Default)
            Device.IsUsingEventHandlers = True

            oDev.RenderState.ZBufferEnable = False
            Dim oOrigSurf As Surface = oDev.GetRenderTarget(0)
            oDev.SetRenderTarget(0, oBBTex.GetSurfaceLevel(0))
            oDev.Clear(ClearFlags.Target, System.Drawing.Color.FromArgb(0, 0, 0, 0), 1, 0)

            'Dim yBackImg As Byte
            'Dim yBackClr As Byte
            'Dim yFore1Img As Byte
            'Dim yFore1Clr As Byte
            'Dim yFore2Img As Byte
            'Dim yFore2Clr As Byte

            'PlayerIconManager.FillIconValues(lIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

            'Dim rcBack As Rectangle = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
            'Dim rcFore1 As Rectangle = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
            'Dim rcFore2 As Rectangle = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

            'Dim clrBack As System.Drawing.Color = PlayerIconManager.GetColorValue(yBackClr)
            'Dim clrFore1 As System.Drawing.Color = PlayerIconManager.GetColorValue(yFore1Clr)
            'Dim clrFore2 As System.Drawing.Color = PlayerIconManager.GetColorValue(yFore2Clr)

            'Dim rcDest As Rectangle = New Rectangle(0, 0, 128, 128)
            'Dim ptDest As Point = New Point(0, 0)
 
            'BPSprite.Draw2DOnce(oDev, ctlDiplomacy.moIconBack, rcBack, rcDest, clrBack, 256, 256)
            'BPSprite.Draw2DOnce(oDev, ctlDiplomacy.moIconFore, rcFore1, rcDest, clrFore1, 512, 512)
            'BPSprite.Draw2DOnce(oDev, ctlDiplomacy.moIconFore, rcFore2, rcDest, clrFore2, 512, 512)

            BPFont.DrawText(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0), msOrigGuildName, New Rectangle(0, 0, 128, 128), DrawTextFormat.Center Or DrawTextFormat.VerticalCenter Or DrawTextFormat.WordBreak, System.Drawing.Color.FromArgb(255, 0, 255, 0))

            oDev.SetRenderTarget(0, oOrigSurf)
            oDev.RenderState.ZBufferEnable = True
            oOrigSurf.Dispose() : oOrigSurf = Nothing
        End Sub
    End Structure
    Public uBillboards(5) As GuildBillboardSlot 

	Private Shared moTutorialPlanet As Planet = Nothing
    Public Shared Function GetTutorialPlanet() As Planet
        If moTutorialPlanet Is Nothing Then
            moTutorialPlanet = New Planet(glPlayerID + 500000000, 0, PlanetType.eGeoPlastic)
            With moTutorialPlanet
                .PlanetName = "Tutorial One"
                .LocX = 100000
                .LocY = 0
                .LocZ = 0
                .Vegetation = 0
                .Atmosphere = 100
                .Hydrosphere = 0
                .Gravity = 40
                .SurfaceTemperature = 90
                .RotationDelay = 120
                .ObjTypeID = ObjectType.ePlanet
                .ParentSystem = New SolarSystem()
            End With
            With moTutorialPlanet.ParentSystem
                .AddPlanet(moTutorialPlanet)
                .CurrentPlanetIdx = 0
                .ObjectID = glPlayerID + 500000000
                .ObjTypeID = ObjectType.eSolarSystem
                .StarType1Idx = 0
                For X As Int32 = 0 To glStarTypeUB
                    If goStarTypes(X).StarTypeID = 8 Then
                        .StarType1Idx = X
                    End If
                Next X
                .SystemName = "Tutorial System"
            End With
        End If
        Return moTutorialPlanet
    End Function

    Private mlOwnerID As Int32 = -1
    Private mlRingMineralID As Int32 = -1
    Private mlRingConcentration As Int32 = 0
    Private mlColonyCount As Int32 = -1
    Private mlLastOwnerIDRequest As Int32 = -1
    Public ReadOnly Property OwnerID() As Int32
        Get
            If glCurrentCycle - mlLastOwnerIDRequest > 150 Then
                mlLastOwnerIDRequest = glCurrentCycle
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlanetOwnership).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yMsg)
            End If
            Return mlOwnerID
        End Get
    End Property
    Public ReadOnly Property RingMineralID() As Int32
        Get
            If glCurrentCycle - mlLastOwnerIDRequest > 150 Then
                mlLastOwnerIDRequest = glCurrentCycle
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlanetOwnership).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yMsg)
            End If
            Return mlRingMineralID
        End Get
    End Property
    Public ReadOnly Property RingMineralConcentration() As Int32
        Get
            If glCurrentCycle - mlLastOwnerIDRequest > 150 Then
                mlLastOwnerIDRequest = glCurrentCycle
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlanetOwnership).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yMsg)
            End If
            Return mlRingConcentration
        End Get
    End Property
    Public ReadOnly Property ColonyCount() As Int32
        Get
            If glCurrentCycle - mlLastOwnerIDRequest > 150 Then
                mlLastOwnerIDRequest = glCurrentCycle
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlanetOwnership).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yMsg)
            End If
            Return mlColonyCount
        End Get
    End Property
    Public Sub HandleUpdatePlanetOwnership(ByVal yData() As Byte)
        Dim lPos As Int32 = 6   'msgcode, PlanetID
        mlOwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += 4   'systemid

        mlRingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        mlRingConcentration = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        mlColonyCount = yData(lPos) : lPos += 1 'System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        myPlanetSizeID = yData(lPos) : lPos += 1 'System.BitConverter.ToInt32(yData, lPos) : lPos += 2
    End Sub

#Region "For Space Rendering"
    Private moMesh As Mesh          'planet's mesh
	Private mbMeshTBNReady As Boolean = False
    Private moMaterial As Material  'planet's Material
    Private moTexture As Texture    'Planet's texture - we get this from moTerrain!!!
	Private moLightMap As Texture	'lightmap of the planet - only used for lava planets
    Private moNormalMap As Texture

    Public Sub ReleaseRenderTargets()
        If moTexture Is Nothing = False Then moTexture.Dispose()
        moTexture = Nothing
        If moLightMap Is Nothing = False Then moLightMap.Dispose()
        moLightMap = Nothing
        If moNormalMap Is Nothing = False Then moNormalMap.Dispose()
        moNormalMap = Nothing
        If moMesh Is Nothing = False Then moMesh.Dispose()
        moMesh = Nothing
        If moCloud Is Nothing = False Then moCloud.Dispose()
        moCloud = Nothing
        If uBillboards Is Nothing = False Then
            For X As Int32 = 0 To uBillboards.GetUpperBound(0)
                If uBillboards(X).oBBTex Is Nothing = False Then uBillboards(X).oBBTex.Dispose()
                uBillboards(X).oBBTex = Nothing
            Next X
        End If
        mbSkyboxReady = False
        If mvbSkybox Is Nothing = False Then mvbSkybox.Dispose()
        mvbSkybox = Nothing

        If moTerrain Is Nothing = False Then moTerrain.CleanResources()
    End Sub

    Public AxisAngle As Int32       'axis angle (yaw)
    Public RotateAngle As Int32   'rotation angle
    Private mlLastUpdate As Int32
    Private Shared moMapTex As Texture      'for system view 2
    Private Shared moSM_Mat As Material     'for system view 2

    Private Shared moPlanetGlowTex As Texture
    Private mvecGlowLoc As Vector3
    Private mbGlowLocSet As Boolean = False
    Private mfScale As Single
    Private mcolGlowColor As System.Drawing.Color
    Private moCloud As Mesh
    Private mfCloudRot As Single = 0.0F
	Private Shared mmatClouds As Material

	Private Shared moBumpPlanet As BumpedPlanet

    Public Sub SetLastUpdate(ByVal lVal As Int32)
        mlLastUpdate = lVal
    End Sub
#End Region

#Region "For Terrain Rendering"
    Private Sub CreateSkybox()
        Dim lIdx As Int32
        Dim fX As Single
        Dim fY As Single
        Dim fZ As Single
        Dim fD As Single

        Dim lColor As System.Drawing.Color
        Dim lCVal As Int32

        'now, generate our starfield...
		ReDim moSkyboxVerts(muSettings.StarfieldParticlesPlanet - 1)
		For lIdx = 0 To muSettings.StarfieldParticlesPlanet - 1
			'Now, we want to position our stars... anywhere along the sphere created by lFarPlane...
			fX = 0
			fZ = muSettings.FarClippingPlane - 1
			fY = 0

            fD = Int(Rnd() * 720) / 2
            RotatePoint(0, 0, fX, fZ, fD)
            fD = Int(Rnd() * 360) / 2
            RotatePoint(0, 0, fX, fY, fD)
            fD = Int(Rnd() * 360) / 2
            RotatePoint(0, 0, fZ, fY, fD)

            fY = Math.Abs(fY)

            lCVal = CInt(Int(Rnd() * 192) + 63)
            lColor = System.Drawing.Color.FromArgb(255, lCVal, lCVal, lCVal)

            moSkyboxVerts(lIdx) = New CustomVertex.PositionColored(fX, fY, fZ, lColor.ToArgb)
        Next lIdx

        'Now, create our buffer
        Device.IsUsingEventHandlers = False
        mvbSkybox = New VertexBuffer(moSkyboxVerts(0).GetType, muSettings.StarfieldParticlesPlanet, GFXEngine.moDevice, Usage.Points, CustomVertex.PositionColored.Format, Pool.Managed)
        Device.IsUsingEventHandlers = True
        mvbSkybox.SetData(moSkyboxVerts, 0, LockFlags.None)

        With moSkyboxMat
            .Ambient = System.Drawing.Color.White
            .Diffuse = .Ambient
            .Emissive = .Ambient
            .Specular = System.Drawing.Color.Black
        End With

        mbSkyboxReady = True
    End Sub
    'Private Sub RenderClouds(ByVal lExtent As Int32, ByVal oColor As System.Drawing.Color)
    '    Static fTex As Single

    '    If Atmosphere = 0 Then Return

    '    Dim X As Int32
    '    Dim fMult As Single = 2.0F   'Tweak this if clouds look unrealistic
    '    Dim oMat As Material
    '    Dim matWorld As Matrix

    '    fTex += 0.0001F
    '    If fTex > 1 Then fTex = 0

    '    If moCloudTex Is Nothing OrElse moCloudTex.Disposed = True Then
    '        moCloudTex = goResMgr.GetTexture("cloud1.dds", GFXResourceManager.eGetTextureType.ModelTexture)
    '    End If

    '    If mvbClouds Is Nothing OrElse mvbClouds.Disposed = True Then
    '        'ok, create it, we'll set it at the end
    '        Device.IsUsingEventHandlers = False
    '        mvbClouds = New VertexBuffer(New CustomVertex.PositionTextured().GetType, 20, moDevice, Usage.None, CustomVertex.PositionTextured.Format, Pool.Managed)
    '        mibClouds = New IndexBuffer(X.GetType, 30, moDevice, Usage.None, Pool.Managed)
    '        Device.IsUsingEventHandlers = True

    '        'now, set up our index buffer...
    '        Dim lIndices(29) As Int32
    '        lIndices(0) = 2 : lIndices(1) = 1 : lIndices(2) = 0 : lIndices(3) = 1 : lIndices(4) = 2 : lIndices(5) = 3
    '        lIndices(6) = 6 : lIndices(7) = 5 : lIndices(8) = 4 : lIndices(9) = 5 : lIndices(10) = 6 : lIndices(11) = 7
    '        lIndices(12) = 10 : lIndices(13) = 9 : lIndices(14) = 8 : lIndices(15) = 9 : lIndices(16) = 10 : lIndices(17) = 11
    '        lIndices(18) = 14 : lIndices(19) = 13 : lIndices(20) = 12 : lIndices(21) = 13 : lIndices(22) = 14 : lIndices(23) = 15
    '        lIndices(24) = 18 : lIndices(25) = 17 : lIndices(26) = 16 : lIndices(27) = 17 : lIndices(28) = 18 : lIndices(29) = 19

    '        mibClouds.SetData(lIndices, 0, LockFlags.None)

    '        Dim uVerts(19) As CustomVertex.PositionTextured
    '        Dim lLowHeight As Int32 = -10000    'tweak this for bottom of the box
    '        Dim lHiHeight As Int32 = 13000      'tweak this for top of the box
    '        Dim fTopExt As Single = lExtent * 0.9F

    '        'front 2 quads
    '        uVerts(0) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, -fTopExt, fTex * 5, (1 + fTex) * 5)
    '        uVerts(1) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, fTopExt, fTex * 5, fTex * 5)
    '        uVerts(2) = New CustomVertex.PositionTextured(0, lHiHeight, -fTopExt, (fTex + 1) * 5, (fTex + 1) * 5)
    '        uVerts(3) = New CustomVertex.PositionTextured(0, lHiHeight, fTopExt, (fTex + 1) * 5, fTex * 5)
    '        uVerts(4) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, -fTopExt, (fTex + 1) * 5, (fTex + 1) * 5)
    '        uVerts(5) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, fTopExt, (fTex + 1) * 5, fTex * 5)


    '        ''Top Quad
    '        'uVerts(0) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, -fTopExt, fTex * 5, (1 + fTex) * 5)
    '        'uVerts(1) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, fTopExt, fTex * 5, fTex * 5)
    '        'uVerts(2) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, -fTopExt, (fTex + 1) * 5, (fTex + 1) * 5)
    '        'uVerts(3) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, fTopExt, (fTex + 1) * 5, fTex * 5)

    '        ''Back Quad 
    '        'uVerts(4) = New CustomVertex.PositionTextured(-lExtent, lLowHeight, -lExtent, fTex * fMult, (1 + fTex) * fMult)
    '        'uVerts(5) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, -fTopExt, fTex * fMult, fTex * fMult)
    '        'uVerts(6) = New CustomVertex.PositionTextured(lExtent, lLowHeight, -lExtent, (fTex + 1) * fMult, (fTex + 1) * fMult)
    '        'uVerts(7) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, -fTopExt, (fTex + 1) * fMult, fTex * fMult)

    '        ''Right Quad
    '        'uVerts(8) = New CustomVertex.PositionTextured(lExtent, lLowHeight, -lExtent, (fTex + 1) * fMult, (fTex + 1) * fMult)
    '        'uVerts(9) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, -fTopExt, (fTex + 1) * fMult, fTex * fMult)
    '        'uVerts(10) = New CustomVertex.PositionTextured(lExtent, lLowHeight, lExtent, fTex * fMult, (1 + fTex) * fMult)
    '        'uVerts(11) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, fTopExt, fTex * fMult, fTex * fMult)

    '        ''Front Quad
    '        'uVerts(12) = New CustomVertex.PositionTextured(lExtent, lLowHeight, lExtent, (fTex + 1) * fMult, (fTex - 1) * fMult)
    '        'uVerts(13) = New CustomVertex.PositionTextured(fTopExt, lHiHeight, fTopExt, (fTex + 1) * fMult, fTex * fMult)
    '        'uVerts(14) = New CustomVertex.PositionTextured(-lExtent, lLowHeight, lExtent, fTex * fMult, (fTex - 1) * fMult)
    '        'uVerts(15) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, fTopExt, fTex * fMult, fTex * fMult)

    '        ''Left Quad
    '        'uVerts(16) = New CustomVertex.PositionTextured(-lExtent, lLowHeight, lExtent, fTex * fMult, (fTex - 1) * fMult)
    '        'uVerts(17) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, fTopExt, fTex * fMult, fTex * fMult)
    '        'uVerts(18) = New CustomVertex.PositionTextured(-lExtent, lLowHeight, -lExtent, (fTex + 1) * fMult, (fTex - 1) * fMult)
    '        'uVerts(19) = New CustomVertex.PositionTextured(-fTopExt, lHiHeight, -fTopExt, (fTex + 1) * fMult, fTex * fMult)

    '        mvbClouds.SetData(uVerts, 0, LockFlags.Discard)
    '    End If

    '    'First do our cloud animation...
    '    'Ok, lock our vertex buffer...
    '    Dim ranks(0) As Integer
    '    ranks(0) = 20
    '    Dim arr As System.Array = mvbClouds.Lock(0, (New CustomVertex.PositionTextured()).GetType(), LockFlags.None, ranks)
    '    For X = 0 To arr.Length - 1
    '        Dim pn As Direct3D.CustomVertex.PositionTextured = CType(arr.GetValue(X), CustomVertex.PositionTextured)
    '        Select Case X
    '            Case 0 : pn.Tu = fTex * 5 : pn.Tv = (1 + fTex) * 5
    '            Case 1 : pn.Tu = fTex * 5 : pn.Tv = fTex * 5
    '            Case 2 : pn.Tu = (fTex + 1) * 5 : pn.Tv = (fTex + 1) * 5
    '            Case 3 : pn.Tu = (fTex + 1) * 5 : pn.Tv = fTex * 5
    '            Case 5, 11, 15, 17
    '                pn.Tu = fTex * fMult : pn.Tv = fTex * fMult
    '            Case 7, 9, 13, 19
    '                pn.Tu = (fTex + 1) * fMult : pn.Tv = fTex * fMult
    '            Case 4, 10
    '                pn.Tu = fTex * fMult : pn.Tv = (fTex + 1) * fMult
    '            Case 6, 8
    '                pn.Tu = (fTex + 1) * fMult : pn.Tv = (fTex + 1) * fMult
    '            Case Else
    '                If X = 12 OrElse X = 18 Then
    '                    pn.Tu = (fTex + 1) * fMult
    '                Else
    '                    pn.Tu = fTex * fMult
    '                End If
    '                pn.Tv = (fTex - 1) * fMult
    '        End Select
    '        arr.SetValue(pn, X)
    '    Next X
    '    mvbClouds.Unlock()
    '    arr = Nothing

    '    'Now, render our clouds
    '    With oMat
    '        Dim lR As Int32
    '        Dim lG As Int32
    '        Dim lB As Int32

    '        lR = oColor.R + 20
    '        lG = oColor.G + 20
    '        lB = oColor.B + 20

    '        If lR > 255 Then lR = 255
    '        If lG > 255 Then lG = 255
    '        If lB > 255 Then lB = 255
    '        oColor = System.Drawing.Color.FromArgb(oColor.A, lR, lG, lB)

    '        .Ambient = oColor

    '        .Diffuse = oColor
    '        .Emissive = System.Drawing.Color.Black
    '        .Specular = .Emissive
    '    End With

    '    With moDevice
    '        matWorld = Matrix.Identity
    '        matWorld.Translate(goCamera.mlCameraAtX, 0, goCamera.mlCameraAtZ)
    '        .Transform.World = matWorld

    '        .RenderState.SourceBlend = Blend.SourceAlpha
    '        .RenderState.DestinationBlend = Blend.InvSourceAlpha
    '        .RenderState.AlphaBlendEnable = True
    '        .RenderState.ZBufferWriteEnable = False

    '        Try
    '            'set the information the device needs...
    '            .Material = oMat
    '            .SetTexture(0, moCloudTex)
    '            .VertexFormat = CustomVertex.PositionTextured.Format
    '            .Indices = mibClouds
    '            .SetStreamSource(0, mvbClouds, 0)

    '            'Render our primitives
    '            .DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, 20, 0, 10)
    '        Catch
    '            'do nothing, for now
    '        End Try

    '        'reset state data
    '        .RenderState.SourceBlend = Blend.SourceAlpha        'Blend.One
    '        .RenderState.DestinationBlend = Blend.InvSourceAlpha    'Blend.Zero
    '        .RenderState.ZBufferWriteEnable = True
    '    End With
    'End Sub
    Private Sub RenderClouds(ByVal lExtent As Int32, ByVal oColor As System.Drawing.Color)
        Static fTex As Single

        If Atmosphere = 0 Then Return

        Dim X As Int32
        Dim fMult As Single = 2.0F   'Tweak this if clouds look unrealistic
        Dim oMat As Material
        Dim matWorld As Matrix
        Dim moDevice As Device = GFXEngine.moDevice

        fTex += 0.0001F
        If fTex > 1 Then fTex = 0

        If moCloudTex Is Nothing OrElse moCloudTex.Disposed = True Then
            moCloudTex = goResMgr.GetTexture("cloud1.dds", GFXResourceManager.eGetTextureType.ModelTexture)
        End If

        If mvbClouds Is Nothing OrElse mvbClouds.Disposed = True Then
            'ok, create it, we'll set it at the end
            Device.IsUsingEventHandlers = False
            mvbClouds = New VertexBuffer(New CustomVertex.PositionColoredTextured().GetType, 9, moDevice, Usage.None, CustomVertex.PositionColoredTextured.Format, Pool.Managed)
            mibClouds = New IndexBuffer(X.GetType, 24, moDevice, Usage.None, Pool.Managed)
            Device.IsUsingEventHandlers = True

            'now, set up our index buffer...
            Dim lIndices(23) As Int32
            lIndices(0) = 1 : lIndices(1) = 0 : lIndices(2) = 3 : lIndices(3) = 0 : lIndices(4) = 2 : lIndices(5) = 3
            lIndices(6) = 3 : lIndices(7) = 2 : lIndices(8) = 5 : lIndices(9) = 2 : lIndices(10) = 4 : lIndices(11) = 5
            lIndices(12) = 6 : lIndices(13) = 1 : lIndices(14) = 7 : lIndices(15) = 7 : lIndices(16) = 1 : lIndices(17) = 3
            lIndices(18) = 7 : lIndices(19) = 3 : lIndices(20) = 8 : lIndices(21) = 8 : lIndices(22) = 3 : lIndices(23) = 5

            mibClouds.SetData(lIndices, 0, LockFlags.None)

            Dim uVerts(8) As CustomVertex.PositionColoredTextured
            Dim lLowHeight As Int32 = 5000    'tweak this for bottom of the box
            Dim lHiHeight As Int32 = 13000      'tweak this for top of the box
            Dim fTopExt As Single = lExtent * 0.9F

            Dim lCenterClr As Int32 = System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb
            Dim lFarClr As Int32 = System.Drawing.Color.FromArgb(0, 255, 255, 255).ToArgb

            'uVerts(0) = New CustomVertex.PositionColoredTextured(-fTopExt, lHiHeight, -fTopExt, lFarClr, fTex * 5, (1 + fTex) * 5)
            'uVerts(1) = New CustomVertex.PositionColoredTextured(-fTopExt, lHiHeight, 0, lFarClr, fTex * 5, fTex * 5)
            'uVerts(2) = New CustomVertex.PositionColoredTextured(0, lHiHeight, -fTopExt, lFarClr, (fTex + 1) * 5, (fTex + 1) * 5)
            'uVerts(3) = New CustomVertex.PositionColoredTextured(0, lHiHeight, 0, lCenterClr, (fTex + 1) * 5, fTex * 5)
            'uVerts(4) = New CustomVertex.PositionColoredTextured(fTopExt, lHiHeight, -fTopExt, lFarClr, (fTex + 1) * 5, (fTex + 1) * 5)
            'uVerts(5) = New CustomVertex.PositionColoredTextured(fTopExt, lHiHeight, 0, lFarClr, (fTex + 1) * 5, fTex * 5)
            'uVerts(6) = New CustomVertex.PositionColoredTextured(-fTopExt, lHiHeight, fTopExt, lFarClr, fTex * 5, fTex * 5)
            'uVerts(7) = New CustomVertex.PositionColoredTextured(0, lHiHeight, fTopExt, lFarClr, (fTex + 1) * 5, fTex * 5)
            'uVerts(8) = New CustomVertex.PositionColoredTextured(fTopExt, lHiHeight, fTopExt, lFarClr, (fTex + 1) * 5, fTex * 5)

            uVerts(0) = New CustomVertex.PositionColoredTextured(-fTopExt, lLowHeight, -fTopExt, lFarClr, 0, 0)
            uVerts(1) = New CustomVertex.PositionColoredTextured(-fTopExt, lLowHeight, 0, lFarClr, 0, 0.5F)
            uVerts(2) = New CustomVertex.PositionColoredTextured(0, lLowHeight, -fTopExt, lFarClr, 0.5F, 0)
            uVerts(3) = New CustomVertex.PositionColoredTextured(0, lHiHeight, 0, lCenterClr, 0.5F, 0.5F)
            uVerts(4) = New CustomVertex.PositionColoredTextured(fTopExt, lLowHeight, -fTopExt, lFarClr, 1, 0)
            uVerts(5) = New CustomVertex.PositionColoredTextured(fTopExt, lLowHeight, 0, lFarClr, 1, 0.5F)
            uVerts(6) = New CustomVertex.PositionColoredTextured(-fTopExt, lLowHeight, fTopExt, lFarClr, 0, 1)
            uVerts(7) = New CustomVertex.PositionColoredTextured(0, lLowHeight, fTopExt, lFarClr, 0.5F, 1)
            uVerts(8) = New CustomVertex.PositionColoredTextured(fTopExt, lLowHeight, fTopExt, lFarClr, 1, 1)


            mvbClouds.SetData(uVerts, 0, LockFlags.Discard)
        End If

        'First do our cloud animation...
        'Ok, lock our vertex buffer...
        Dim ranks(0) As Integer
        ranks(0) = 20
        Dim arr As System.Array = Nothing
        Dim bLocked As Boolean = False

        Try
            arr = mvbClouds.Lock(0, (New CustomVertex.PositionColoredTextured()).GetType(), LockFlags.None, ranks)
            bLocked = True
            For X = 0 To arr.Length - 1
                Dim pn As Direct3D.CustomVertex.PositionColoredTextured = CType(arr.GetValue(X), CustomVertex.PositionColoredTextured)
                '        Select Case X
                '            Case 0 : pn.Tu = fTex * 5 : pn.Tv = (1 + fTex) * 5
                '            Case 1 : pn.Tu = fTex * 5 : pn.Tv = fTex * 5
                '            Case 2 : pn.Tu = (fTex + 1) * 5 : pn.Tv = (fTex + 1) * 5
                '            Case 3 : pn.Tu = (fTex + 1) * 5 : pn.Tv = fTex * 5
                '            Case 5, 11, 15, 17
                '                pn.Tu = fTex * fMult : pn.Tv = fTex * fMult
                '            Case 7, 9, 13, 19
                '                pn.Tu = (fTex + 1) * fMult : pn.Tv = fTex * fMult
                '            Case 4, 10
                '                pn.Tu = fTex * fMult : pn.Tv = (fTex + 1) * fMult
                '            Case 6, 8
                '                pn.Tu = (fTex + 1) * fMult : pn.Tv = (fTex + 1) * fMult
                '            Case Else
                '                If X = 12 OrElse X = 18 Then
                '                    pn.Tu = (fTex + 1) * fMult
                '                Else
                '                    pn.Tu = fTex * fMult
                '                End If
                '                pn.Tv = (fTex - 1) * fMult
                '        End Select
                Select Case X
                    Case 0
                        pn.Tu = fTex * 5 : pn.Tv = fTex * 5
                    Case 1
                        pn.Tu = fTex * 5 : pn.Tv = (1 + fTex) * 5
                    Case 2
                        pn.Tu = (1 + fTex) * 5 : pn.Tv = fTex * 5
                    Case 3
                        pn.Tu = (1 + fTex) * 5 : pn.Tv = (1 + fTex) * 5
                    Case 4
                        pn.Tu = (2 + fTex) * 5 : pn.Tv = fTex * 5
                    Case 5
                        pn.Tu = (2 + fTex) * 5 : pn.Tv = (1 + fTex) * 5
                    Case 6
                        pn.Tu = fTex * 5 : pn.Tv = (2 + fTex) * 5
                    Case 7
                        pn.Tu = (1 + fTex) * 5 : pn.Tv = (2 + fTex) * 5
                    Case 8
                        pn.Tu = (2 + fTex) * 5 : pn.Tv = (2 + fTex) * 5
                End Select
                'pn.Tu = (pn.Tu + 1) * fTexMult
                'pn.Tv = (pn.Tv + 1) * fTexMult
                arr.SetValue(pn, X)
            Next X
        Catch
        Finally
            If mvbClouds Is Nothing = False Then
                Try
                    If bLocked = True Then mvbClouds.Unlock()
                Catch
                    If mvbClouds Is Nothing = False Then mvbClouds.Dispose()
                    mvbClouds = Nothing
                End Try
            End If
        End Try
        arr = Nothing

        'Now, render our clouds
        With oMat
            Dim lR As Int32
            Dim lG As Int32
            Dim lB As Int32

            lR = oColor.R + 20
            lG = oColor.G + 20
            lB = oColor.B + 20

            If lR > 255 Then lR = 255
            If lG > 255 Then lG = 255
            If lB > 255 Then lB = 255
            oColor = System.Drawing.Color.FromArgb(oColor.A, lR, lG, lB)

            .Ambient = oColor

            .Diffuse = oColor
            .Emissive = System.Drawing.Color.Black
            .Specular = .Emissive
        End With

        With moDevice
            matWorld = Matrix.Identity
            matWorld.Translate(goCamera.mlCameraAtX, 0, goCamera.mlCameraAtZ)
            .Transform.World = matWorld

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)

            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.InvSourceAlpha
            .RenderState.AlphaBlendEnable = True
            .RenderState.ZBufferWriteEnable = False

            Try
                'set the information the device needs...
                .Material = oMat
                .SetTexture(0, moCloudTex)
                .VertexFormat = CustomVertex.PositionColoredTextured.Format
                .Indices = mibClouds
                .SetStreamSource(0, mvbClouds, 0)

                'Render our primitives
                .DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, 9, 0, 8)
            Catch
                'do nothing, for now
            End Try

            'reset state data
            .RenderState.SourceBlend = Blend.SourceAlpha        'Blend.One
            .RenderState.DestinationBlend = Blend.InvSourceAlpha    'Blend.Zero
            .RenderState.ZBufferWriteEnable = True

            .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
            .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
            .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
        End With
    End Sub

    Private Shared mbSkyboxReady As Boolean = False
    Private Shared moSkyboxVerts() As CustomVertex.PositionColored
    Private Shared mvbSkybox As VertexBuffer
    Private Shared moSkyboxMat As Material

    Private Shared mfSphereU() As Single
    Private Shared mfSphereV() As Single
    Private Shared moColorSphere As Mesh
    Private Shared myWestR As Byte
    Private Shared myWestG As Byte
    Private Shared myWestB As Byte
    Private Shared myEastR As Byte
    Private Shared myEastG As Byte
    Private Shared myEastB As Byte

    Private Shared moCloudTex As Texture
    Private Shared mvbClouds As VertexBuffer
    Private Shared mibClouds As IndexBuffer

    Private Shared moLensFlare As Texture

    Private mbCullBoxSet As Boolean = False
    Private muCullBox As CullBox
    Public ReadOnly Property uCullBox() As CullBox
        Get
            If mbCullBoxSet = False Then
                muCullBox = CullBox.GetCullBox(Me.LocX, Me.LocY, Me.LocZ, -Me.PlanetRadius, -Me.PlanetRadius, -Me.PlanetRadius, Me.PlanetRadius, Me.PlanetRadius, Me.PlanetRadius)
                mbCullBoxSet = True
            End If
            Return muCullBox
        End Get
    End Property
#End Region

    Public Property PlanetRadius() As Int16
        Get
            Return miPlanetRadius
        End Get
        Set(ByVal Value As Int16)
            miPlanetRadius = Value

            moMesh = goResMgr.CreateTexturedSphere(miPlanetRadius, 32, 32, 0)
            mbMeshTBNReady = False

            moCloud = goResMgr.CreateTexturedSphere(miPlanetRadius * 1.01F, 32, 32, 0)
        End Set
    End Property
    Public ReadOnly Property PlanetSizeID() As Byte
        Get
            Return myPlanetSizeID
        End Get
    End Property

    Public Sub New(ByVal lID As Int32, ByVal ySizeID As Byte, ByVal yMapTypeID As Byte)

        For X As Int32 = 0 To 5
            uBillboards(X).lGuildID = -1
        Next

        ObjectID = lID
        myPlanetSizeID = ySizeID

        Me.MapTypeID = yMapTypeID

        moTerrain = New TerrainClass(lID)
        If myPlanetSizeID = 0 Then moTerrain.ml_Y_Mult = 9.0F
        moTerrain.MapType = yMapTypeID

        Select Case myPlanetSizeID
            Case 0 : moTerrain.CellSpacing = gl_TINY_PLANET_CELL_SPACING
            Case 1 : moTerrain.CellSpacing = gl_SMALL_PLANET_CELL_SPACING
            Case 2 : moTerrain.CellSpacing = gl_MEDIUM_PLANET_CELL_SPACING
            Case 3 : moTerrain.CellSpacing = gl_LARGE_PLANET_CELL_SPACING
            Case 4 : moTerrain.CellSpacing = gl_HUGE_PLANET_CELL_SPACING
        End Select

        Select Case yMapTypeID
            Case PlanetType.eAcidic
                moEnvirColor = System.Drawing.Color.FromArgb(255, 64, 115, 64)
                mcolGlowColor = System.Drawing.Color.FromArgb(255, 32, 192, 64)
            Case PlanetType.eAdaptable
                moEnvirColor = System.Drawing.Color.FromArgb(255, 128, 0, 0)
                mcolGlowColor = System.Drawing.Color.FromArgb(128, 192, 192, 255)
            Case PlanetType.eBarren
                moEnvirColor = System.Drawing.Color.Black
                mcolGlowColor = System.Drawing.Color.FromArgb(0, 0, 0, 0)
            Case PlanetType.eDesert
                moEnvirColor = System.Drawing.Color.FromArgb(255, 255, 255, 128)
                mcolGlowColor = System.Drawing.Color.FromArgb(255, 192, 192, 128)
            Case PlanetType.eGeoPlastic
                moEnvirColor = System.Drawing.Color.FromArgb(255, 128, 64, 0)
                mcolGlowColor = System.Drawing.Color.FromArgb(255, 255, 64, 32)
            Case PlanetType.eTerran, PlanetType.eTundra, PlanetType.eWaterWorld
                moEnvirColor = System.Drawing.Color.FromArgb(255, 0, 128, 255)
                mcolGlowColor = System.Drawing.Color.FromArgb(255, 128, 128, 192)
        End Select

        If mbSkyboxReady = False Then
            CreateSkybox()
        End If

        EnsureSpritesValid(GFXEngine.moDevice)

        If moPlanetGlowTex Is Nothing OrElse moPlanetGlowTex.Disposed = True Then
            moPlanetGlowTex = goResMgr.GetTexture("PlanetGlow.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "SpacePlanet.pak")
            With mmatClouds
                .Ambient = Color.Black
                .Diffuse = Color.White
                .Emissive = Color.Black
            End With
        End If
    End Sub

    'MSC - 08/18/08 - remarked to see if nagging crash bug goes away
    '   Private Shared Sub SpriteDispose(ByVal sender As Object, ByVal e As EventArgs)
    '       moSprite = Nothing
    '   End Sub

    '   Private Shared Sub SpriteLost(ByVal sender As Object, ByVal e As EventArgs)
    '       If moSprite Is Nothing = False Then moSprite.Dispose()
    '       moSprite = Nothing
    'End Sub

    Private mbInMakeMeshTBNReady As Boolean = False
    Private Function MakeMeshTBNReady(ByRef oMesh As Mesh) As Mesh
        If mbInMakeMeshTBNReady = True Then Return Nothing
        mbInMakeMeshTBNReady = True

        Dim oTmpMesh As Mesh = Nothing

        Try
            Dim elems(5) As VertexElement
            elems(0) = New VertexElement(0, 0, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Position, 0)
            elems(1) = New VertexElement(0, 12, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Normal, 0)
            elems(2) = New VertexElement(0, 24, DeclarationType.Float2, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 0)
            elems(3) = New VertexElement(0, 32, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Tangent, 0)
            elems(4) = New VertexElement(0, 44, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.BiNormal, 0)
            elems(5) = VertexElement.VertexDeclarationEnd

            oTmpMesh = oMesh.Clone(MeshFlags.Managed, elems, GFXEngine.moDevice)
            oTmpMesh.ComputeTangent(0, 0, 0, 0)
        Catch 'catch all exceptions
            'just recreate the Mesh
            Try
                If moMesh Is Nothing = False Then moMesh.Dispose()
            Catch
            End Try
            moMesh = Nothing
            moMesh = goResMgr.CreateTexturedSphere(miPlanetRadius, 32, 32, 0)

            oTmpMesh = MakeMeshTBNReady(moMesh)
        End Try

        mbInMakeMeshTBNReady = False
        Return oTmpMesh
    End Function

    Public Shared b_LOCK_TIME_OF_DAY As Boolean = False
    Private bSaved As Boolean = False
    Private bSaved2 As Boolean = False
    Public Sub Render(ByVal lEnvirType As Int32)

        Dim moDevice As Device = GFXEngine.moDevice

        If mbGlowLocSet = False Then
            mfScale = (Me.PlanetRadius / 0.87F) / 256.0F
            mvecGlowLoc = New Vector3(LocX / mfScale, LocY / mfScale, LocZ / mfScale)
            mbGlowLocSet = True
        End If

        If mlLastUpdate = 0 Then
            mlLastUpdate = timeGetTime
        Else
            'Ok, now... the primary server determines rotation using the following:
            '   Return Math.Floor((timeGetTime / 30) / RotationDelay) Mod 3600
            'RotationDelay = 10
            If b_LOCK_TIME_OF_DAY = False AndAlso (timeGetTime - mlLastUpdate) > RotationDelay Then
                Dim fPasses As Single = CSng(timeGetTime - mlLastUpdate) / RotationDelay
                RotateAngle += CInt(Int(fPasses))
                If RotateAngle > 3600 Then RotateAngle = RotateAngle Mod 3600
                mlLastUpdate = CInt(timeGetTime - ((fPasses - Int(fPasses)) * RotationDelay))
            End If
            'Dim fPasses As Single = CSng(timeGetTime - mlLastUpdate) / 10
            'RotateAngle += CInt(Int(fPasses))
            'If RotateAngle > 3600 Then RotateAngle = RotateAngle Mod 3600
            'mlLastUpdate = CInt(timeGetTime - ((fPasses - Int(fPasses)) * 10))

        End If
        If lEnvirType = CurrentView.ePlanetMapView AndAlso goCurrentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso NewTutorialManager.TutorialOn = False Then
            'Show our Strategic Map Control window...
            If goUILib Is Nothing = False Then
                Dim oFrm As UIWindow = goUILib.GetWindow("frmStrategicMapControl")
                If oFrm Is Nothing Then
                    oFrm = New frmStrategicMapControl(goUILib)
                    oFrm.Visible = True
                ElseIf oFrm.Visible = False Then
                    oFrm.Visible = True
                End If
                oFrm = Nothing
            End If
        ElseIf lEnvirType = CurrentView.eSystemMapView2 AndAlso goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso NewTutorialManager.TutorialOn = False Then
            If goUILib Is Nothing = False Then
                Dim oFrm As UIWindow = goUILib.GetWindow("frmStrategicMapControl")
                If oFrm Is Nothing Then
                    oFrm = New frmStrategicMapControl(goUILib)
                    oFrm.Visible = True
                ElseIf oFrm.Visible = False Then
                    oFrm.Visible = True
                End If
                oFrm = Nothing
            End If
        End If
        If lEnvirType = CurrentView.eSystemView OrElse lEnvirType = CurrentView.eSystemMapView2 Then
            'Ok, render the planet as a globe on the system view (normal)

            Dim dAngle As Single
            Dim matYaw As Matrix
            Dim matAngle As Matrix
            Dim matWorld As Matrix
            Dim matTemp As Matrix
            Dim matScale As Matrix
            Dim fScale As Single

            Dim bCullObject As Boolean = False

            Dim bInSpecialFXRange As Boolean = Distance(goCamera.mlCameraX, goCamera.mlCameraZ, Me.LocX, Me.LocZ) < muSettings.EntityClipPlane

            'Check for system view
            If lEnvirType = CurrentView.eSystemView Then
                bCullObject = goCamera.CullObject_NoMaxDistance(uCullBox)
                If bCullObject = False Then
                    If Me.ParentSystem Is Nothing = False AndAlso Me.ParentSystem.ObjectID = 36 Then
                        If glCurrentCycle - lLastBBRequest > 900 Then
                            lLastBBRequest = glCurrentCycle
                            Dim yMsg(7) As Byte
                            Dim lPos As Int32 = 0
                            System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildBillboards).CopyTo(yMsg, lPos) : lPos += 2
                            Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                            goUILib.SendMsgToPrimary(yMsg)
                        End If
                    End If

                    With moDevice
                        .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                        .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Diffuse)
                        .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Modulate)
                        .RenderState.SourceBlend = Blend.SourceAlpha
                        .RenderState.DestinationBlend = Blend.One
                        .RenderState.AlphaBlendEnable = True
                        .RenderState.ZBufferWriteEnable = False
                    End With
                    If moPlanetGlowTex Is Nothing OrElse moPlanetGlowTex.Disposed = True Then moPlanetGlowTex = goResMgr.GetTexture("PlanetGlow.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "SpacePlanet.pak")

                    If moPGlowSpr Is Nothing Then
                        Device.IsUsingEventHandlers = False
                        moPGlowSpr = New Sprite(moDevice)
                        'AddHandler moPGlowSpr.Disposing, AddressOf SpriteDispose
                        'AddHandler moPGlowSpr.Lost, AddressOf SpriteLost
                        'AddHandler moPGlowSpr.Reset, AddressOf SpriteLost
                        Device.IsUsingEventHandlers = True
                    End If

                    'Draw the planet glow...
                    With moPGlowSpr
                        matTemp = Matrix.Scaling(mfScale, mfScale, mfScale)
                        matWorld = Matrix.Identity
                        matWorld.Multiply(matTemp)
                        matTemp = Matrix.Identity
                        matTemp.Translate(LocX, LocY, LocZ)
                        matWorld.Multiply(matTemp)

                        moDevice.Transform.World = matWorld
                        .SetWorldViewLH(moDevice.Transform.World, moDevice.Transform.View)
                        .Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)
                        If moPlanetGlowTex Is Nothing = False Then .Draw(moPlanetGlowTex, System.Drawing.Rectangle.Empty, New Vector3(256, 256, 0), New Vector3(0, 0, 0), mcolGlowColor)
                        .End()
                    End With

                    With moDevice
                        .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
                        .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
                        .RenderState.AlphaBlendEnable = True
                        .RenderState.ZBufferWriteEnable = True
                        .SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                        .SetTextureStageState(0, TextureStageStates.AlphaArgument2, TextureArgument.Current)
                        .SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                    End With
                End If
            End If

            If bCullObject = False Then
                If moTexture Is Nothing AndAlso lEnvirType = CurrentView.eSystemView Then

                    If goCamera.CullObject(uCullBox) = False Then
                        'goUILib.AddNotification("Planet Texture created: " & Me.PlanetName, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If muSettings.CachePlanetTextures = False OrElse LoadCachedPlanet() = False Then
                            moTerrain.CreateTerrainPlanetTexture(moTexture)
                        End If
                        If Me.MapTypeID = PlanetType.eGeoPlastic Then moLightMap = moTerrain.CreateTerrainLavaLightMap()
                        With moMaterial
                            .Diffuse = System.Drawing.Color.White
                            .Ambient = System.Drawing.Color.FromArgb(255, 18, 18, 18)
                            .Emissive = System.Drawing.Color.Black
                            .Specular = Color.FromArgb(128, 128, 128, 128)
                            .SpecularSharpness = 30
                        End With
                    Else
                        With moMaterial
                            Select Case MapTypeID
                                Case PlanetType.eAcidic
                                    .Diffuse = System.Drawing.Color.FromArgb(255, 64, 192, 64)
                                Case PlanetType.eAdaptable
                                    .Diffuse = System.Drawing.Color.FromArgb(255, 255, 64, 0)
                                Case PlanetType.eBarren
                                    .Diffuse = System.Drawing.Color.FromArgb(255, 192, 192, 192)
                                Case PlanetType.eDesert
                                    .Diffuse = System.Drawing.Color.FromArgb(255, 255, 192, 128)
                                Case PlanetType.eGeoPlastic
                                    .Diffuse = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                                Case PlanetType.eTerran
                                    .Diffuse = System.Drawing.Color.FromArgb(255, 128, 128, 255)
                                Case PlanetType.eWaterWorld
                                    .Diffuse = System.Drawing.Color.FromArgb(255, 0, 128, 255)
                                Case Else       'tundra and any others we miss
                                    .Diffuse = System.Drawing.Color.White
                            End Select
                            .Ambient = System.Drawing.Color.FromArgb(255, 18, 18, 18)
                            .Emissive = System.Drawing.Color.Black
                            .Specular = Color.FromArgb(128, 128, 128, 128)
                            .SpecularSharpness = 30
                        End With
                    End If
                End If

                If muSettings.CachePlanetTextures = True AndAlso moTexture Is Nothing = False AndAlso bSaved = False Then
                    SavePlanetTexture()
                End If
                matWorld = Matrix.Identity

                'Rotate Yaw (Axis)
                dAngle = ((AxisAngle) / 10.0F) * CSng(Math.PI / 180.0F)
                'matYaw = Matrix.Identity
                'matYaw = Matrix.Multiply(matYaw, Matrix.RotationX(CSng(dAngle)))
                'matYaw.RotateX(CSng(dAngle))

                'matWorld.Multiply(matYaw)

                'Rotate Angle (rotateangle)
                'matAngle = Matrix.Identity
                Dim dAngle2 As Single = (RotateAngle / 10.0F) * CSng(Math.PI / 180.0F)
                'matAngle.RotateY(CSng(dAngle))
                'matAngle = Matrix.Multiply(matAngle, Matrix.RotationY(CSng(dAngle)))

                'matWorld.Multiply(matAngle)
                matWorld = Matrix.Multiply(matWorld, Matrix.RotationYawPitchRoll(0, dAngle, dAngle2))

                'Create our world
                'matWorld = Matrix.Multiply(matAngle, matYaw)
                matYaw = Nothing
                matAngle = Nothing
                If lEnvirType = CurrentView.eSystemMapView2 Then
                    fScale = 1 / 30
                    matScale = Matrix.Identity
                    matScale.Scale(fScale, fScale, fScale)
                    matWorld.Multiply(matScale)
                    matTemp = Matrix.Identity
                    matTemp.Translate(CSng(LocX / 30), CSng(LocY / 30), CSng(LocZ / 30))
                Else
                    matTemp = Matrix.Identity
                    matTemp.Translate(LocX, LocY, LocZ)
                End If

                matWorld.Multiply(matTemp)      'set our loc
                matTemp = Nothing

                'set the world and render
                moDevice.Transform.World = matWorld

                'Make sure our planet exists...
                moDevice.RenderState.Wrap0 = WrapCoordinates.Zero
                moDevice.RenderState.Wrap1 = WrapCoordinates.Zero
                moDevice.RenderState.Wrap2 = WrapCoordinates.Zero

                If lEnvirType = CurrentView.eSystemMapView2 Then
                    If moMapTex Is Nothing OrElse moMapTex.Disposed = True Then
                        moMapTex = goResMgr.GetTexture("Unplotted.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "MinCacheTex.pak")
                        'moMapTex = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Unplotted.bmp") ', 0, 0, 1, Usage.None, Format.Unknown, Pool.Default, Filter.Triangle, Filter.Triangle, 0)
                        With moSM_Mat
                            .Ambient = System.Drawing.Color.White
                            .Diffuse = System.Drawing.Color.White
                            .Emissive = System.Drawing.Color.White
                            .Specular = System.Drawing.Color.Black
                        End With
                    End If
                    moDevice.SetTexture(0, moMapTex)
                    moDevice.Material = moSM_Mat
                Else
                    moDevice.Material = moMaterial
                    moDevice.SetTexture(0, moTexture)

                    If muSettings.BumpMapPlanetModel = True AndAlso bInSpecialFXRange = True Then

                        If mbMeshTBNReady = False Then
                            mbMeshTBNReady = True
                            moMesh = MakeMeshTBNReady(moMesh)
                            If moMesh Is Nothing Then
                                mbMeshTBNReady = False
                                Return
                            End If
                        End If

                        If muSettings.CachePlanetTextures = False Then
                            If moNormalMap Is Nothing Then moNormalMap = moTerrain.GeneratePlanetModelNormalMap()
                        Else
                            If moNormalMap Is Nothing Then
                                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                                If sPath.EndsWith("\") = False Then sPath &= "\"
                                sPath &= "cache\planetBump_" & Me.ObjectID & ".bmp"

                                If Exists(sPath) Then
                                    moNormalMap = TextureLoader.FromFile(moDevice, sPath, 0, 0, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Box, 0)
                                    'Debug.Print("Loaded Cached Bump " & Me.ObjectID)
                                End If

                                If moNormalMap Is Nothing Then
                                    moNormalMap = moTerrain.GeneratePlanetModelNormalMap()
                                End If
                            End If
                            If moNormalMap Is Nothing = False And bSaved2 = False Then
                                Try
                                    Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                                    If sPath.EndsWith("\") = False Then sPath &= "\"
                                    sPath &= "cache"
                                    If Exists(sPath) = False Then
                                        MkDir(sPath)
                                    End If
                                    sPath &= "\planetBump_" & Me.ObjectID

                                    If Not Exists(sPath & ".bmp") Then
                                        'Debug.Print("Saving Bump " & Me.ObjectID)
                                        TextureLoader.Save(sPath & ".bmp", ImageFileFormat.Bmp, moNormalMap)
                                    End If
                                    bSaved2 = True
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                        Dim vecTemp As Vector3 = New Vector3(Me.LocX, Me.LocY, Me.LocZ)
                        If moBumpPlanet Is Nothing Then moBumpPlanet = New BumpedPlanet()

                        moBumpPlanet.PrepareToRender(vecTemp, moDevice.Lights(0).Diffuse, moDevice.Lights(0).Ambient, moDevice.Lights(0).Specular)
                        moBumpPlanet.SetTextures(moTexture, moNormalMap, moLightMap)
                        moBumpPlanet.UpdateMatrices()
                    End If
                End If

                If MapTypeID = PlanetType.eGeoPlastic AndAlso mbShowLightMap = True AndAlso muSettings.BumpMapPlanetModel = False Then
                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                    moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                    moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.SelectArg1)
                    'moDevice.SetTextureStageState(0, TextureStageStates.TextureCoordinateIndex, 0)

                    moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Modulate)
                    moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.Current)
                    moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
                    moDevice.SetTextureStageState(1, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
                    moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)

                    moDevice.SetSamplerState(2, SamplerStageStates.MinFilter, TextureFilter.Linear)
                    moDevice.SetSamplerState(2, SamplerStageStates.MagFilter, TextureFilter.Linear)
                    moDevice.SetSamplerState(2, SamplerStageStates.MipFilter, TextureFilter.Linear)
                    moDevice.SetTextureStageState(2, TextureStageStates.AlphaArgument1, TextureArgument.Diffuse)
                    moDevice.SetTextureStageState(2, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                    moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                    moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument2, TextureArgument.Current)
                    moDevice.SetTextureStageState(2, TextureStageStates.ColorOperation, TextureOperation.Add)
                    'moDevice.SetTextureStageState(2, TextureStageStates.TextureCoordinateIndex, 0)
                    'If moLightMap Is Nothing OrElse moLightMap.Disposed = True Then moLightMap = moTerrain.CreateTerrainLavaLightMap()
                    moDevice.SetTexture(2, moLightMap)
                End If
                moDevice.SetTextureStageState(0, TextureStageStates.TextureCoordinateIndex, 0)
                moDevice.SetTextureStageState(2, TextureStageStates.TextureCoordinateIndex, 0)

                If moMesh Is Nothing OrElse moMesh.Disposed = True Then moMesh = goResMgr.CreateTexturedSphere(miPlanetRadius, 32, 32, 0)
                moDevice.RenderState.AlphaBlendEnable = False
                moMesh.DrawSubset(0)
                moDevice.RenderState.AlphaBlendEnable = True

                If MapTypeID = PlanetType.eGeoPlastic AndAlso mbShowLightMap = True AndAlso muSettings.BumpMapPlanetModel = False Then
                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                    moDevice.SetTextureStageState(1, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)
                    moDevice.SetTextureStageState(2, TextureStageStates.AlphaArgument1, TextureArgument.TextureColor)

                    moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.Modulate)
                    moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Disable)
                    moDevice.SetTextureStageState(2, TextureStageStates.ColorOperation, TextureOperation.Disable)

                    'moDevice.SetTextureStageState(0, TextureStageStates.TextureCoordinateIndex, 0)
                    'moDevice.SetTextureStageState(2, TextureStageStates.TextureCoordinateIndex, 1)

                    moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                    moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                    moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)

                    moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                    moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
                    moDevice.SetTextureStageState(2, TextureStageStates.AlphaOperation, TextureOperation.Disable)

                    moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.Current)
                    moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument2, TextureArgument.Current)

                    moDevice.SetSamplerState(2, SamplerStageStates.MinFilter, TextureFilter.None)
                    moDevice.SetSamplerState(2, SamplerStageStates.MagFilter, TextureFilter.None)
                    moDevice.SetSamplerState(2, SamplerStageStates.MipFilter, TextureFilter.None)
                End If


                moDevice.SetTextureStageState(0, TextureStageStates.TextureCoordinateIndex, 0)
                ' moDevice.SetTextureStageState(1, TextureStageStates.TextureCoordinateIndex, 1)
                moDevice.SetTextureStageState(2, TextureStageStates.TextureCoordinateIndex, 2)

                If muSettings.BumpMapPlanetModel = True AndAlso lEnvirType = CurrentView.eSystemView AndAlso bInSpecialFXRange = True Then
                    moBumpPlanet.EndRender()
                End If

                'Now, do we have a ring
                If muSettings.RenderPlanetRings = True AndAlso mbHasRing = True Then
                    RenderMyPlanetRing()
                End If
                'If frmMain.mbAltKeyDown AndAlso GFXEngine.mbCaptureScreenshot = False AndAlso isAdmin() = True Then
                '    RenderPlanetAxis(Me)
                'End If
                If bInSpecialFXRange = True Then
                    If lEnvirType = CurrentView.eSystemView AndAlso Me.Atmosphere <> 0 Then
                        'Only enter here if distance from planet is within Entity clip plane

                        mfCloudRot -= 0.0005F
                        'If mfCloudRot > Math.PI * 2.0F Then mfCloudRot = 0.0F
                        If mfCloudRot < 0 Then mfCloudRot = Math.PI * 2.0F
                        matYaw = Matrix.RotationX(Math.PI / 2.0F)
                        matWorld = Matrix.Identity
                        matWorld.Multiply(matYaw)
                        matYaw = Matrix.RotationY(mfCloudRot)
                        matWorld.Multiply(matYaw)
                        matTemp = Matrix.Translation(LocX, LocY, LocZ)
                        matWorld.Multiply(matTemp)

                        moDevice.Transform.World = matWorld

                        If moCloudTex Is Nothing Then
                            moCloudTex = goResMgr.GetTexture("cloud1.dds", GFXResourceManager.eGetTextureType.ModelTexture)
                        End If
                        moDevice.SetTexture(0, moCloudTex)
                        moDevice.Material = mmatClouds
                        If moCloud Is Nothing OrElse moCloud.Disposed = True Then moCloud = goResMgr.CreateTexturedSphere(miPlanetRadius * 1.01F, 32, 32, 0)
                        moCloud.DrawSubset(0)
                    End If

                    'Now, check for guild billboards
                    If uBillboards Is Nothing = False Then
                        Dim matBase As Matrix = Matrix.Identity
                        Dim oBBMat As Material
                        With oBBMat
                            .Ambient = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                            .Diffuse = .Ambient
                            '.Specular = .Ambient
                            .Emissive = .Ambient
                        End With
                        moDevice.Material = oBBMat

                        matBase = Matrix.Identity
                        matBase = Matrix.Multiply(matBase, Matrix.RotationZ(timeGetTime * 0.002F))
                        matBase = Matrix.Multiply(matBase, Matrix.RotationX(-0.7853982F))

                        For X As Int32 = 0 To uBillboards.GetUpperBound(0)
                            If uBillboards(X).lGuildID > 0 Then
                                'ok, got one
                                If moBillboardMesh Is Nothing OrElse moBillboardMesh.Disposed = True Then moBillboardMesh = goResMgr.CreateTexturedSphere(150, 16, 16, 0, False)
                                uBillboards(X).CheckBBTex()
                                If uBillboards(X).oBBTex Is Nothing Then uBillboards(X).GenerateBBTex(moDevice)

                                Dim lOX As Int32 = 0
                                Dim lOZ As Int32 = 0
                                Select Case X
                                    Case 0
                                        lOX = -PlanetRadius : lOZ = 0
                                    Case 1
                                        lOX = -PlanetRadius : lOZ = -PlanetRadius
                                    Case 2
                                        lOX = PlanetRadius : lOZ = -PlanetRadius
                                    Case 3
                                        lOX = PlanetRadius : lOZ = 0
                                    Case 4
                                        lOX = PlanetRadius : lOZ = PlanetRadius
                                    Case 5
                                        lOX = -PlanetRadius : lOZ = PlanetRadius
                                End Select

                                matWorld = Matrix.Multiply(matBase, Matrix.Translation(Me.LocX + lOX, -1000, Me.LocZ + lOZ))
                                moDevice.Transform.World = matWorld

                                moDevice.SetTexture(0, uBillboards(X).oBBTex)
                                moBillboardMesh.DrawSubset(0)
                            End If
                        Next X
                    End If
                End If

                moDevice.RenderState.Wrap0 = 0 ' WrapCoordinates.Two
                moDevice.RenderState.Wrap1 = 0
                moDevice.RenderState.Wrap2 = 0

                matWorld = Nothing
            End If

        ElseIf lEnvirType = CurrentView.ePlanetMapView Then
            If Me.ParentSystem Is Nothing = False AndAlso Me.ParentSystem.ObjectID = 36 Then
                If glCurrentCycle - lLastBBRequest > 900 Then
                    lLastBBRequest = glCurrentCycle
                    Dim yMsg(7) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildBillboards).CopyTo(yMsg, lPos) : lPos += 2
                    Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                    goUILib.SendMsgToPrimary(yMsg)
                End If
            End If

            Dim lSmaller As Int32
            Dim lLarger As Int32
            Dim ptPos As Point
            Dim lTempW As Int32
            Dim lTempH As Int32
            Dim fMultX As Single
            Dim fMultY As Single

            'Ok, first clear our buffer
            moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.Black, 1, 0)

            If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

            lTempW = moDevice.PresentationParameters.BackBufferWidth
            lTempH = moDevice.PresentationParameters.BackBufferHeight
            lSmaller = CInt(Math.Min(lTempW, lTempH) * 0.8)
            lLarger = Math.Max(lTempW, lTempH)
            ptPos.X = CInt((lTempW - lSmaller) / 2)
            ptPos.Y = CInt((lTempH - lSmaller) / 2)

            If muSettings.gbPlanetMapDontRenderTerrain = True Then
                'Dim oTmpMat As Material
                'moDevice.Transform.World = Matrix.Identity
                'moDevice.RenderState.ZBufferWriteEnable = False
                'moDevice.RenderState.Lighting = False

                'oTmpMat.Ambient = System.Drawing.Color.FromArgb(255, 64, 64, 64)
                'oTmpMat.Diffuse = System.Drawing.Color.White
                'oTmpMat.Emissive = System.Drawing.Color.Black
                'oTmpMat.Specular = oTmpMat.Emissive
                'moDevice.Material = oTmpMat
                'moDevice.SetTexture(0, Nothing)     'explicitly tell the device no texture, a texture may still be there from a previous render action
                'moDevice.VertexFormat = CustomVertex.PositionColored.Format
                ''moSprite.Draw2D()
                'Dim oVerts(3) As CustomVertex.PositionColored
                'oVerts(0) = New CustomVertex.PositionColored(ptPos.X - 100, ptPos.Y - 100, -1, System.Drawing.Color.FromArgb(64, 255, 0, 0).ToArgb)
                'oVerts(1) = New CustomVertex.PositionColored(ptPos.X - 100, ptPos.Y, -1, System.Drawing.Color.FromArgb(64, 255, 0, 0).ToArgb)
                'oVerts(2) = New CustomVertex.PositionColored(ptPos.X, ptPos.Y - 100, -1, System.Drawing.Color.FromArgb(64, 255, 0, 0).ToArgb)
                'oVerts(3) = New CustomVertex.PositionColored(ptPos.X, ptPos.Y, -1, System.Drawing.Color.FromArgb(64, 255, 0, 0).ToArgb)
                'moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, oVerts)

                'oVerts(0) = New CustomVertex.PositionColored(ptPos.X - 100, -1, ptPos.Y - 100, System.Drawing.Color.FromArgb(64, 0, 255, 0).ToArgb)
                'oVerts(1) = New CustomVertex.PositionColored(ptPos.X - 100, -1, ptPos.Y, System.Drawing.Color.FromArgb(64, 0, 255, 0).ToArgb)
                'oVerts(2) = New CustomVertex.PositionColored(ptPos.X, -1, ptPos.Y - 100, System.Drawing.Color.FromArgb(64, 0, 255, 0).ToArgb)
                'oVerts(3) = New CustomVertex.PositionColored(ptPos.X, -1, ptPos.Y, System.Drawing.Color.FromArgb(64, 0, 255, 0).ToArgb)
                'moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, oVerts)

                'moDevice.RenderState.ZBufferEnable = True
                'moDevice.RenderState.Lighting = True
            Else
                If moTexture Is Nothing OrElse moTexture.Disposed = True Then
                    moTerrain.CreateTerrainPlanetTexture(moTexture)
                    If Me.MapTypeID = PlanetType.eGeoPlastic Then moLightMap = moTerrain.CreateTerrainLavaLightMap()
                    With moMaterial
                        .Diffuse = System.Drawing.Color.White
                        .Ambient = System.Drawing.Color.FromArgb(255, 18, 18, 18)
                        .Emissive = System.Drawing.Color.Black
                        .Specular = Color.FromArgb(128, 128, 128, 128)
                        .SpecularSharpness = 30
                    End With
                End If


                Dim lSrcRectWH As Int32 = muSettings.PlanetModelTextureWH - 1
                If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = False OrElse muSettings.HiResPlanetTexture = False Then lSrcRectWH = 255

                fMultX = CSng(lSrcRectWH / lSmaller)       'was terrainclass.width instead of 255      
                fMultY = CSng(lSrcRectWH / lSmaller)       'was terrainclass.height instead of 255

                'New Point(oLoc.X * fMultX, oLoc.Y * fMultY)
                'Now, set our map onto the screen, not sure I like the results but it works...


                moDevice.RenderState.Lighting = False
                moDevice.RenderState.ZBufferEnable = False
                EnsureSpritesValid(moDevice)
                moSprite.SetWorldViewLH(Matrix.Identity, moDevice.Transform.View)
                moSprite.Begin(SpriteFlags.None)
                'Draw2D: Texture, SrcRect, SizeRect, RotPoint, RotAngle, Point, Color
                'oSpr.Draw2D(moTexture, System.Drawing.Rectangle.FromLTRB(0, 0, 255, 255), System.Drawing.Rectangle.FromLTRB(0, 0, lSmaller, lSmaller), System.Drawing.Point.Empty, 0, ptPos, System.Drawing.Color.White)
                moSprite.Draw2D(moTexture, System.Drawing.Rectangle.FromLTRB(0, 0, lSrcRectWH, lSrcRectWH), System.Drawing.Rectangle.FromLTRB(ptPos.X, ptPos.Y, ptPos.X + lSmaller, ptPos.Y + lSmaller), System.Drawing.Point.Empty, 0, New Point(CInt(ptPos.X * fMultX), CInt(ptPos.Y * fMultY)), System.Drawing.Color.White)
                'moSprite.Draw2D(moTexture, System.Drawing.Rectangle.FromLTRB(0, 0, TerrainClass.Width - 1, TerrainClass.Height - 1), New System.Drawing.SizeF(lSmaller, lSmaller), PointF.Empty, 0, New PointF(ptPos.X * fMultX, ptPos.Y * fMultY), Color.White)
                moSprite.End()
                moDevice.RenderState.ZBufferEnable = True
                moDevice.RenderState.Lighting = True
            End If
            If goUILib Is Nothing = False Then
                If uBillboards Is Nothing = False Then
                    For X As Int32 = 0 To uBillboards.GetUpperBound(0)
                        Dim sBBName As String = "frmGuildBillboard" & (X + 1)
                        If uBillboards(X).lGuildID > 0 Then
                            Dim oWin As UIWindow = goUILib.GetWindow(sBBName)
                            If oWin Is Nothing Then oWin = New frmGuildBillboard(goUILib, X + 1)
                            If oWin.Visible = False Then oWin.Visible = True
                        Else
                            goUILib.RemoveWindow(sBBName)
                        End If
                    Next X
                Else
                    For X As Int32 = 1 To 6
                        goUILib.RemoveWindow("frmGuildBillboard" & X)
                    Next
                End If
            End If

        ElseIf lEnvirType = CurrentView.ePlanetView Then
            If goUILib Is Nothing = False Then
                Dim oFrm As UIWindow = goUILib.GetWindow("frmPlanetMapControl")
                If Not oFrm Is Nothing Then
                    oFrm.Visible = False
                End If
            End If
            Dim lRadius1 As Int32 = 0

            SetPlanetAtmosphere()

            Dim lTmpIdx As Int32 = 0
            If ParentSystem.StarType1Idx <> -1 Then
                lRadius1 = goStarTypes(ParentSystem.StarType1Idx).StarRadius
                SetLight(1, goStarTypes(ParentSystem.StarType1Idx), lRadius1)
                lTmpIdx += 1
            End If
            If ParentSystem.StarType2Idx <> -1 Then
                SetLight(2, goStarTypes(ParentSystem.StarType2Idx), lRadius1)
                lTmpIdx += 2
            End If
            If ParentSystem.StarType3Idx <> -1 Then
                SetLight(3, goStarTypes(ParentSystem.StarType3Idx), lRadius1)
                lTmpIdx += 3
            End If

            For X As Int32 = lTmpIdx To moDevice.Lights.Count - 1
                If moDevice.Lights(X).Enabled = True Then moDevice.Lights(X).Enabled = False
            Next X

            For X As Int32 = 0 To FireworksMgr.lLightUB
                If FireworksMgr.yLightUsed(X) <> 0 Then
                    moDevice.Lights(lTmpIdx).FromLight(FireworksMgr.Lights(X))
                    moDevice.Lights(lTmpIdx).Enabled = True
                    moDevice.Lights(lTmpIdx).Update()
                    lTmpIdx += 1
                End If
            Next X

            moTerrain.Render(muSettings.TerrainFarClippingPlane)

        End If
    End Sub

    Public Sub DoWaterRenderPass()
        If moTerrain Is Nothing = False Then moTerrain.DoWaterRenderPass()
    End Sub

    Private Sub SavePlanetTexture()
        Try
            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            sPath &= "cache"
            If Exists(sPath) = False Then
                MkDir(sPath)
            End If
            sPath &= "\planet_" & Me.ObjectID
            If muSettings.HiResPlanetTexture = True Then
                sPath &= "_H"
            Else
                sPath &= "_L"
            End If


            If Not Exists(sPath & ".dds") Then
                'Debug.Print("Saving " & Me.ObjectID)
                TextureLoader.Save(sPath & ".dds", ImageFileFormat.Dds, moTexture)
                'TextureLoader.Save(sPath & ".png", ImageFileFormat.Png, moTexture)
                'TextureLoader.Save(sPath & ".jpg", ImageFileFormat.Jpg, moTexture)
                'TextureLoader.Save(sPath & ".bmp", ImageFileFormat.Bmp, moTexture)
            End If
            bSaved = True
        Catch ex As Exception
        End Try
    End Sub

    Private Function LoadCachedPlanet() As Boolean
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then
            moTexture = Nothing
            Return False
        End If

        Dim moDevice As Device = GFXEngine.moDevice

        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"
        sPath &= "cache\planet_"
        sPath &= Me.ObjectID
        If muSettings.HiResPlanetTexture = True Then
            sPath &= "_H"
        Else
            sPath &= "_L"
        End If
        sPath &= ".dds"

        If Exists(sPath) Then
            Try
                moTexture = TextureLoader.FromFile(moDevice, sPath, 0, 0, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Box, 0)
                'Debug.Print("Loaded Cached " & Me.ObjectID)
            Catch
                Return False
            End Try
            Return True
        End If
    End Function

#Region "  For Planet Ring Rendering  "

#Region "  Ring Resource Management  "
    Private Const ml_RING_TYPE_CNT As Int32 = 8
    Private Const ml_RING_MAT_CNT As Int32 = 15

    Private Shared moRingMats() As Material
    Private Shared moRingTexs() As Texture
    Private Shared Sub ConfigureRingResources()
        If moRingTexs Is Nothing Then
            ReDim moRingTexs(ml_RING_TYPE_CNT - 1)
        End If
        For X As Int32 = 0 To moRingTexs.GetUpperBound(0)
            If moRingTexs(X) Is Nothing OrElse moRingTexs(X).Disposed = True Then
                moRingTexs(X) = goResMgr.GetTexture("Rings" & (X + 1) & ".dds", GFXResourceManager.eGetTextureType.ModelTexture, "pring.pak")
            End If
        Next X

        If moRingMats Is Nothing Then
            ReDim moRingMats(ml_RING_MAT_CNT - 1)
            For X As Int32 = 0 To ml_RING_MAT_CNT - 1
                With moRingMats(X)
                    .Diffuse = GetClrFromIndex(X)
                    .Specular = .Diffuse
                    '.Ambient = Color.FromArgb(255, 92, 92, 92)

                    .Ambient = System.Drawing.Color.FromArgb(255, _
                     CInt((.Diffuse.R / 255.0F) * 128), CInt((.Diffuse.G / 255.0F) * 128), CInt((.Diffuse.B / 255.0F) * 128))
                    If .Ambient.R < 64 AndAlso .Ambient.G < 64 AndAlso .Ambient.B < 64 Then
                        .Ambient = System.Drawing.Color.FromArgb(255, .Ambient.R * 2, .Ambient.G * 2, .Ambient.B * 2)
                    End If

                    .SpecularSharpness = 10
                End With
            Next X
        End If
    End Sub
    Private Shared Function GetClrFromIndex(ByVal lIndex As Int32) As System.Drawing.Color
        Select Case lIndex
            Case 1  'silver
                Return Color.FromArgb(255, 192, 192, 192)
            Case 2  'gray
                Return Color.FromArgb(255, 128, 128, 128)
            Case 3  'dirt
                Return Color.FromArgb(255, 64, 64, 64)
            Case 4  'brown
                Return Color.FromArgb(255, 128, 64, 0)
            Case 5  'red
                Return Color.FromArgb(255, 255, 0, 0)
            Case 6  'dark red
                Return Color.FromArgb(255, 128, 0, 0)
            Case 7  'green
                'Return Color.FromArgb(255, 0, 255, 0)
                Return Color.FromArgb(255, 0, 64, 0)
            Case 8  'dark green
                Return Color.FromArgb(255, 0, 128, 0)
            Case 9  'orange
                Return Color.FromArgb(255, 255, 128, 0)
            Case 10 'yellow
                'Return Color.FromArgb(255, 255, 255, 0)
                Return Color.FromArgb(255, 128, 128, 0)
            Case 11 'blue
                Return Color.FromArgb(255, 0, 0, 255)
            Case 12 'dark blue
                Return Color.FromArgb(255, 0, 64, 128)
            Case 13 'teal
                Return Color.FromArgb(255, 0, 255, 255)
            Case 14 'light blue
                Return Color.FromArgb(255, 0, 128, 255)
            Case Else       '0 = white
                Return Color.FromArgb(255, 255, 255, 255)
        End Select
    End Function
#End Region

    Public mbHasRing As Boolean = False   'indicates the planet has a ring...
    Private moRing As Mesh          'planet ring mesh
    Private mfRingAngle As Single

    Private Structure RingLayer
        Public fRotateSpeedDir As Single
        Public fRotation As Single
        Public lMaterialIdx As Int32
        Public lTextureIdx As Int32
    End Structure
    Private muRingLayers() As RingLayer

    Private mlInnerRingRadius As Int32 = 0
    Private mlOuterRingRadius As Int32 = 0
    Public Sub SetRingProps(ByVal lInnerRingRadius As Int32, ByVal lOuterRingRadius As Int32, ByVal lRingDiffuse As Int32)
        mlInnerRingRadius = lInnerRingRadius
        mlOuterRingRadius = lOuterRingRadius
        If lInnerRingRadius = 0 OrElse lOuterRingRadius = 0 Then
            mbHasRing = False
        Else
            mbHasRing = True
        End If
        If moRing Is Nothing = False Then moRing.Dispose()
        moRing = Nothing
    End Sub
    Private Sub SetUpRingMesh()
        moRing = goResMgr.CreatePlanetRing(mlInnerRingRadius, mlOuterRingRadius, 64)
        moRing.ComputeNormals()

        'set our seed
        'Call Rnd(-1)
        Dim oRandom As New Random(Me.ObjectID)
        'Randomize(Me.ObjectID)

        'all rings have at least 2 layers...
        Dim lRingCnt As Int32 = 2
        '75% of the time they have 3
        If oRandom.NextDouble() * 100 < 75 Then
            lRingCnt += 1
        End If

        mfRingAngle = CSng(oRandom.NextDouble() * 0.3F + 0.2F)
        Dim bNegDir As Boolean = oRandom.NextDouble() * 100 < 50
        If bNegDir = True Then mfRingAngle = -mfRingAngle

        ReDim muRingLayers(lRingCnt - 1)
        For X As Int32 = 0 To muRingLayers.GetUpperBound(0)
            With muRingLayers(X)
                .fRotateSpeedDir = CSng(oRandom.NextDouble() / mlOuterRingRadius)
                If bNegDir = True Then .fRotateSpeedDir = -.fRotateSpeedDir
                .fRotation = 0
                .lMaterialIdx = CInt(Math.Floor(oRandom.NextDouble() * ml_RING_MAT_CNT))
                .lTextureIdx = CInt(Math.Floor(oRandom.NextDouble() * ml_RING_TYPE_CNT))
            End With
        Next X

        oRandom = Nothing 
    End Sub
    Private Sub RenderMyPlanetRing()
        Dim moDevice As Device = GFXEngine.moDevice

        If mbHasRing = False Then Return
        If moRing Is Nothing OrElse moRing.Disposed = True Then
            SetUpRingMesh()
            Threading.Thread.Sleep(50)
        End If

        '1) ensure our resources exist
        ConfigureRingResources()
        If moRing Is Nothing OrElse moRing.Disposed = True Then
            moRing = goResMgr.CreatePlanetRing(mlInnerRingRadius, mlOuterRingRadius, 64)
            Threading.Thread.Sleep(50)
        End If

        '2) Set our renderstates
        Dim bAlphaBlend As Boolean = moDevice.RenderState.AlphaBlendEnable
        Dim bZBufferWriteEnable As Boolean = moDevice.RenderState.ZBufferWriteEnable
        Dim uCull As Cull = moDevice.RenderState.CullMode
        'Dim fPSC As Single = moDevice.RenderState.PointScaleC
        'Dim fPSA As Single = moDevice.RenderState.PointScaleA
        'Dim bPSEnable As Boolean = moDevice.RenderState.PointScaleEnable
        'Dim fPointSize As Single = moDevice.RenderState.PointSize
        Dim uWrap0 As WrapCoordinates = moDevice.RenderState.Wrap0

        moDevice.RenderState.AlphaBlendEnable = True
        moDevice.RenderState.ZBufferWriteEnable = False
        moDevice.RenderState.CullMode = Cull.None
        'moDevice.RenderState.PointScaleC = 0
        'moDevice.RenderState.PointScaleA = 1
        'moDevice.RenderState.PointScaleEnable = False
        'moDevice.RenderState.PointSize = 1
        moDevice.RenderState.Wrap0 = 0

        '3) Loop thru our layers and do our rendering
        Dim lYVal As Int32 = -10
        For X As Int32 = muRingLayers.GetUpperBound(0) To 0 Step -1
            'ok, get our world matrix
            Dim matWorld As Matrix = Matrix.Identity

            With muRingLayers(X)
                .fRotation += .fRotateSpeedDir
                matWorld.Multiply(Matrix.RotationY(.fRotation))
                matWorld.Multiply(Matrix.RotationZ(mfRingAngle)) '0.5235988F))
                matWorld.Multiply(Matrix.Translation(LocX, LocY + lYVal, LocZ))
                moDevice.Transform.World = matWorld
                moDevice.SetTexture(0, moRingTexs(.lTextureIdx))
                moDevice.Material = moRingMats(.lMaterialIdx)
            End With
            moRing.DrawSubset(0)
            'increment our y val
            lYVal += 5
        Next X

        '4) Reset our renderstates
        moDevice.RenderState.AlphaBlendEnable = bAlphaBlend
        moDevice.RenderState.ZBufferWriteEnable = bZBufferWriteEnable
        moDevice.RenderState.CullMode = uCull
        'moDevice.RenderState.PointScaleC = fPSC
        'moDevice.RenderState.PointScaleA = fPSA
        'moDevice.RenderState.PointScaleEnable = bPSEnable
        'moDevice.RenderState.PointSize = fPointSize
        moDevice.RenderState.Wrap0 = uWrap0


        'Dim matWorld As Matrix = Matrix.Identity
        'moDevice.RenderState.CullMode = Cull.None

        'matWorld.Multiply(Matrix.RotationZ(0.5235988F))
        'matWorld.Multiply(Matrix.Translation(LocX, LocY, LocZ))
        'moDevice.Transform.World = matWorld

        'If moRingTex Is Nothing OrElse moRingTex.Disposed = True Then moRingTex = goResMgr.GetTexture("rings.dds", GFXResourceManager.eGetTextureType.NoSpecifics)
        'moDevice.SetTexture(0, moRingTex)
        'moDevice.Material = moRingMat
        'If moRing Is Nothing OrElse moRing.Disposed = True Then moRing = goResMgr.CreatePlanetRing(mlInnerRingRadius, mlOuterRingRadius, 32)
        'moRing.DrawSubset(0)
        'moDevice.RenderState.CullMode = Cull.CounterClockwise
    End Sub

    Private Sub RenderPlanetAxis(ByVal oPlanet As Planet)
        Dim moDevice As Device = GFXEngine.moDevice

        Dim lPrevCullMode As Cull = moDevice.RenderState.CullMode
        Dim lPrevVertexFormat As VertexFormats = moDevice.VertexFormat
        Dim lPrevSrcBlnd As Blend = moDevice.RenderState.SourceBlend
        Dim lPrevDestBlnd As Blend = moDevice.RenderState.DestinationBlend
        Dim lPrevAlplaBlnd As Boolean = moDevice.RenderState.AlphaBlendEnable

        'Debug.Print("Planet Line @ (" & Me.LocX & ", " & Me.LocY & ", " & Me.LocZ & ")")
        Dim oMat As Material
        With oMat
            .Ambient = Color.Black
            .Diffuse = Color.Black
            .Emissive = Color.Teal
            .Specular = Color.Black
        End With

        With moDevice
            .SetTexture(0, Nothing)
            .Material = oMat
            .RenderState.CullMode = Cull.Clockwise

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

        Dim uVerts(3) As CustomVertex.PositionColored

        Dim zSize As Single = 25
        Dim ySize As Int32 = 2500
        Dim yOffset As Int32 = oPlanet.miPlanetRadius

        Dim lLineColor As Int32 = System.Drawing.Color.FromArgb(192, 255, 255, 255).ToArgb
        uVerts(0) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset, Me.LocZ + zSize, lLineColor)
        uVerts(1) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset, Me.LocZ - zSize, lLineColor)
        uVerts(2) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset + ySize, Me.LocZ + zSize, lLineColor)
        uVerts(3) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset + ySize, Me.LocZ - zSize, lLineColor)
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

        uVerts(0) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset, Me.LocZ - zSize, lLineColor)
        uVerts(1) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset, Me.LocZ + zSize, lLineColor)
        uVerts(2) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset + ySize, Me.LocZ - zSize, lLineColor)
        uVerts(3) = New CustomVertex.PositionColored(Me.LocX, Me.LocY + yOffset + ySize, Me.LocZ + zSize, lLineColor)
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

        uVerts(0) = New CustomVertex.PositionColored(Me.LocX + zSize, Me.LocY + yOffset, Me.LocZ, lLineColor)
        uVerts(1) = New CustomVertex.PositionColored(Me.LocX - zSize, Me.LocY + yOffset, Me.LocZ, lLineColor)
        uVerts(2) = New CustomVertex.PositionColored(Me.LocX + zSize, Me.LocY + yOffset + ySize, Me.LocZ, lLineColor)
        uVerts(3) = New CustomVertex.PositionColored(Me.LocX - zSize, Me.LocY + yOffset + ySize, Me.LocZ, lLineColor)
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

        uVerts(0) = New CustomVertex.PositionColored(Me.LocX - zSize, Me.LocY + yOffset, Me.LocZ, lLineColor)
        uVerts(1) = New CustomVertex.PositionColored(Me.LocX + zSize, Me.LocY + yOffset, Me.LocZ, lLineColor)
        uVerts(2) = New CustomVertex.PositionColored(Me.LocX - zSize, Me.LocY + yOffset + ySize, Me.LocZ, lLineColor)
        uVerts(3) = New CustomVertex.PositionColored(Me.LocX + zSize, Me.LocY + yOffset + ySize, Me.LocZ, lLineColor)
        moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, uVerts)

        'Reset our renderstates
        With moDevice
            .RenderState.ZBufferEnable = True
            .RenderState.Lighting = True
            .SetTexture(0, Nothing)

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

    Public ReadOnly Property InnerRingRadius() As Int32
        Get
            Return mlInnerRingRadius
        End Get
    End Property
    Public ReadOnly Property OuterRingRadius() As Int32
        Get
            Return mlOuterRingRadius
        End Get
    End Property
#End Region

	Public Sub RenderMiniMap(ByRef oMapBlips As Texture, ByVal yBlipsAlpha As Byte)
        'Now, make our minimap
        If goUILib.yRenderUI = 0 Then Return

        Try
            Dim moDevice As Device = GFXEngine.moDevice
            If muSettings.ShowMiniMap = True OrElse (goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255) Then

                If moTexture Is Nothing OrElse moTexture.Disposed = True Then
                    moTerrain.CreateTerrainPlanetTexture(moTexture)
                    If Me.MapTypeID = PlanetType.eGeoPlastic Then moLightMap = moTerrain.CreateTerrainLavaLightMap()
                End If

                'Figure out our draw area
                Dim rcSrc As Rectangle
                Dim lCX As Int32 = 0
                Dim lCZ As Int32 = 0
                Dim lSize As Int32 = 255

                Dim lCameraOffsetX As Int32 = 0
                Dim lCameraOffsetZ As Int32 = 0

                If muSettings.PlanetMinimapZoomLevel = 0 Then
                    'Zoomed all the way out... no changes
                ElseIf muSettings.PlanetMinimapZoomLevel = 1 Then
                    lCX = CInt(((CInt(goCamera.mlCameraX / moTerrain.CellSpacing) + (TerrainClass.Width \ 2)) / 240) * 255)
                    lCZ = CInt(((CInt(goCamera.mlCameraZ / moTerrain.CellSpacing) + (TerrainClass.Height \ 2)) / 240) * 255)
                    lSize = TerrainClass.Width \ 2
                ElseIf muSettings.PlanetMinimapZoomLevel = 2 Then
                    lCX = CInt(((CInt(goCamera.mlCameraX / moTerrain.CellSpacing) + (TerrainClass.Width \ 2)) / 240) * 255)
                    lCZ = CInt(((CInt(goCamera.mlCameraZ / moTerrain.CellSpacing) + (TerrainClass.Height \ 2)) / 240) * 255)
                    lSize = TerrainClass.Width \ 4
                Else
                    lCX = CInt(((CInt(goCamera.mlCameraX / moTerrain.CellSpacing) + (TerrainClass.Width \ 2)) / 240) * 255)
                    lCZ = CInt(((CInt(goCamera.mlCameraZ / moTerrain.CellSpacing) + (TerrainClass.Height \ 2)) / 240) * 255)
                    lSize = TerrainClass.Width \ 6
                End If
                Dim lHalfSize As Int32 = lSize \ 2

                lCX = 255 - lCX

                If lCX < lHalfSize Then
                    lCameraOffsetX = lCX
                    lCX = lHalfSize
                    lCameraOffsetX -= lCX
                End If
                If lCZ < lHalfSize Then
                    lCameraOffsetZ = lCZ
                    lCZ = lHalfSize
                    lCameraOffsetZ -= lCZ
                End If
                If lCX + lHalfSize > 255 Then
                    lCameraOffsetX = lCX
                    lCX = 255 - lHalfSize
                    lCameraOffsetX -= lCX
                End If
                If lCZ + lHalfSize > 255 Then
                    lCameraOffsetZ = lCZ
                    lCZ = 255 - lHalfSize
                    lCameraOffsetZ -= lCZ
                End If

                rcSrc = New Rectangle(lCX - lHalfSize, lCZ - lHalfSize, lSize, lSize)

                Dim rcTexSrc As Rectangle = rcSrc
                Dim lTempVal As Int32 = muSettings.PlanetModelTextureWH
                If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = False OrElse muSettings.HiResPlanetTexture = False Then lTempVal = 255
                Dim fScaleVal As Single = lTempVal / 255.0F
                rcTexSrc.X = CInt(rcTexSrc.X * fScaleVal)
                rcTexSrc.Y = CInt(rcTexSrc.Y * fScaleVal)
                rcTexSrc.Width = CInt(rcTexSrc.Width * fScaleVal)
                rcTexSrc.Height = CInt(rcTexSrc.Height * fScaleVal)

                If NewTutorialManager.TutorialOn = True Then
                    muSettings.MiniMapWidthHeight = 120
                End If

                Dim fMultX As Single = CSng(rcSrc.Width / muSettings.MiniMapWidthHeight)
                Dim fMultY As Single = CSng(rcSrc.Height / muSettings.MiniMapWidthHeight)

                moDevice.RenderState.Lighting = False
                moDevice.RenderState.ZBufferWriteEnable = False
                EnsureSpritesValid(moDevice)
                moSprite.SetWorldViewLH(Matrix.Identity, moDevice.Transform.View)
                moSprite.Begin(SpriteFlags.None)
                moSprite.Draw2D(moTexture, rcTexSrc, System.Drawing.Rectangle.FromLTRB(muSettings.MiniMapLocX, muSettings.MiniMapLocY, muSettings.MiniMapLocX + muSettings.MiniMapWidthHeight, muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight), System.Drawing.Point.Empty, 0, New Point(CInt(muSettings.MiniMapLocX * fMultX), CInt(muSettings.MiniMapLocY * fMultY)), System.Drawing.Color.White)
                If oMapBlips Is Nothing = False Then moSprite.Draw2D(oMapBlips, rcSrc, System.Drawing.Rectangle.FromLTRB(muSettings.MiniMapLocX, muSettings.MiniMapLocY, muSettings.MiniMapLocX + muSettings.MiniMapWidthHeight, muSettings.MiniMapLocY + muSettings.MiniMapWidthHeight), System.Drawing.Point.Empty, 0, New Point(CInt(muSettings.MiniMapLocX * fMultX), CInt(muSettings.MiniMapLocY * fMultY)), System.Drawing.Color.FromArgb(yBlipsAlpha, 255, 255, 255))
                moSprite.End()

                moDevice.RenderState.Lighting = True
                moDevice.RenderState.ZBufferWriteEnable = True

                'Show our Minimap Zoom control window...
                If goUILib Is Nothing = False Then
                    If goUILib.GetWindow("frmMinimapZoom") Is Nothing Then
                        Dim frmMinMapZoom As New frmMinimapZoom(goUILib)
                        frmMinMapZoom.Visible = True
                        frmMinMapZoom = Nothing
                    End If
                End If

                'Now, render the Camera Angle
                Dim bFog As Boolean = moDevice.RenderState.FogEnable
                moDevice.RenderState.FogEnable = False
                'Render our camera view...
                'Dim uVerts(2) As CustomVertex.TransformedColored
                'Dim vecTemp As Vector3 = New Vector3(goCamera.mlCameraAtX - goCamera.mlCameraX, 0, goCamera.mlCameraAtZ - goCamera.mlCameraZ)
                'vecTemp.Normalize()
                'vecTemp.Multiply(24)


                'Dim fCameraX As Single = muSettings.MiniMapLocX + (muSettings.MiniMapWidthHeight \ 2) + lCameraOffsetX
                '((CSng(goCamera.mlCameraX / moTerrain.CellSpacing) + (TerrainClass.Width \ 2)) / 240.0F) * 255.0F
                'Dim fCameraY As Single = muSettings.MiniMapLocY + (muSettings.MiniMapWidthHeight \ 2) + lCameraOffsetZ
                ' ((CSng(goCamera.mlCameraZ / moTerrain.CellSpacing) + (TerrainClass.Height \ 2)) / 240.0F) * 255.0F

                Dim lBaseR As Int32 = 255
                Dim lBaseG As Int32 = 255
                Dim lBaseB As Int32 = 255

                Select Case Me.MapTypeID
                    Case PlanetType.eAcidic
                        With muSettings.AcidMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                    Case PlanetType.eAdaptable
                        With muSettings.AdaptableMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                    Case PlanetType.eBarren
                        With muSettings.BarrenMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                    Case PlanetType.eDesert
                        With muSettings.DesertMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                    Case PlanetType.eGeoPlastic
                        With muSettings.LavaMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                    Case PlanetType.eTerran
                        With muSettings.TerranMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                    Case PlanetType.eTundra
                        With muSettings.IceMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                    Case PlanetType.eWaterWorld
                        With muSettings.WaterworldMinimapAngle
                            lBaseR = .R : lBaseG = .G : lBaseB = .B
                        End With
                End Select

                Dim fLineAngle As Single = 180 - LineAngleDegrees(goCamera.mlCameraX, goCamera.mlCameraZ, goCamera.mlCameraAtX, goCamera.mlCameraAtZ) '- 180
                If fLineAngle < 0 Then fLineAngle += 360
                Dim lHalfWidth As Int32 = (muSettings.MiniMapWidthHeight \ 2) - 8
                BPSprite.Draw2DOnceRotation(moDevice, goUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eDirectionalArrow), New Rectangle(muSettings.MiniMapLocX + lHalfWidth, muSettings.MiniMapLocY + lHalfWidth, 16, 16), System.Drawing.Color.FromArgb(255, lBaseR, lBaseG, lBaseB), 256, 256, fLineAngle)

                'uVerts(0) = New CustomVertex.TransformedColored(muSettings.MiniMapLocX + fCameraX, muSettings.MiniMapLocY + fCameraY, 0.0F, 1.0F, System.Drawing.Color.FromArgb(255, lBaseR, lBaseG, lBaseB).ToArgb)

                'Dim fX As Single = vecTemp.X
                'Dim fZ As Single = vecTemp.Z
                'RotatePoint(0, 0, fX, fZ, -22.5F)
                ''RotatePoint(0, 0, fX, fZ, 157.5F)
                'uVerts(1) = New CustomVertex.TransformedColored(fCameraX - fX, fCameraY + fZ, 1.0F, 0.0F, System.Drawing.Color.FromArgb(64, lBaseR, lBaseG, lBaseB).ToArgb)

                'fX = vecTemp.X
                'fZ = vecTemp.Z
                'RotatePoint(0, 0, fX, fZ, 22.5F)
                ''RotatePoint(0, 0, fX, fZ, 202.5F)
                'uVerts(2) = New CustomVertex.TransformedColored(fCameraX - fX, fCameraY + fZ, 1.0F, 0.0F, System.Drawing.Color.FromArgb(64, lBaseR, lBaseG, lBaseB).ToArgb)

                'moDevice.Transform.World = Matrix.Identity
                'moDevice.VertexFormat = CustomVertex.TransformedColored.Format
                'moDevice.RenderState.Lighting = False
                'moDevice.RenderState.CullMode = Cull.None
                'moDevice.RenderState.AlphaBlendEnable = True
                'moDevice.SetTexture(0, Nothing)
                'moDevice.DrawUserPrimitives(PrimitiveType.TriangleList, 1, uVerts)
                moDevice.RenderState.CullMode = Cull.CounterClockwise
                moDevice.RenderState.Lighting = True
                moDevice.RenderState.FogEnable = bFog
            End If
        Catch
        End Try
    End Sub

	Private Sub SetLight(ByVal lStarID As Int32, ByVal oStarType As StarType, ByVal lStar1Radius As Int32)
		Dim lSunX As Single
		Dim lSunY As Single
		Dim lSunZ As Single
		Dim fOffset As Single = (CSng(timeGetTime) - CSng(mlLastUpdate)) / CSng(RotationDelay)
        Dim fAngle As Single
        If b_LOCK_TIME_OF_DAY = True Then fOffset = 0

		Dim lLightID As Int32
		Dim lTemp As Int32

        Dim moDevice As Device = GFXEngine.moDevice

		'Set up our initial locations for MIDNIGHT
		lSunX = 0 : lSunY = Math.Abs(LocX) + Math.Abs(LocZ) : lSunZ = 0
		Select Case lStarID
			Case 1
				lSunX = 0
				lSunZ = 0
				fAngle = LineAngleDegrees(CInt(lSunX), CInt(lSunZ), LocX, LocZ)
				lLightID = ge_Light_Idx.eStar1Light
			Case 2	'secondary star
				lSunX = lStar1Radius + oStarType.StarRadius + 1000
				lSunZ = 0
				fAngle = LineAngleDegrees(CInt(lSunX), CInt(lSunZ), LocX, LocZ)
				lLightID = ge_Light_Idx.eStar2Light
			Case 3	'third star
				lSunX = -(lStar1Radius + oStarType.StarRadius + 1000)
				lSunZ = -lSunX
				fAngle = LineAngleDegrees(CInt(lSunX), CInt(lSunZ), LocX, LocZ)
				lLightID = ge_Light_Idx.eStar3Light
		End Select

		fAngle += (((CSng(RotateAngle + fOffset)) - 150) / 10.0F)
		If fAngle > 360 Then fAngle -= 360
		If fAngle < 0 Then fAngle += 360

		RotatePoint(0, 0, lSunX, lSunY, fAngle)

		'a quick normalize and scale...
		lTemp = CInt(Math.Abs(lSunX) + Math.Abs(lSunY) + Math.Abs(lSunZ))
		lSunX = (lSunX / lTemp) * 500
		lSunY = (lSunY / lTemp) * 500
		lSunZ = (lSunZ / lTemp) * 500

		If lSunX > 0 Then
			lSunY += 65
		End If

		moDevice.Lights(lLightID).FromLight(oStarType.StarLight)
		With moDevice.Lights(lLightID)
			.Type = LightType.Directional
			.Direction = Nothing
			.Direction = New Vector3(lSunX, lSunY, lSunZ)
			.Enabled = True
			.Update()
		End With

		'Now, represent our sun in sprite form
		If moLensFlare Is Nothing OrElse moLensFlare.Disposed = True Then moLensFlare = goResMgr.GetTexture("LensFlare.dds", GFXResourceManager.eGetTextureType.NoSpecifics)

		Dim fMult As Single = CSng(muSettings.TerrainFarClippingPlane / 300)
		Dim matWorld As Matrix = Matrix.Identity
		Dim matTemp As Matrix = Matrix.Identity
		Dim lCamXMod As Int32
		matTemp.Scale(fMult, fMult, fMult)
		matWorld.Multiply(matTemp)
		matTemp = Matrix.Identity
		lCamXMod = goCamera.mlCameraAtX
		matTemp.Translate(((-lSunX) * fMult) + lCamXMod, (-lSunY) * fMult, goCamera.mlCameraAtZ)
		matWorld.Multiply(matTemp)
		matTemp = Matrix.Identity
		matTemp = Nothing

		moDevice.Transform.World = matWorld

		moDevice.RenderState.SourceBlend = Blend.SourceColor
		moDevice.RenderState.DestinationBlend = Blend.DestinationColor

		Dim lPrevBlendOp As BlendOperation = moDevice.RenderState.BlendOperation

		If moStarSprite Is Nothing Then
			Device.IsUsingEventHandlers = False
			moStarSprite = New Sprite(moDevice)
			Device.IsUsingEventHandlers = True
		End If

		With moStarSprite
			.SetWorldViewLH(matWorld, moDevice.Transform.View)
			.Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)

			moDevice.RenderState.SourceBlend = Blend.SourceColor
			moDevice.RenderState.DestinationBlend = Blend.One
			moDevice.RenderState.BlendOperation = BlendOperation.Add
			moDevice.RenderState.ZBufferEnable = False

            .Draw(moLensFlare, System.Drawing.Rectangle.Empty, New Vector3(128, 128, 0), New Vector3(0, 0, 0), oStarType.MatDiffuse)
            .End()
        End With

        'fMult *= 
        Dim fSizeMult As Single = (oStarType.StarRadius) / 50000.0F
        matWorld = Matrix.Identity
        matTemp = Matrix.Identity
        matTemp.Scale(fMult * fSizeMult, fMult * fSizeMult, fMult * fSizeMult)
        matWorld.Multiply(matTemp)
        matTemp = Matrix.Identity
        lCamXMod = goCamera.mlCameraAtX
        matTemp.Translate(((-lSunX) * fMult) + lCamXMod, (-lSunY) * fMult, goCamera.mlCameraAtZ)
        matWorld.Multiply(matTemp)
        matTemp = Matrix.Identity
        matTemp = Nothing

        moDevice.Transform.World = matWorld

        With moStarSprite
            .SetWorldViewLH(matWorld, moDevice.Transform.View)
            .Begin(SpriteFlags.AlphaBlend Or SpriteFlags.Billboard Or SpriteFlags.ObjectSpace)

            moDevice.RenderState.SourceBlend = Blend.SourceColor
            moDevice.RenderState.DestinationBlend = Blend.One
            moDevice.RenderState.BlendOperation = BlendOperation.Add
            moDevice.RenderState.ZBufferEnable = False

            .Draw(goResMgr.GetTexture("WpnFire.dds", GFXResourceManager.eGetTextureType.NoSpecifics), New Rectangle(96, 32, 31, 31), New Vector3(16, 16, 0), New Vector3(0, 0, 0), Color.White)
            .End()
        End With



		moDevice.RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
		moDevice.RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
		moDevice.RenderState.BlendOperation = lPrevBlendOp
		moDevice.RenderState.ZBufferEnable = True
	End Sub

	Private Sub SetPlanetAtmosphere()
		Dim lColor As System.Drawing.Color
		'Dim lColorTemp As System.Drawing.Color
		'Dim X As Int32
		'Dim lTemp As Int32
		Dim matWorld As Matrix = Matrix.Identity
        Dim matTemp As Matrix = Matrix.Identity
        Dim moDevice As Device = GFXEngine.moDevice

        If Atmosphere <> 0 Then
            Try
                lColor = ColorSphereSetColor()
            Catch
            End Try
        Else
            lColor = moEnvirColor
        End If

		'If we are calling this, then clear our modevice...
		moDevice.Clear(ClearFlags.Target Or ClearFlags.ZBuffer, lColor, 1, 0)

		'disable our fog, and clear our material and texture
		moDevice.RenderState.FogEnable = False

		'ParentSystem.Test()
		RenderOrbitalBodies()

		moDevice.Material = moSkyboxMat
		moDevice.SetTexture(0, Nothing)

        If Atmosphere <> 0 AndAlso moColorSphere Is Nothing = False Then
            'now, render our sphere first

            matTemp.Translate(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ)
            matWorld.Multiply(matTemp)
            moDevice.Transform.World = matWorld

            Dim lPreSrcBlnd As Blend = moDevice.RenderState.SourceBlend
            Dim lPreDstBlnd As Blend = moDevice.RenderState.DestinationBlend
            Dim lPreBlendOp As BlendOperation = moDevice.RenderState.BlendOperation

            moDevice.RenderState.BlendOperation = BlendOperation.Max
            moDevice.RenderState.SourceBlend = Blend.One
            moDevice.RenderState.DestinationBlend = Blend.DestinationColor

            'now, draw our sphere
            moDevice.RenderState.CullMode = Cull.None
            moDevice.RenderState.Lighting = False
            moDevice.RenderState.ZBufferWriteEnable = False
            moColorSphere.DrawSubset(0)
            moDevice.RenderState.ZBufferWriteEnable = True
            moDevice.RenderState.Lighting = True
            moDevice.RenderState.CullMode = Cull.CounterClockwise

            moDevice.RenderState.BlendOperation = lPreBlendOp
            moDevice.RenderState.SourceBlend = lPreSrcBlnd
            moDevice.RenderState.DestinationBlend = lPreDstBlnd
        End If

		With moDevice
			.RenderState.FogColor = lColor
			.RenderState.FogVertexMode = FogMode.Linear

			'If bUseNewFogDistance = True Then
			'	.RenderState.FogStart = CSng(muSettings.EntityClipPlane * 0.66F)
			'	.RenderState.FogEnd = Math.Max(muSettings.TerrainFarClippingPlane, muSettings.EntityClipPlane)
			'Else
			.RenderState.FogStart = CSng(muSettings.TerrainFarClippingPlane * 0.66)
			.RenderState.FogEnd = muSettings.TerrainFarClippingPlane
			'End If

			.RenderState.FogEnable = True
		End With

		With moDevice
			'Render the skybox (stars, distant nebulae, etc...)
			If mbSkyboxReady Then
				.RenderState.FogEnable = False

				.SetTexture(0, Nothing)
				.Material = moSkyboxMat
				'set up our renderstates
				.RenderState.Lighting = False

				'if modevice.

				'Ok, if our device was created with mixed...
				If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
					'Set us up for software vertex processing as point sprites always work in software
					moDevice.SoftwareVertexProcessing = True
				End If

				.RenderState.PointSpriteEnable = True

				'.RenderState.PointSize = 2
				.RenderState.PointSize = 0.8

				'render our points
				Dim omatTemp As Matrix = Matrix.Identity
				Dim omatTrans As Matrix = Matrix.Translation(goCamera.mlCameraAtX, goCamera.mlCameraAtY, goCamera.mlCameraAtZ)
				omatTemp.Multiply(omatTrans)
				.Transform.World = omatTemp
				omatTemp = Nothing : omatTrans = Nothing

				.SetStreamSource(0, mvbSkybox, 0)
				.VertexFormat = CustomVertex.PositionColored.Format
				.DrawPrimitives(PrimitiveType.PointList, 0, muSettings.StarfieldParticlesPlanet)
				'now reset our renderstates
				.RenderState.PointSpriteEnable = False

				'now, reset us to hardware if we are in mixed
				If moDevice.CreationParameters.Behavior.MixedVertexProcessing = True Then
					moDevice.SoftwareVertexProcessing = False
				End If

				.RenderState.Lighting = True

				.RenderState.FogEnable = True
			Else
				CreateSkybox()
			End If
		End With

		moDevice.RenderState.FogEnable = False
		RenderClouds(200000, lColor)
		moDevice.RenderState.FogEnable = True

	End Sub

	Private Sub RenderOrbitalBodies()

		'TODO: Render moons of me... and moons of my parent... and my parent

		'moDevice.RenderState.ZBufferEnable = False

		'Dim matWorld As Matrix
		'Dim dAngle As Single
		'Dim matYaw As Matrix
		'Dim matAngle As Matrix

		''Now, render this planet as a moon (for grins and giggles)
		'If moTexture Is Nothing Then
		'    moTerrain.CreateTerrainPlanetTexture(moTexture)
		'    With moMaterial
		'        .Diffuse = System.Drawing.Color.White
		'        .Ambient = System.Drawing.Color.FromArgb(255, 18, 18, 18)
		'        .Emissive = System.Drawing.Color.Black
		'        .Specular = System.Drawing.Color.Gray
		'        .SpecularSharpness = 10
		'    End With
		'End If

		''Rotate Yaw (Axis)
		'dAngle = ((AxisAngle) / 10.0F) * CSng(Math.PI / 180.0F)
		'matYaw = Matrix.Identity
		'matYaw.RotateX(CSng(dAngle))

		''Rotate Angle (rotateangle)
		'matAngle = Matrix.Identity
		'dAngle = (RotateAngle / 10.0F) * CSng(Math.PI / 180.0F)
		'matAngle.RotateY(CSng(dAngle))

		''Create our world
		'matWorld = Matrix.Multiply(matYaw, matAngle)
		'matYaw = Nothing
		'matAngle = Nothing

		'matWorld.Multiply(Matrix.Translation(0, 320 * moTerrain.ml_Y_Mult, goCamera.mlCameraAtZ + 50000))


		''set the world and render
		'moDevice.Transform.World = matWorld
		'moDevice.Material = moMaterial
		'moDevice.SetTexture(0, moTexture)

		''Make sure our planet exists...
		'If moMesh Is Nothing Then moMesh = goResMgr.CreateTexturedSphere(miPlanetRadius, 32, 32, 0)
		'moDevice.RenderState.Wrap0 = WrapCoordinates.Zero

		'moMesh.DrawSubset(0)

		''Now, do we have a ring
		'If mbHasRing = True Then
		'    moDevice.RenderState.CullMode = Cull.None
		'    moDevice.SetTexture(0, moRingTex)
		'    moDevice.Material = moRingMat
		'    moRing.DrawSubset(0)
		'    moDevice.RenderState.CullMode = Cull.CounterClockwise
		'End If

		'moDevice.RenderState.Wrap0 = 0 ' WrapCoordinates.Two

		'matWorld = Nothing

		'moDevice.RenderState.ZBufferEnable = True
	End Sub

    Public Sub CleanResources(ByVal bComplete As Boolean)
        If moTerrain Is Nothing = False Then moTerrain.CleanResources()

        'if complete clean, then we clean all but the basic data of the planet...
        If bComplete Then
            'since planet meshes are all scraps...
            If moMesh Is Nothing = False Then moMesh.Dispose()
            moMesh = Nothing
            If moTexture Is Nothing = False Then moTexture.Dispose() 'we call dispose here because the texture is a scrap
            moTexture = Nothing
            If moLightMap Is Nothing = False Then moLightMap.Dispose()
            moLightMap = Nothing
            If moNormalMap Is Nothing = False Then moNormalMap.Dispose()
            moNormalMap = Nothing
            If moRing Is Nothing = False Then moRing.Dispose()
            moRing = Nothing
            mbHasRing = False   'no longer has the ring...
            'moRingTex = Nothing		'do not call dispose here because this is not a scrap
            If moCloud Is Nothing = False Then moCloud.Dispose()
            moCloud = Nothing
            Erase miSpeedMod

            'If another planet needs it, it will recreate it
            If moStarSprite Is Nothing = False Then moStarSprite.Dispose()
            moStarSprite = Nothing
        End If
    End Sub

	Protected Overrides Sub Finalize()
		'Note: The only time we would call this is if we were COMPLETELY done with the planet... otherwise, call CleanResources
		CleanResources(True)
		moTerrain = Nothing
		If moMesh Is Nothing = False Then moMesh.Dispose()
		moMesh = Nothing
		If moRing Is Nothing = False Then moRing.Dispose()
		moRing = Nothing

		'For textures, we set them to nothing, but we do not call DISPOSE!!! it is the Resource Manager's responsibility
		'moRingTex = Nothing

		ParentSystem = Nothing	'remove our reference when setting to nothing

		moSprite = Nothing	'DO NOT DISPOSE!!!
        MyBase.Finalize()
	End Sub

	Public Function GetMapClickedCoords(ByVal lX As Int32, ByVal lY As Int32) As Point
		Dim lSmaller As Int32
		Dim lLarger As Int32
		Dim ptPos As Point

        Dim moDevice As Device = GFXEngine.moDevice

		'Ok, let's get the rectangle for where the map is actually being rendered...
		Dim lTempW As Int32 = moDevice.PresentationParameters.BackBufferWidth
		Dim lTempH As Int32 = moDevice.PresentationParameters.BackBufferHeight
		lSmaller = CInt(Math.Min(lTempW, lTempH) * 0.8)
		lLarger = Math.Max(lTempW, lTempH)
		ptPos.X = CInt((lTempW - lSmaller) / 2)
		ptPos.Y = CInt((lTempH - lSmaller) / 2)

		Dim fMultX As Single = CSng(TerrainClass.Width / lSmaller)
		Dim fMultY As Single = CSng(TerrainClass.Height / lSmaller)

		'Now get our display settings...
		lSmaller = CInt(Math.Min(moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight) * 0.8F)
		lLarger = Math.Max(moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight)
		ptPos.X = CInt((moDevice.PresentationParameters.BackBufferWidth - lSmaller) / 2)
		ptPos.Y = CInt((moDevice.PresentationParameters.BackBufferHeight - lSmaller) / 2)

		If lX >= ptPos.X And lX < ptPos.X + lSmaller Then
			If lY >= ptPos.Y And lY < ptPos.Y + lSmaller Then
				'ok, we're inside, the point should reflect our map...
				ptPos.X = (lX - ptPos.X)
				ptPos.Y = (lY - ptPos.Y)

				'ptPos now has our MAP locational coordinate, but this coordinate is transformed from the stretch...
				ptPos.X = CInt((TerrainClass.Width * ptPos.X) / lSmaller)
				ptPos.Y = CInt((TerrainClass.Height * ptPos.Y) / lSmaller)

				'Now, ptPos tells us where on the map they clicked in Vertex numbers, so multiply by cell spacing
				ptPos.X *= moTerrain.CellSpacing
				ptPos.Y *= moTerrain.CellSpacing

				'But then, reduce that by half the map width to offset center of map
				ptPos.X -= CInt((TerrainClass.Width * moTerrain.CellSpacing) / 2)
				ptPos.Y -= CInt((TerrainClass.Height * moTerrain.CellSpacing) / 2)

				'and finally, EVERYTHING on the map is mirrored on the X axis
				ptPos.X *= -1
			Else
				ptPos.X = Integer.MinValue
				ptPos.Y = Integer.MinValue
			End If
		Else
			ptPos.X = Integer.MinValue
			ptPos.Y = Integer.MinValue
		End If
		Return ptPos
	End Function

	Public Function GetHeightAtPoint(ByVal lX As Single, ByVal lZ As Single, ByVal bWaterHeightMin As Boolean) As Single
		Return moTerrain.GetHeightAtLocation(lX, lZ, bWaterHeightMin)
	End Function

	Public Function GetTriangleNormal(ByVal lX As Single, ByVal lZ As Single) As Vector3
		Return moTerrain.GetTerrainNormal(lX, lZ)
	End Function

	Public Function MiniMapClick(ByVal lX As Int32, ByVal lY As Int32) As Vector3
		'Ok, get the single mults
		'dim fXVal as Single = lx 
		Dim lCX As Int32 = 0
		Dim lCZ As Int32 = 0
		Dim lSize As Int32 = 255

		If muSettings.PlanetMinimapZoomLevel = 0 Then
			'Zoomed all the way out, so Point is point
			lSize = 255
		ElseIf muSettings.PlanetMinimapZoomLevel = 1 Then
			lSize = TerrainClass.Width \ 2
		ElseIf muSettings.PlanetMinimapZoomLevel = 2 Then
			lSize = TerrainClass.Width \ 4
		Else
			lSize = TerrainClass.Width \ 6
		End If
		Dim lHalfSize As Int32 = lSize \ 2

		'Get the center of the minimap picture...
		lCX = CInt(goCamera.mlCameraX / moTerrain.CellSpacing) + (TerrainClass.Width \ 2)
		lCZ = CInt(goCamera.mlCameraZ / moTerrain.CellSpacing) + (TerrainClass.Height \ 2)

		If lCX < lHalfSize Then lCX = lHalfSize
		If lCZ < lHalfSize Then lCZ = lHalfSize
		If lCX + lHalfSize > 255 Then lCX = 255 - lHalfSize
		If lCZ + lHalfSize > 255 Then lCZ = 255 - lHalfSize

		'from that and the size, we can get the upper left corner's loc
		Dim lCornerX As Int32 = lCX - lHalfSize
		Dim lCornerY As Int32 = lCZ - lHalfSize

		'now, get the multiplier of the ratio of X and Y over the Width/Height of the minimap
		Dim fMultX As Single = CSng((lX - muSettings.MiniMapLocX) / muSettings.MiniMapWidthHeight)
		Dim fMultY As Single = CSng((lY - muSettings.MiniMapLocY) / muSettings.MiniMapWidthHeight)

		'Take the multiplier and multiply by the size, this tells us the size from corner... so add the corner
		lCX = CInt(lCornerX + (lSize * (1 - fMultX)))
		lCZ = CInt(lCornerY + (lSize * fMultY))

		'That is our VERT clicked on... to get the translated to Object Space position, subtract half the width/height of the terrain
		lCX -= (TerrainClass.Width \ 2)
		lCZ -= (TerrainClass.Height \ 2)

		'And multiply it by the cellspacing
		lCX *= moTerrain.CellSpacing
		lCZ *= moTerrain.CellSpacing

		Return New Vector3(lCX, GetHeightAtPoint(lCX, lCZ, True), lCZ)
	End Function

	Public Function GetMaxHeight() As Single
		'the max heightmap value is 256
		Return moTerrain.ml_Y_Mult * 256
	End Function

    Private Sub CreateColorSphere()
        Dim moDevice As Device = GFXEngine.moDevice
        Device.IsUsingEventHandlers = False
        Dim oTemp As Mesh = Mesh.Sphere(moDevice, 20000, 32, 32)

        moColorSphere = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionColored.Format, moDevice)

        ' Get the original mesh's vertex buffer.
        Dim ranks(0) As Integer
        ranks(0) = oTemp.NumberVertices
        Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

        ' Set the vertex buffer
        Dim data As System.Array = moColorSphere.VertexBuffer.Lock(0, (New CustomVertex.PositionColored()).GetType(), LockFlags.None, ranks)

        'Dim lR As Int32
        'Dim lG As Int32
        'Dim lB As Int32

        'Dim phi As Single
        Dim u As Single
        Dim v As Single
        Dim i As Integer

        ReDim mfSphereU(arr.Length - 1)
        ReDim mfSphereV(arr.Length - 1)

        For i = 0 To arr.Length - 1
            Dim pn As Direct3D.CustomVertex.PositionNormal = CType(arr.GetValue(i), CustomVertex.PositionNormal)
            Dim pnt As Direct3D.CustomVertex.PositionColored = CType(data.GetValue(i), CustomVertex.PositionColored)

            pnt.X = pn.X
            pnt.Y = pn.Y
            pnt.Z = pn.Z

            u = CSng(Math.Asin(pn.Nx) / Math.PI + 0.5)
            v = CSng(Math.Asin(pn.Ny) / Math.PI + 0.5)

            mfSphereU(i) = u
            mfSphereV(i) = v

            pnt.Color = System.Drawing.Color.FromArgb(255, 0, 0, 0).ToArgb

            data.SetValue(pnt, i)
        Next i

        moColorSphere.VertexBuffer.Unlock()
        oTemp.VertexBuffer.Unlock()

        ' Set the index buffer. 
        ranks(0) = oTemp.NumberFaces * 3
        arr = oTemp.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
        moColorSphere.IndexBuffer.SetData(arr, 0, LockFlags.None)

        Device.IsUsingEventHandlers = True
        oTemp = Nothing
        arr = Nothing
        data = Nothing

    End Sub

	Private Function ColorSphereSetColor() As System.Drawing.Color

        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return Color.Black

		If moColorSphere Is Nothing Then CreateColorSphere()

        Dim X As Int32
        
		Dim u As Single
		Dim v As Single

		Dim lR As Int32
		Dim lG As Int32
		Dim lB As Int32

		Dim fEastR As Single
		Dim fEastG As Single
		Dim fEastB As Single
		Dim fWestR As Single
		Dim fWestG As Single
		Dim fWestB As Single

		Dim fFinalEastR As Single = 0
		Dim fFinalEastG As Single = 0
		Dim fFinalEastB As Single = 0
		Dim fFinalWestR As Single = 0
		Dim fFinalWestG As Single = 0
		Dim fFinalWestB As Single = 0
		Dim lStarID As Int32
		'Dim lCurStarTypeID As Int32
		Dim fAngle As Single
		Dim lStar1Radius As Int32
		Dim lSunX As Int32

		'Dim fDayTimeRatio As Single '= RotateAngle / 3600
		Dim fTemp As Single

		Dim oResultColor As System.Drawing.Color

		'Need to take into account multiple stars in a system
		For lStarID = 1 To 3
            If (lStarID = 1 AndAlso ParentSystem.StarType1Idx > -1) OrElse _
               (lStarID = 2 AndAlso ParentSystem.StarType2Idx > -1) OrElse _
               (lStarID = 3 AndAlso ParentSystem.StarType3Idx > -1) Then

                Select Case lStarID
                    Case 1
                        lStar1Radius = goStarTypes(ParentSystem.StarType1Idx).StarRadius
                        fAngle = LineAngleDegrees(0, 0, LocX, LocZ) * 10
                    Case 2
                        lSunX = lStar1Radius + goStarTypes(ParentSystem.StarType2Idx).StarRadius + 1000
                        fAngle = LineAngleDegrees(lSunX, 0, LocX, LocZ) * 10
                    Case 3
                        lSunX = -(lStar1Radius + goStarTypes(ParentSystem.StarType3Idx).StarRadius + 1000)
                        fAngle = LineAngleDegrees(lSunX, -lSunX, LocX, LocZ) * 10
                End Select
                fAngle = fAngle + CSng(RotateAngle) - 150
                If fAngle > 3600 Then fAngle -= 3600
                If fAngle < 0 Then fAngle += 3600

                fDayTimeRatio = fAngle / 3600

                Select Case Me.MapTypeID
                    Case PlanetType.eAcidic
                        Select Case fDayTimeRatio
                            Case Is < 0.2
                                fEastR = 0 : fWestR = 0
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.25
                                fTemp = fDayTimeRatio - 0.2F    'result is .000 to .05
                                fTemp /= 0.05F                   'result is 0 to 1
                                fEastR = 0 : fWestR = 0
                                fEastG = fTemp * 32 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.3
                                fTemp = fDayTimeRatio - 0.25F    'result is .000 to .05
                                fTemp /= 0.05F                   'result is 0 to 1
                                fEastR = 0 : fWestR = 0
                                fEastG = (fTemp * 16) + 32 : fWestG = fTemp * 32
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.333333
                                fTemp = fDayTimeRatio - 0.3F     'result is .000 to .033333
                                fTemp /= 0.033333F               'result is 0 to 1
                                fEastR = fTemp * 255 : fWestR = 0
                                fEastG = 32 + (fTemp * 220) : fWestG = 32
                                fEastB = fTemp * 64 : fWestB = 0
                            Case Is < 0.375
                                fTemp = fDayTimeRatio - 0.333333F 'result is .000 to .041667
                                fTemp /= 0.041667F               'result is 0 to 1
                                fEastR = 255 - (fTemp * 192) : fWestR = fTemp * 64
                                fEastG = 255 - (fTemp * 64) : fWestG = 32 + (fTemp * 192)
                                fEastB = 64 : fWestB = fTemp * 64
                            Case Is < 0.41667
                                fTemp = fDayTimeRatio - 0.375F   'result is .000 to .041667
                                fTemp /= 0.041667F               'result is 0 to 1
                                fEastR = (64 * fTemp) + 64 : fWestR = fEastR
                                fEastG = 191 + (fTemp * 64) : fWestG = fEastG
                                fEastB = (64 * fTemp) + 64 : fWestB = fEastB
                            Case Is < 0.625
                                fEastR = 128 : fWestR = fEastR
                                fEastG = 255 : fWestG = fEastG
                                fEastB = 128 : fWestB = fEastB
                            Case Is < 0.666667
                                fTemp = fDayTimeRatio - 0.625F   'result is .000 to .041667
                                fTemp /= 0.041667F
                                fEastR = 128 - (64 * fTemp) : fWestR = fEastR
                                fEastG = 255 : fWestG = 255 - (64 * fTemp)
                                fEastB = 128 - (64 * fTemp) : fWestB = fEastB
                            Case Is < 0.708333
                                fTemp = fDayTimeRatio - 0.666667F 'result is .000 to 
                                fTemp /= 0.041667F
                                fEastR = 64 : fWestR = (fTemp * 191) + 64
                                fEastG = (1 - fTemp) * 255 : fWestG = 192
                                fEastB = (1 - fTemp) * 64 : fWestB = fEastB
                            Case Is < 0.75
                                fTemp = fDayTimeRatio - 0.708333F
                                fTemp /= 0.041667F
                                fEastR = 64 - (fTemp * 64) : fWestR = 255
                                fEastG = fTemp * 32 : fWestG = 192 '- (fEastG)
                                fEastB = 0 : fWestB = 0 '(1 - fTemp) * 128
                            Case Is < 0.791667
                                fTemp = fDayTimeRatio - 0.75F
                                fTemp /= 0.041667F
                                fEastR = 0 : fWestR = 255 - (192 * fTemp)
                                fEastG = 32 : fWestG = 192 - (64 * fTemp)
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.833333
                                fTemp = fDayTimeRatio - 0.791667F
                                fTemp /= 0.041667F
                                fEastR = 0 : fWestR = (1 - fTemp) * 64
                                fEastG = 32 : fWestG = 128 - (fTemp * 80)
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.875
                                fTemp = fDayTimeRatio - 0.833333F
                                fTemp /= 0.041667F
                                fEastR = 0 : fWestR = 0
                                fEastG = (1 - fTemp) * 32 : fWestG = (1 - fTemp) * 48
                                fEastB = 0 : fWestB = 0
                            Case Else
                                fEastR = 0 : fWestR = 0
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                        End Select
                    Case PlanetType.eGeoPlastic
                        Select Case fDayTimeRatio
                            Case Is < 0.2
                                fEastR = 64 : fWestR = 64
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.25
                                fTemp = fDayTimeRatio - 0.2F    'result is .000 to .05
                                fTemp /= 0.05F                   'result is 0 to 1
                                fEastR = 64 + (fTemp * 16) : fWestR = 64
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.3
                                fTemp = fDayTimeRatio - 0.25F    'result is .000 to .05
                                fTemp /= 0.05F                   'result is 0 to 1
                                fEastR = 80 + (fTemp * 16) : fWestR = 64 + (fTemp * 32)
                                fEastG = (fTemp * 32) : fWestG = 0
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.333333
                                fTemp = fDayTimeRatio - 0.3F     'result is .000 to .033333
                                fTemp /= 0.033333F               'result is 0 to 1
                                fEastR = 96 + (fTemp * 64) : fWestR = 96 + (fTemp * 64)
                                fEastG = 32 + (fTemp * 96) : fWestG = fTemp * 32
                                fEastB = fTemp * 32 : fWestB = fTemp * 32
                            Case Is < 0.375
                                fTemp = fDayTimeRatio - 0.333333F 'result is .000 to .041667
                                fTemp /= 0.041667F               'result is 0 to 1
                                fEastR = 160 + (fTemp * 32) : fWestR = fEastR
                                fEastG = 128 - (fTemp * 64) : fWestG = 32 + (fTemp * 32)
                                fEastB = 32 + (fTemp * 16) : fWestB = 32 + (fTemp * 16)
                            Case Is < 0.41667
                                fTemp = fDayTimeRatio - 0.375F   'result is .000 to .041667
                                fTemp /= 0.041667F               'result is 0 to 1
                                fEastR = 192 + (fTemp * 63) : fWestR = fEastR
                                fEastG = 64 + (fTemp * 64) : fWestG = fEastG
                                fEastB = 48 + (fTemp * 64) : fWestB = fEastB
                            Case Is < 0.625
                                fEastR = 255 : fWestR = fEastR
                                fEastG = 128 : fWestG = fEastG
                                fEastB = 112 : fWestB = fEastB
                            Case Is < 0.666667
                                fTemp = fDayTimeRatio - 0.625F   'result is .000 to .041667
                                fTemp /= 0.041667F
                                fEastR = 255 - (fTemp * 64) : fWestR = fEastR
                                fEastG = 128 - (fTemp * 32) : fWestG = 128
                                fEastB = 112 - (fTemp * 16) : fWestB = fEastB
                            Case Is < 0.708333
                                fTemp = fDayTimeRatio - 0.666667F 'result is .000 to 
                                fTemp /= 0.041667F
                                fEastR = 192 - (fTemp * 32) : fWestR = 192
                                fEastG = 96 - (fTemp * 32) : fWestG = 128
                                fEastB = 96 - (fTemp * 32) : fWestB = fEastB
                            Case Is < 0.75
                                fTemp = fDayTimeRatio - 0.708333F
                                fTemp /= 0.041667F
                                fEastR = 160 - (fTemp * 64) : fWestR = 192 - (fTemp * 32)
                                fEastG = 64 - (fTemp * 32) : fWestG = 128 - (fTemp * 32)
                                fEastB = 64 - (fTemp * 32) : fWestB = 64
                            Case Is < 0.791667
                                fTemp = fDayTimeRatio - 0.75F
                                fTemp /= 0.041667F
                                fEastR = 96 : fWestR = 160 - (64 * fTemp)
                                fEastG = 32 : fWestG = 96 - (fTemp * 64)
                                fEastB = 32 : fWestB = 64 - (fTemp * 32)
                            Case Is < 0.833333
                                fTemp = fDayTimeRatio - 0.791667F
                                fTemp /= 0.041667F
                                fEastR = 96 - (fTemp * 32) : fWestR = fEastR
                                fEastG = 32 - (fTemp * 32) : fWestG = fEastG
                                fEastB = 32 - (fTemp * 32) : fWestB = 32 - (fTemp * 32)
                            Case Is < 0.875
                                fTemp = fDayTimeRatio - 0.833333F
                                fTemp /= 0.041667F
                                fEastR = 64 : fWestR = 64
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                            Case Else
                                fEastR = 64 : fWestR = 64
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                        End Select

                    Case Else       'TODO: Expand on this further when the other atmosphere colors are done
                        Select Case fDayTimeRatio
                            Case Is < 0.2
                                fEastR = 0 : fWestR = 0
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                            Case Is < 0.25
                                fTemp = fDayTimeRatio - 0.2F    'result is .000 to .05
                                fTemp /= 0.05F                   'result is 0 to 1
                                fEastR = 0 : fWestR = 0
                                fEastG = 0 : fWestG = 0
                                fEastB = fTemp * 32 : fWestB = 0
                            Case Is < 0.3
                                fTemp = fDayTimeRatio - 0.25F    'result is .000 to .05
                                fTemp /= 0.05F                   'result is 0 to 1
                                fEastR = 0 : fWestR = 0
                                fEastG = 0 : fWestG = 0
                                fEastB = (fTemp * 16) + 32 : fWestB = fTemp * 32
                            Case Is < 0.333333
                                fTemp = fDayTimeRatio - 0.3F     'result is .000 to .033333
                                fTemp /= 0.033333F               'result is 0 to 1
                                fEastR = fTemp * 255 : fWestR = 0
                                fEastG = fTemp * 128 : fWestG = 0
                                fEastB = (1 - fTemp) * 48 : fWestB = (fTemp * 16) + 32
                            Case Is < 0.375
                                fTemp = fDayTimeRatio - 0.333333F 'result is .000 to .041667
                                fTemp /= 0.041667F               'result is 0 to 1
                                fEastR = 255 - (fTemp * 192) : fWestR = fTemp * 64
                                fEastG = 128 : fWestG = fTemp * 128
                                fEastB = fTemp * 255 : fWestB = (fTemp * 206) + 48
                            Case Is < 0.41667
                                fTemp = fDayTimeRatio - 0.375F   'result is .000 to .041667
                                fTemp /= 0.041667F              'result is 0 to 1
                                fEastR = (64 * fTemp) + 64 : fWestR = fEastR
                                fEastG = (64 * fTemp) + 128 : fWestG = fEastG
                                fEastB = 255 : fWestB = 255
                            Case Is < 0.625
                                fEastR = 128 : fWestR = fEastR
                                fEastG = 192 : fWestG = fEastG
                                fEastB = 255 : fWestB = fEastB
                            Case Is < 0.666667
                                fTemp = fDayTimeRatio - 0.625F   'result is .000 to .041667
                                fTemp /= 0.041667F
                                fEastR = 128 - (64 * fTemp) : fWestR = fEastR
                                fEastG = 192 - (64 * fTemp) : fWestG = 128 + (64 * fTemp)
                                fEastB = 255 : fWestB = fEastB
                            Case Is < 0.708333
                                fTemp = fDayTimeRatio - 0.666667F 'result is .000 to 
                                fTemp /= 0.041667F
                                fEastR = 64 : fWestR = (fTemp * 191) + 64
                                fEastG = (1 - fTemp) * 128 : fWestG = 192
                                fEastB = (1 - fTemp) * 255 : fWestB = 255 - (128 * fTemp)
                            Case Is < 0.75
                                fTemp = fDayTimeRatio - 0.708333F
                                fTemp /= 0.041667F
                                fEastR = 64 + (fTemp * 128) : fWestR = 255
                                fEastG = fTemp * 64 : fWestG = 192 - (fEastG)
                                fEastB = fTemp * 192 : fWestB = (1 - fTemp) * 128
                            Case Is < 0.791667
                                fTemp = fDayTimeRatio - 0.75F
                                fTemp /= 0.041667F
                                fEastR = (1 - fTemp) * 192 : fWestR = 255 - (64 * fTemp)
                                fEastG = (1 - fTemp) * 64 : fWestG = 128 - (64 * fTemp)
                                fEastB = 192 - (fTemp * 144) : fWestB = fTemp * 192
                            Case Is < 0.833333
                                fTemp = fDayTimeRatio - 0.791667F
                                fTemp /= 0.041667F
                                fEastR = 0 : fWestR = (1 - fTemp) * 192
                                fEastG = 0 : fWestG = (1 - fTemp) * 64
                                fEastB = 48 - (fTemp * 16) : fWestB = 192 - (fTemp * 144)
                            Case Is < 0.875
                                fTemp = fDayTimeRatio - 0.833333F
                                fTemp /= 0.041667F
                                fEastR = 0 : fWestR = 0
                                fEastG = 0 : fWestG = 0
                                fEastB = (1 - fTemp) * 32 : fWestB = (1 - fTemp) * 48
                            Case Else
                                fEastR = 0 : fWestR = 0
                                fEastG = 0 : fWestG = 0
                                fEastB = 0 : fWestB = 0
                        End Select
                End Select

                fFinalEastR = Math.Max(fEastR, fFinalEastR)
                fFinalEastG = Math.Max(fEastG, fFinalEastG)
                fFinalEastB = Math.Max(fEastB, fFinalEastB)
                fFinalWestR = Math.Max(fWestR, fFinalWestR)
                fFinalWestG = Math.Max(fWestG, fFinalWestG)
                fFinalWestB = Math.Max(fWestB, fFinalWestB)
            End If
		Next lStarID

        Dim vecOverallLight As Vector3 = New Vector3(0, 0, 0)
        Try
            Dim oDev As Device = GFXEngine.moDevice
            For X = 0 To oDev.Lights.Count - 1
                If oDev.Lights(X).Enabled = True Then
                    vecOverallLight.X = Math.Max(vecOverallLight.X, oDev.Lights(X).Diffuse.R / 255.0F)
                    vecOverallLight.Y = Math.Max(vecOverallLight.Y, oDev.Lights(X).Diffuse.G / 255.0F)
                    vecOverallLight.Z = Math.Max(vecOverallLight.Z, oDev.Lights(X).Diffuse.B / 255.0F)
                End If
            Next X
        Catch
        End Try

        fEastR = (fFinalEastR * vecOverallLight.X)
        fEastG = (fFinalEastG * vecOverallLight.Y)
        fEastB = (fFinalEastB * vecOverallLight.Z)
        fWestR = (fFinalWestR * vecOverallLight.X)
        fWestG = (fFinalWestG * vecOverallLight.Y)
        fWestB = (fFinalWestB * vecOverallLight.Z)

        oResultColor = System.Drawing.Color.FromArgb(255, CInt((fEastR + fWestR) / 2), CInt((fEastG + fWestG) / 2), CInt((fEastB + fWestB) / 2))

        'now, determine if we need to update our sphere
        If myEastR = Math.Floor(fEastR) AndAlso myEastG = Math.Floor(fEastG) AndAlso _
          myEastB = Math.Floor(fEastB) AndAlso myWestR = Math.Floor(fWestR) AndAlso _
          myWestG = Math.Floor(fWestG) AndAlso myWestB = Math.Floor(fWestB) Then Return oResultColor 'Exit Function

        Dim ranks(0) As Integer
        ranks(0) = moColorSphere.NumberVertices

        ' Set the vertex buffer
        Dim data As System.Array '= moColorSphere.VertexBuffer.Lock(0, (New CustomVertex.PositionColored()).GetType(), LockFlags.None, ranks)
        Try
            data = moColorSphere.VertexBuffer.Lock(0, (New CustomVertex.PositionColored()).GetType(), LockFlags.None, ranks)
        Catch
            moColorSphere = Nothing
            Return Color.Black
        End Try

        myEastR = CByte(fEastR)
        myEastG = CByte(fEastG)
        myEastB = CByte(fEastB)
        myWestR = CByte(fWestR)
        myWestG = CByte(fWestG)
        myWestB = CByte(fWestB)

        Dim fDR As Single = fWestR - fEastR
        Dim fDG As Single = fWestG - fEastG
        Dim fDB As Single = fWestB - fEastB

        For X = 0 To data.Length - 1
            Dim pnt As Direct3D.CustomVertex.PositionColored = CType(data.GetValue(X), CustomVertex.PositionColored)

            u = mfSphereU(X)
            v = mfSphereV(X)

            lR = CInt((fDR * u) + fEastR)
            lG = CInt((fDG * u) + fEastG)
            lB = CInt((fDB * u) + fEastB)

            pnt.Color = System.Drawing.Color.FromArgb(255, lR, lG, lB).ToArgb

            data.SetValue(pnt, X)
        Next X

        moColorSphere.VertexBuffer.Unlock()

        data = Nothing

        Return oResultColor
    End Function

	Private Function FillSpeedModArray() As Boolean
		'Call this when terrain is ready...

		If moTerrain.NormalsReady = False Then Return False

		ReDim miSpeedMod((TerrainClass.Width * 2) - 1, (TerrainClass.Height * 2) - 1, 7)
		'Ok, now, fill it up
		Dim X As Int32
		Dim Y As Int32
		Dim lHalfCell As Int32 = CInt(moTerrain.CellSpacing / 2)

		Dim fTX As Single
		Dim fTY As Single

		Dim vTemp As Vector3

		mlFullSizeWH = moTerrain.CellSpacing * TerrainClass.Width	   'NOTE: we assume Width = Height here
		mlDoubleWH = TerrainClass.Width * 2							   'NOTE: we assume Width = Height here

		'For each quad...
		For Y = 0 To (TerrainClass.Height * 2) - 1
			fTY = CSng(Y / 2)
			For X = 0 To (TerrainClass.Width * 2) - 1
				fTX = CSng(X / 2)

				'Ok, get our value
				vTemp = moTerrain.GetTerrainNormalEx(fTX, fTY)

				''Right
				'miSpeedMod(X, Y, 0) = (1 + vTemp.X) * 100
				''Up Right
				'miSpeedMod(X, Y, 1) = (1 + (vTemp.Z + vTemp.X)) * 100
				''Up
				'miSpeedMod(X, Y, 2) = (1 + vTemp.Z) * 100
				''Up Left
				'miSpeedMod(X, Y, 3) = (1 + (vTemp.Z + -vTemp.X)) * 100
				''Left
				'miSpeedMod(X, Y, 4) = (1 + (-vTemp.X)) * 100
				''Down
				'miSpeedMod(X, Y, 6) = (1 + (-vTemp.Z)) * 100
				''Down Left
				'miSpeedMod(X, Y, 5) = (1 + (-vTemp.Z + -vTemp.X)) * 100
				''Down Right
				'miSpeedMod(X, Y, 7) = (1 + (-vTemp.Z + vTemp.X)) * 100
				'Right
				miSpeedMod(X, Y, 0) = CShort((1 + vTemp.X) * 100)
				'Up Right
				miSpeedMod(X, Y, 7) = CShort((1 + (vTemp.Z + vTemp.X)) * 100)
				'Up
				miSpeedMod(X, Y, 6) = CShort((1 + vTemp.Z) * 100)
				'Up Left
				miSpeedMod(X, Y, 5) = CShort((1 + (vTemp.Z + -vTemp.X)) * 100)
				'Left
				miSpeedMod(X, Y, 4) = CShort((1 + (-vTemp.X)) * 100)
				'Down Left
				miSpeedMod(X, Y, 3) = CShort((1 + (-vTemp.Z + -vTemp.X)) * 100)
				'Down
				miSpeedMod(X, Y, 2) = CShort((1 + (-vTemp.Z)) * 100)
				'Down Right
				miSpeedMod(X, Y, 1) = CShort((1 + (-vTemp.Z + vTemp.X)) * 100)
			Next X
		Next Y

		Return True
	End Function

	Public Function GetSpeedMod(ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iAngle As Int16) As Single
        'If lLocX < goCurrentEnvir.lMinXPos Then lLocX += goCurrentEnvir.lMapWrapAdjustX
        'If lLocX > goCurrentEnvir.lMaxXPos Then lLocX -= goCurrentEnvir.lMapWrapAdjustX

		Dim lTX As Int32 = CInt((2 * lLocX + mlFullSizeWH) / moTerrain.CellSpacing)
		Dim lTZ As Int32 = CInt((2 * lLocZ + mlFullSizeWH) / moTerrain.CellSpacing)
		Dim lTA As Int32

        If mbSpeedModRdy = False OrElse miSpeedMod Is Nothing Then
            mbSpeedModRdy = FillSpeedModArray()
            If mbSpeedModRdy = False Then Return 1.0F
        End If

		'Ensure angle is in valid range
		If iAngle < 0S Then
			iAngle += 3600S
		ElseIf iAngle > 3600S Then
			iAngle -= 3600S
		End If
		lTA = CInt(Math.Floor(iAngle / 450I))

		If (lTX < 0 OrElse lTX > mlDoubleWH - 1) OrElse (lTZ < 0 OrElse lTZ > mlDoubleWH - 1) Then
			'Range is outside bounds...
			Return 1.0F
		Else
			Return (miSpeedMod(lTX, lTZ, lTA) / 100.0F)
		End If
	End Function

	Public Function GetExtent() As Int32
		Return moTerrain.CellSpacing * TerrainClass.Width
	End Function

	Public ReadOnly Property WaterHeight() As Int32
		Get
			Return CInt(moTerrain.WaterHeight * moTerrain.ml_Y_Mult)
		End Get
	End Property

	Public ReadOnly Property CellSpacing() As Int32
		Get
			Return moTerrain.CellSpacing
		End Get
	End Property

    Public Sub SaveToFile()
        moTerrain.SaveToFile(Me.ObjectID)
    End Sub

    Public Sub SaveTexToFile()
        SurfaceLoader.Save("C:\Diffuse.bmp", ImageFileFormat.Bmp, moTexture.GetSurfaceLevel(0))
        If moLightMap Is Nothing = False Then SurfaceLoader.Save("C:\Light.bmp", ImageFileFormat.Bmp, moLightMap.GetSurfaceLevel(0))
        If moNormalMap Is Nothing = False Then SurfaceLoader.Save("C:\Normal.bmp", ImageFileFormat.Bmp, moNormalMap.GetSurfaceLevel(0))
    End Sub

	Public Sub SetCityCreepLoc(ByVal fLocX As Single, ByVal fLocZ As Single, ByVal bSetValue As Boolean, ByVal fSize As Single)
        If HasAliasedRights(AliasingRights.eViewUnitsAndFacilities) = False Then Return
        If bSetValue = True Then moTerrain.AdjustCityCreepRef(fLocX, fLocZ, 1, fSize) Else moTerrain.AdjustCityCreepRef(fLocX, fLocZ, -1, fSize)
    End Sub

	Public Sub ClearPlanetTexture()
        If moTexture Is Nothing = False Then moTexture.Dispose()
        bSaved = False
		moTexture = Nothing
	End Sub

End Class
