Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Structure CullBox
    Public vecMin As Vector3
    Public vecMax As Vector3

    Public Shared Function GetCullBox(ByVal fLocX As Single, ByVal fLocY As Single, ByVal fLocZ As Single, ByVal fMinX As Single, ByVal fMinY As Single, ByVal fMinZ As Single, ByVal fMaxX As Single, ByVal fMaxY As Single, ByVal fMaxZ As Single) As CullBox
        Dim oBox As CullBox
        oBox.vecMax.X = fLocX + fMaxX
        oBox.vecMax.Y = fLocY + fMaxY
        oBox.vecMax.Z = fLocZ + fMaxZ
        oBox.vecMin.X = fLocX + fMinX
        oBox.vecMin.Y = fLocY + fMinY
        oBox.vecMin.Z = fLocZ + fMinZ
        Return oBox
    End Function
End Structure

Public Class TerrainClass

    Private Structure terVert
        Public p As Vector3
        Public n As Vector3
        Public tu As Single
        Public tv As Single
        Public tw As Single

        Public tu2 As Single
        Public tv2 As Single
        Public tw2 As Single
    End Structure
    Private terVertFmt As VertexFormats = VertexFormats.Position Or VertexFormats.Normal Or VertexFormats.Texture1 Or VertexFormats.Texture2 Or VertexTextureCoordinate.Size3(0) Or VertexTextureCoordinate.Size3(1)

#Region "  Normal and Illumination Rendering  "
    Private Structure terVertTBN
        Public p As Vector3
        Public n As Vector3
        Public tu As Single
        Public tv As Single
        Public tw As Single

        Public tangent As Vector3
        Public binormal As Vector3

        'Public tu2 As Single
        'Public tv2 As Single
    End Structure
    'Private terVertFmtTBN As VertexFormats = VertexFormats.Position Or VertexFormats.Normal Or VertexFormats.Texture1 Or VertexFormats.Texture2 Or VertexFormats.Texture3 Or VertexFormats.Texture4 Or VertexTextureCoordinate.Size3(0) Or VertexTextureCoordinate.Size3(1) Or VertexTextureCoordinate.Size3(2)
    'Private terVertFmtTBN As VertexFormats = VertexFormats.Position Or VertexFormats.Normal Or VertexFormats.Texture1 Or VertexFormats.Texture2 Or VertexFormats.Texture3 Or VertexTextureCoordinate.Size3(0) Or VertexTextureCoordinate.Size3(1) Or VertexTextureCoordinate.Size3(2)
    Private terVertFmtTBN As VertexFormats = VertexFormats.Position Or VertexFormats.Normal Or VertexFormats.Texture3 Or VertexTextureCoordinate.Size3(0) Or VertexTextureCoordinate.Size3(1) Or VertexTextureCoordinate.Size3(2)

    'Private Sub DoTangentCalcs()
    '	Dim uVerts(VertsTotal - 1) As terVertTBN

    '	'Load our entire vertex buffer into an array
    '	For i As Int32 = 0 To VertsTotal - 1
    '		Dim pnSrc As terVert = moVerts(i)

    '		With uVerts(i)
    '			.p = pnSrc.p
    '			.n = pnSrc.n
    '			.tu = pnSrc.tu
    '			.tv = pnSrc.tv
    '			.tw = pnSrc.tw

    '			.tu2 = pnSrc.tu2
    '			.tv2 = pnSrc.tv2
    '			'.tw2 = pnSrc.tw2

    '			.binormal = New Vector3(0, 0, 0)
    '			.tangent = New Vector3(0, 0, 0)
    '		End With
    '	Next i

    '	Device.IsUsingEventHandlers = False
    '	Dim oTmpVB As VertexBuffer = New VertexBuffer(uVerts(0).GetType, uVerts.Length, moDevice, Usage.WriteOnly, terVertFmtTBN, Pool.Managed)
    '	Device.IsUsingEventHandlers = True

    '	Dim oIndexArr As Array = moIB.Lock(0, LockFlags.ReadOnly)
    '	'Assume it is a triangle list?
    '	For i As Int32 = 0 To oIndexArr.Length - 3 'Step 3
    '		Dim iIndex1 As Int32 = CInt(oIndexArr.GetValue(i))
    '		Dim iIndex2 As Int32 = CInt(oIndexArr.GetValue(i + 1))
    '		Dim iIndex3 As Int32 = CInt(oIndexArr.GetValue(i + 2))

    '		'Now, get our vecs and UVs
    '		Dim vA, vB, vC As Vector3
    '		Dim uva, uvb, uvc As Vector2
    '		With uVerts(iIndex1)
    '			vA.X = .p.X : vA.Y = .p.Y : vA.Z = .p.Z
    '			uva.X = .tu : uva.Y = .tv
    '		End With
    '		With uVerts(iIndex2)
    '			vB.X = .p.X : vB.Y = .p.Y : vB.Z = .p.Z
    '			uvb.X = .tu : uvb.Y = .tv
    '		End With
    '		With uVerts(iIndex3)
    '			vC.X = .p.X : vC.Y = .p.Y : vC.Z = .p.Z
    '			uvc.X = .tu : uvc.Y = .tv
    '		End With

    '		Dim fX1 As Single = vB.X - vA.X
    '		Dim fX2 As Single = vC.X - vA.X
    '		Dim fY1 As Single = vB.Y - vA.Y
    '		Dim fY2 As Single = vC.Y - vA.Y
    '		Dim fZ1 As Single = vB.Z - vA.Z
    '		Dim fZ2 As Single = vC.Z - vA.Z
    '		Dim fS1 As Single = uvb.X - uva.X
    '		Dim fS2 As Single = uvc.X - uva.X
    '		Dim fT1 As Single = uvb.Y - uva.Y
    '		Dim fT2 As Single = uvc.Y - uva.Y

    '		Dim fDenom As Single = fS1 * fT2 - fS2 * fT1
    '		If Math.Abs(fDenom) < 0.0001 Then Continue For

    '		Dim fR As Single = 1.0F / fDenom
    '		Dim vecS As Vector3
    '		vecS.X = (fT2 * fX1 - fT1 * fX2) * fR
    '		vecS.Y = (fT2 * fY1 - fT1 * fY2) * fR
    '		vecS.Z = (fT2 * fZ1 - fT1 * fZ2) * fR

    '		Dim vecT As Vector3
    '		vecT.X = (fS1 * fX2 - fS2 * fX1) * fR
    '		vecT.Y = (fS1 * fY2 - fS2 * fY1) * fR
    '		vecT.Z = (fS1 * fZ2 - fS2 * fZ1) * fR

    '		With uVerts(iIndex1)
    '			.tangent = Vector3.Add(.tangent, vecS)
    '			.binormal = Vector3.Add(.binormal, vecT)
    '		End With
    '		With uVerts(iIndex2)
    '			.tangent = Vector3.Add(.tangent, vecS)
    '			.binormal = Vector3.Add(.binormal, vecT)
    '		End With
    '		With uVerts(iIndex3)
    '			.tangent = Vector3.Add(.tangent, vecS)
    '			.binormal = Vector3.Add(.binormal, vecT)
    '		End With
    '	Next i
    '	moIB.Unlock()

    '	For i As Int32 = 0 To VertsTotal - 1
    '		uVerts(i).binormal.Normalize()
    '		uVerts(i).tangent.Normalize()
    '	Next i

    '	oTmpVB.SetData(uVerts, 0, LockFlags.Discard)
    '	moVB.Dispose()
    '	moVB = Nothing
    '	moVB = oTmpVB
    'End Sub
    Private Function DoTangentCalcs(ByRef oVerts() As terVert, ByVal lIndices() As Int32) As terVertTBN()

        Dim vecT(oVerts.Length - 1) As Vector3
        Dim vecB(oVerts.Length - 1) As Vector3

        Dim lCount As Int32 = lIndices.Length \ 3

        For i As Int32 = 0 To lCount - 1
            Dim v1 As Vector3 = oVerts(lIndices(i * 3)).p
            Dim v2 As Vector3 = oVerts(lIndices(i * 3 + 1)).p
            Dim v3 As Vector3 = oVerts(lIndices(i * 3 + 2)).p

            Dim w1 As Vector3
            With oVerts(lIndices(i * 3))
                w1 = New Vector3(.tu, .tv, .tw)
            End With
            Dim w2 As Vector3
            With oVerts(lIndices(i * 3 + 1))
                w2 = New Vector3(.tu, .tv, .tw)
            End With
            Dim w3 As Vector3
            With oVerts(lIndices(i * 3 + 2))
                w3 = New Vector3(.tu, .tv, .tw)
            End With

            Dim x1 As Single = v2.X - v1.X
            Dim x2 As Single = v3.X - v1.X
            Dim y1 As Single = v2.Y - v1.Y
            Dim y2 As Single = v3.Y - v1.Y
            Dim z1 As Single = v2.Z - v1.Z
            Dim z2 As Single = v3.Z - v1.Z

            Dim s1 As Single = w2.X - w1.X
            Dim s2 As Single = w3.X - w1.X
            Dim t1 As Single = w2.Y - w1.Y
            Dim t2 As Single = w3.Y - w1.Y

            Dim r As Single = 1.0F / (s1 * t2 - s2 * t1)

            Dim sDir As New Vector3((t2 * x1 - t1 * x2) * r, _
              (t2 * y1 - t1 * y2) * r, _
              (t2 * z1 - t1 * z2) * r)

            Dim tDir As New Vector3((s1 * x2 - s2 * x1) * r, _
              (s1 * y2 - s2 * y1) * r, _
              (s1 * z2 - s2 * z1) * r)

            vecT(lIndices(i * 3)) += sDir
            vecT(lIndices(i * 3 + 1)) += sDir
            vecT(lIndices(i * 3 + 2)) += sDir

            vecB(lIndices(i * 3)) += tDir
            vecB(lIndices(i * 3 + 1)) += tDir
            vecB(lIndices(i * 3 + 2)) += tDir

        Next i

        Dim uVerts(oVerts.Length - 1) As terVertTBN

        Dim fFullSize As Single = TerrainClass.Width * Me.CellSpacing
        Dim fHalfSize As Single = fFullSize / 2.0F

        For i As Int32 = 0 To oVerts.Length - 1

            With uVerts(i)
                .n = oVerts(i).n
                .p = oVerts(i).p
                .tu = oVerts(i).tu
                .tv = oVerts(i).tv
                .tw = oVerts(i).tw
                '.tu2 = (.p.X + fHalfSize) / fFullSize  'oVerts(i).tu2
                '.tv2 = (.p.Z + fHalfSize) / fFullSize 'oVerts(i).tv2

                Dim t As Vector3 = vecT(i)
                .tangent = Vector3.Normalize(t - .n * (Vector3.Dot(.n, t)))

                t = vecB(i)
                .binormal = Vector3.Normalize(t - .n * (Vector3.Dot(.n, t)))
            End With
        Next i

        Return uVerts
    End Function
#End Region

#Region " Constant Expressions "
    Public Const Width As Int32 = 240 '256
    Public Const Height As Int32 = 240 '256
    Private Const mlPASSES As Int32 = 5
    Private Const mlQuads As Int32 = 24      '8x8 quad = 64 quads total, this is to segment the total vertex buffer further
    Private Const mlVertsPerQuad As Int32 = CInt(Width / mlQuads)
    Private Const mlHalfHeight As Int32 = CInt(Height / 2)
    Private Const mlHalfWidth As Int32 = CInt(Width / 2)
    Private Const VertsTotal As Int32 = Width * Height
    Private Const QuadsX As Int32 = Width - 1
    Private Const QuadsZ As Int32 = Height - 1
    Private Const TrisX As Int32 = CInt(QuadsX / 2)
#End Region

    Public CellSpacing As Int32 = 200
    Public ml_Y_Mult As Single = 5.0
    Public DrawWater As Boolean = False
    Public MapType As Int32
    Public WaterHeight As Byte      'height of the water level

    Public NormalsReady As Boolean = False

    'FOR ADVANCED RENDER (BUMP MAPPED TERRAIN)
    Private mlTris As Int32 = 0
    Private mlTris_XOrZ As Int32 = 0
    Private mlTris_XZ As Int32 = 0
    Private Shared moShader As TerrainShader

    Public Sub ClearShader()
        If moShader Is Nothing = False Then moShader.ReleaseMe()
        moShader = Nothing
        If moNormalVol Is Nothing = False Then moNormalVol.Dispose()
        moNormalVol = Nothing
        If moIllum Is Nothing = False Then moIllum.Dispose()
        moIllum = Nothing
        If moVolTex Is Nothing = False Then moVolTex.Dispose()
        If moTex Is Nothing = False Then moTex.Dispose()
        moVolTex = Nothing
        moTex = Nothing
    End Sub
#Region " Private Variables "
    'Device for rendering... be sure NEVER to dispose the device, it is the master device
    Private moDevice As Device
    'Our heightmap array, we need to keep this for... getting heights at locations
    Private HeightMap() As Byte
    'Indicates if the Heightmap has been generated yet
    Private mbHMReady As Boolean = False
    'Vertex buffer for *THIS* terrain, when I am not viewing *THIS* terrain, I need to destroy the buffer
    'MSC 06/05/08 - remarked this
    'Private moVB As VertexBuffer
    'Index Buffer, it is shared between all Terrain classes, it should be destroyed when the engine shuts down
    'Private Shared moIB As IndexBuffer
    'the vertex array.... it is cleared when viewing this terrain only
    Private moVerts() As terVert
    Private moVertsTBN() As terVertTBN
    'the Parent Planet ID... set in New()
    Private mlSeed As Int32
    'Triangle Normals for determining unit alignment along terrain
    Private HMNormals() As Vector3
    'Quad Start Index is the first index for the quad, used during rendering (makes engine faster)
    Private mlQuadStartIndex() As Int32
    'mlQuadX is the Left most location of the quad
    Private mlQuadX() As Int32
    'mlQuadZ is the top most location of the quad
    Private mlQuadZ() As Int32
    'mlQuadFOWIndex is the index of the oFOWGrid in goCurrentEnvir that this quad uses for a texture
    Private mlQuadFOWXIndex() As Int32
    Private mlQuadFOWZIndex() As Int32
    'moVolTex is our texture for use when volume textures are supported
    Private moVolTex As VolumeTexture
    'moTex is our texture for use when volume textures are not supported
    Private moTex As Texture
    'our normal tex, only volume texture
    Private Shared moNormalVol As VolumeTexture
    'illum for city creep
    Private Shared moIllum As Texture
    'material for the terrain
    Private moTerrMat As Material 

    Private mcbQuadBox() As CullBox

    Private mcbQB_East() As CullBox     'cullboxes for when the camera is on the EAST edge
    Private mcbQB_West() As CullBox     'cullboxes for when the camera is on the WEST edge

    'MSC 06/05/08 - for the new rendering method
    Private moQuadVB() As VertexBuffer
    Private myQuadVBDirty() As Byte
    Private moQuadIB As IndexBuffer
    Private moQuad_RIB As IndexBuffer   'the right sided index buffer 

    Private mlBaseIndices() As Int32
#End Region

#Region " Water Rendering "
    Private Shared muWaterVerts() As CustomVertex.PositionNormal
    Private Shared mfWaterMod() As Single
    Private Shared myWaterModCnt() As Int16
    Private Shared mlWaterWidthHeight As Int32 = 128
    Private Shared mlWaterSpacing As Int32
    Private Shared mvbWater As VertexBuffer
    Private Shared mibWater As IndexBuffer
    Private Shared mfMaxWaterHeight As Single = 2.5
    Private Shared mlWaterPrimCnt As Int32 = 32510 '8062 is for 64
    Private muWaterMat As Material

    Private Shared muWater() As CustomVertex.PositionNormal  'CustomVertex.PositionNormalColored

    Private mlLastCameraAtX As Int32
    Private mlLastCameraAtZ As Int32

    Public Shared Sub ForceWaterCreation()
        If mvbWater Is Nothing = False Then mvbWater.Dispose()
        mvbWater = Nothing
        If mibWater Is Nothing = False Then mibWater.Dispose()
        mibWater = Nothing
        Erase muWaterVerts
        Erase mfWaterMod
        Erase myWaterModCnt
    End Sub

    Private Sub CreateWater()
        Dim X As Int32
        Dim Z As Int32
        Dim fHeight As Single
        Dim lIdx As Int32

        'mlWaterWidthHeight = CInt(Val(oINI.GetString("GRAPHICS", "WaterQuadRes", "128")))
        mlWaterWidthHeight = 32
        mlWaterPrimCnt = ((mlWaterWidthHeight * (mlWaterWidthHeight - 1)) * 2) - 2

        'If muSettings.DrawProceduralWater = True Then

        ReDim muWaterVerts((mlWaterWidthHeight * mlWaterWidthHeight) - 1)
        ReDim mfWaterMod((mlWaterWidthHeight * mlWaterWidthHeight) - 1)
        ReDim myWaterModCnt((mlWaterWidthHeight * mlWaterWidthHeight) - 1)

        'now, calcualte our water spacing
        mlWaterSpacing = CInt(25000 / mlWaterWidthHeight)
        Dim lOffset As Int32 = CInt((mlWaterWidthHeight / 2) * mlWaterSpacing)
        mfMaxWaterHeight = ml_Y_Mult '2 * ml_Y_Mult

        For Z = 0 To mlWaterWidthHeight - 1
            For X = 0 To mlWaterWidthHeight - 1
                lIdx = (Z * mlWaterWidthHeight) + X
                fHeight = (Rnd() * (mfMaxWaterHeight * 2)) - mfMaxWaterHeight
                fHeight += (WaterHeight * ml_Y_Mult)
                'u = X / mlWaterWidthHeight
                'v = Z / mlWaterWidthHeight
                'u = X / 5
                'v = Z / 5
                'muWaterVerts(lIdx) = New CustomVertex.PositionNormalTextured((X * mlWaterSpacing) - lExt, fHeight, (Z * mlWaterSpacing) - lExt, 0, 1, 0, u, v)
                muWaterVerts(lIdx) = New CustomVertex.PositionNormal((X * mlWaterSpacing), fHeight, (Z * mlWaterSpacing) - lOffset, 0, 1, 0)
                mfWaterMod(lIdx) = CSng(((Rnd() * 1) - 0.5) * (ml_Y_Mult / 5))
                myWaterModCnt(lIdx) = CShort((Rnd() * 6) + 1)
            Next X
        Next Z

        'now, go back thru and smooth our vals
        Dim fSides As Single
        Dim fCorners As Single
        Dim fCenter As Single
        For Z = 0 To mlWaterWidthHeight - 1
            For X = 0 To mlWaterWidthHeight - 1
                fCorners = 0 : fSides = 0 : fCenter = 0
                lIdx = (Z * mlWaterWidthHeight) + X
                If X - 1 > -1 Then fSides += muWaterVerts(lIdx - 1).Y
                If X + 1 < mlWaterWidthHeight Then fSides += muWaterVerts(lIdx + 1).Y
                If Z - 1 > -1 Then fSides += muWaterVerts(lIdx - mlWaterWidthHeight).Y
                If Z + 1 < mlWaterWidthHeight Then fSides += muWaterVerts(lIdx + mlWaterWidthHeight).Y

                If X - 1 > -1 And Z - 1 > -1 Then fCorners += muWaterVerts(lIdx - mlWaterWidthHeight - 1).Y
                If X + 1 < mlWaterWidthHeight And Z - 1 > -1 Then fCorners += muWaterVerts(lIdx - mlWaterWidthHeight + 1).Y
                If X - 1 > -1 And Z + 1 < mlWaterWidthHeight Then fCorners += muWaterVerts(lIdx + mlWaterWidthHeight - 1).Y
                If X + 1 < mlWaterWidthHeight And Z + 1 < mlWaterWidthHeight Then fCorners += muWaterVerts(lIdx + mlWaterWidthHeight + 1).Y

                fCenter = muWaterVerts(lIdx).Y

                muWaterVerts(lIdx).Y = (fCorners / 16) + (fSides / 8) + (fCenter / 4)
            Next X
        Next Z

        Device.IsUsingEventHandlers = False
        mvbWater = New VertexBuffer(muWaterVerts(0).GetType, muWaterVerts.Length, moDevice, Usage.WriteOnly, CustomVertex.PositionNormal.Format, Pool.Managed)
        Device.IsUsingEventHandlers = True
        ComputeWaterNormals()

        mvbWater.SetData(muWaterVerts, 0, LockFlags.None)

        LoadWaterIndexBuffer()


        'Else
        ''create our water
        'Dim lExt As Int32 = ((50000 / CellSpacing) + 1) * CellSpacing

        'ReDim muWater(3)
        ''muWater(0) = New CustomVertex.PositionNormalColored(-lExt, WaterHeight * ml_Y_Mult, -lExt, 0, 1, 0, muWaterMat.Diffuse.ToArgb)
        ''muWater(1) = New CustomVertex.PositionNormalColored(-lExt, WaterHeight * ml_Y_Mult, lExt, 0, 1, 0, muWaterMat.Diffuse.ToArgb)
        ''muWater(2) = New CustomVertex.PositionNormalColored(lExt, WaterHeight * ml_Y_Mult, -lExt, 0, 1, 0, muWaterMat.Diffuse.ToArgb)
        ''muWater(3) = New CustomVertex.PositionNormalColored(lExt, WaterHeight * ml_Y_Mult, lExt, 0, 1, 0, muWaterMat.Diffuse.ToArgb)
        'muWater(0) = New CustomVertex.PositionNormal(-lExt, WaterHeight * ml_Y_Mult, -lExt, 0, 1, 0)
        'muWater(1) = New CustomVertex.PositionNormal(-lExt, WaterHeight * ml_Y_Mult, lExt, 0, 1, 0)
        'muWater(2) = New CustomVertex.PositionNormal(lExt, WaterHeight * ml_Y_Mult, -lExt, 0, 1, 0)
        'muWater(3) = New CustomVertex.PositionNormal(lExt, WaterHeight * ml_Y_Mult, lExt, 0, 1, 0)
        'End If

    End Sub

    Private Sub ComputeWaterNormals()
        Dim Z As Int32
        Dim X As Int32

        Dim vecX As Vector3 = New Vector3()
        Dim vecZ As Vector3 = New Vector3()
        Dim vecN As Vector3

        'Dim fA As Single
        'Dim fB As Single
        'Dim fC As Single

        'Dim lIdx As Int32


        For Z = 1 To mlWaterWidthHeight - 2
            For X = 1 To mlWaterWidthHeight - 2
                vecX = Vector3.Subtract(muWaterVerts(Z * mlWaterWidthHeight + X + 1).Position, muWaterVerts(Z * mlWaterWidthHeight + X - 1).Position)
                vecZ = Vector3.Subtract(muWaterVerts((Z + 1) * mlWaterWidthHeight + X).Position, muWaterVerts((Z - 1) * mlWaterWidthHeight + X).Position)
                vecN = Vector3.Cross(vecZ, vecX)
                vecN.Normalize()
                muWaterVerts(Z * mlWaterWidthHeight + X).Normal = vecN
            Next X
        Next Z

        vecX = Nothing
        vecZ = Nothing
        vecN = Nothing
    End Sub

    Private Sub LoadWaterIndexBuffer()
        Dim lCnt As Int32 = ((mlWaterWidthHeight - 1) * 6) * (mlWaterWidthHeight - 1)
        Dim lIndex As Int32
        Dim Z As Int32
        Dim X As Int32
        Dim mlIndices() As Int32
        Dim lHalfMod As Int32 = CInt((mlWaterWidthHeight / 2) * mlWaterSpacing)
        Dim bBool As Boolean

        ReDim mlIndices(lCnt)
        Device.IsUsingEventHandlers = False
        mibWater = New IndexBuffer(lCnt.GetType, mlIndices.Length, moDevice, Usage.WriteOnly, Pool.Managed)
        Device.IsUsingEventHandlers = True

        lIndex = 0

        For Z = 0 To mlWaterWidthHeight

            If mlIndices.Length < lIndex + 4 Then
                bBool = False
            End If

            If Z Mod 2 = 0 Then
                For X = 0 To mlWaterWidthHeight - 1
                    mlIndices(lIndex) = X + (Z * mlWaterWidthHeight) : lIndex += 1
                    mlIndices(lIndex) = X + (Z * mlWaterWidthHeight) + mlWaterWidthHeight : lIndex += 1
                Next X
                If Z <> mlWaterWidthHeight - 2 Then
                    X -= 1
                    mlIndices(lIndex) = X + (Z * mlWaterWidthHeight)
                End If
            Else
                For X = mlWaterWidthHeight - 1 To 0 Step -1
                    mlIndices(lIndex) = X + (Z * mlWaterWidthHeight) : lIndex += 1
                    mlIndices(lIndex) = X + (Z * mlWaterWidthHeight) + mlWaterWidthHeight : lIndex += 1
                Next X
                If Z <> Height - 2 Then
                    X += 1
                    mlIndices(lIndex) = X + (Z * mlWaterWidthHeight)
                End If
            End If
        Next Z

        mibWater.SetData(mlIndices, 0, LockFlags.None)

        Erase mlIndices
    End Sub

    Private Sub RenderWater()
        Dim matWorld As Matrix
        Dim matTemp As Matrix

        'If gbDrawWater = True Then
        '    If mvbWater Is Nothing Then
        '        CreateWater()
        '        mlLastCameraAtX = goCamera.mlCameraAtX
        '        mlLastCameraAtZ = goCamera.mlCameraAtZ
        '    End If

        '    Dim X As Int32
        '    'Dim Z As Int32
        '    Dim bRandomize As Boolean = mlLastCameraAtX <> goCamera.mlCameraAtX Or mlLastCameraAtZ <> goCamera.mlCameraAtZ
        '    Dim fHeight As Single

        '    mlLastCameraAtX = goCamera.mlCameraAtX
        '    mlLastCameraAtZ = goCamera.mlCameraAtZ

        '    'configure our verts
        '    If bRandomize = True Then
        '        For X = 0 To muWaterVerts.Length - 1
        '            fHeight = (Rnd() * (mfMaxWaterHeight * 2)) - mfMaxWaterHeight
        '            fHeight += (WaterHeight * ml_Y_Mult)
        '            muWaterVerts(X).Y = fHeight
        '            mfWaterMod(X) = CSng(((Rnd() * 1) - 0.5) * (ml_Y_Mult / 5))
        '            myWaterModCnt(X) = CShort((Rnd() * 6) + 1)
        '        Next X
        '    Else
        '        For X = 0 To muWaterVerts.Length - 1
        '            muWaterVerts(X).Y += mfWaterMod(X)
        '            If Math.Abs(muWaterVerts(X).Y - (WaterHeight * ml_Y_Mult)) > mfMaxWaterHeight Then
        '                mfWaterMod(X) *= -1
        '                myWaterModCnt(X) -= 1S
        '                If myWaterModCnt(X) = 0 Then
        '                    If mfWaterMod(X) < 0 Then
        '                        mfWaterMod(X) = Rnd() * -0.5F
        '                    Else
        '                        mfWaterMod(X) = Rnd() * 0.5F
        '                    End If
        '                    mfWaterMod(X) *= (ml_Y_Mult / 5)
        '                    myWaterModCnt(X) = CShort((Rnd() * 6) + 1)
        '                End If
        '            End If
        '        Next X
        '    End If

        '    ComputeWaterNormals()
        '    mvbWater.SetData(muWaterVerts, 0, LockFlags.Discard)

        '    'ensure our renderstates are correct
        '    With moDevice
        '        .RenderState.AlphaBlendEnable = True
        '        .RenderState.SourceBlend = Blend.SourceAlpha
        '        .RenderState.DestinationBlend = Blend.InvSourceAlpha
        '    End With

        '    'NOTE: I was going to have textures, but I just didn't like them...
        '    moDevice.SetTexture(0, Nothing)
        '    moDevice.Material = muWaterMat

        '    'set our stream type, source, and indices... turn off culling
        '    moDevice.VertexFormat = CustomVertex.PositionNormal.Format
        '    moDevice.SetStreamSource(0, mvbWater, 0)
        '    moDevice.Indices = mibWater
        '    moDevice.RenderState.CullMode = Cull.None

        '    'set up our transforms
        '    Dim fAngle As Single = goCamera.GetCameraAngleDegrees()
        '    fAngle /= CSng(gdDegreePerRad)
        '    matWorld = Matrix.Identity
        '    matTemp = Matrix.Identity
        '    matTemp.RotateY(fAngle)
        '    matWorld.Multiply(matTemp)
        '    matTemp = Matrix.Identity
        '    matTemp = Matrix.Translation(goCamera.mlCameraX, 0, goCamera.mlCameraZ)
        '    matWorld.Multiply(matTemp)
        '    moDevice.Transform.World = matWorld

        '    'render our quad
        '    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, muWaterVerts.Length, 0, mlWaterPrimCnt)
        'Else

        'If muWater Is Nothing Then CreateWater()
        'With moDevice
        '    .RenderState.AlphaBlendEnable = True
        '    .RenderState.SourceBlend = Blend.SourceAlpha
        '    .RenderState.DestinationBlend = Blend.InvSourceAlpha
        'End With
        'matTemp = Matrix.Identity
        'matWorld = Matrix.Identity
        'matTemp = matTemp.Translation(goCamera.mlCameraAtX, 0, goCamera.mlCameraAtZ)
        'matWorld.Multiply(matTemp)
        'moDevice.Transform.World = matWorld
        'moDevice.VertexFormat = muWater(0).Format

        'moDevice.SetStreamSource(0, Nothing, 0)
        'moDevice.SetTexture(0, Nothing)
        'moDevice.SetTexture(1, Nothing)
        'moDevice.Material = muWaterMat

        'moDevice.DrawUserPrimitives(PrimitiveType.TriangleStrip, 2, muWater)
        If mvbWater Is Nothing Then
            CreateWater()
            mlLastCameraAtX = goCamera.mlCameraAtX
            mlLastCameraAtZ = goCamera.mlCameraAtZ
        End If
        'ensure our renderstates are correct
        With moDevice
            .RenderState.AlphaBlendEnable = True
            .RenderState.SourceBlend = Blend.SourceAlpha
            .RenderState.DestinationBlend = Blend.InvSourceAlpha
        End With

        'NOTE: I was going to have textures, but I just didn't like them...
        moDevice.SetTexture(0, Nothing)
        moDevice.Material = muWaterMat

        'set our stream type, source, and indices... turn off culling
        moDevice.VertexFormat = CustomVertex.PositionNormal.Format
        moDevice.SetStreamSource(0, mvbWater, 0)
        moDevice.Indices = mibWater
        moDevice.RenderState.CullMode = Cull.None

        'set up our transforms
        Dim fAngle As Single = goCamera.GetCameraAngleDegrees()
        fAngle /= CSng(gdDegreePerRad)
        matWorld = Matrix.Identity
        matTemp = Matrix.Identity
        matTemp.RotateY(fAngle)
        matWorld.Multiply(matTemp)
        matTemp = Matrix.Identity
        matTemp = Matrix.Translation(goCamera.mlCameraX, 0, goCamera.mlCameraZ)
        matWorld.Multiply(matTemp)
        moDevice.Transform.World = matWorld

        'render our quad
        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, muWaterVerts.Length, 0, mlWaterPrimCnt)

        'End If
    End Sub


#End Region

#Region " Fog of War Rendering "
    Private Shared moVisibleTex As Texture
    Private Shared moVisibleMat As Material
    Public Shared moDisk As Mesh
    Private Shared moMaxRangeDisk As Mesh
    Private Shared myDiskMapType As Byte
    Private Const ml_FOW_UPDATE_INTERVAL As Int32 = 500     'every half a second

    Private Shared msw_FOW As Stopwatch

    Public Shared bSaveFOWImage As Boolean = False

    Public Shared Sub ReleaseVisibleTexture()
        If moVisibleTex Is Nothing = False Then moVisibleTex.Dispose()
        moVisibleTex = Nothing
    End Sub

    Private Sub SaveFOWImage()
        Dim oSurf As Surface
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory

        If Right(sFile, 1) <> "\" Then sFile = sFile & "\"

        oSurf = moVisibleTex.GetSurfaceLevel(0)
        SurfaceLoader.Save(sFile & "FOW_" & Now.ToString("MM_dd_yyyy_hhmmss") & ".bmp", ImageFileFormat.Bmp, oSurf)
        oSurf.Dispose()
    End Sub

    Public Sub UpdateFOWTextures(ByVal lFOWCenterX As Int32, ByVal lFOWCenterZ As Int32)
        If muSettings.ShowFOWTerrainShading = False Then Return
        Dim matTemp As Matrix
        Dim matWorld As Matrix

        'Return
        'Store away our Original Surface so we can set it back when we are done
        Dim oOriginal As Surface '= moDevice.GetRenderTarget(0)
        Dim oScene As Surface = Nothing

        Dim matView As Matrix
        Dim matProj As Matrix

        'Check if we need to New any of these
        If moVisibleTex Is Nothing OrElse moVisibleTex.Disposed = True Then
            moVisibleTex = New Texture(moDevice, muSettings.FOWTextureResolution, muSettings.FOWTextureResolution, 1, Usage.RenderTarget, Format.R5G6B5, Pool.Default)
        End If

        If msw_FOW Is Nothing Then msw_FOW = Stopwatch.StartNew
        If msw_FOW.IsRunning = False Then
            msw_FOW.Reset()
            msw_FOW.Start()
        End If
        If msw_FOW.ElapsedMilliseconds < ml_FOW_UPDATE_INTERVAL Then Return
        msw_FOW.Reset()
        msw_FOW.Start()

        'validate that our disk object is something...
        'our visible color is 255,255,255,255 (white)
        'our shroud color is 255,128,128,128 (gray)
        If myDiskMapType <> MapType Then
            'Need to change the disk
            If moDisk Is Nothing = False Then moDisk.Dispose()
            If moMaxRangeDisk Is Nothing = False Then moMaxRangeDisk.Dispose()
            moDisk = Nothing
            moMaxRangeDisk = Nothing
        End If
        If moDisk Is Nothing Then
            'If MapType = PlanetType.eBarren Then
            '    moDisk = CreateFOWDisk(moDevice, System.Drawing.Color.FromArgb(255, 128, 128, 128), System.Drawing.Color.FromArgb(255, 48, 48, 48))
            '    moMaxRangeDisk = CreateFOWDisk(moDevice, System.Drawing.Color.FromArgb(255, 96, 96, 96), System.Drawing.Color.FromArgb(255, 48, 48, 48))
            'ElseIf MapType = PlanetType.eGeoPlastic Then
            '    moDisk = CreateFOWDisk(moDevice, System.Drawing.Color.FromArgb(255, 255, 192, 192), System.Drawing.Color.FromArgb(255, 48, 48, 48))
            '    moMaxRangeDisk = CreateFOWDisk(moDevice, System.Drawing.Color.FromArgb(255, 128, 48, 48), System.Drawing.Color.FromArgb(255, 48, 48, 48))
            'Else
            moDisk = CreateFOWDisk(moDevice, System.Drawing.Color.FromArgb(255, 255, 255, 255), System.Drawing.Color.FromArgb(255, 48, 48, 48))
            moMaxRangeDisk = CreateFOWDisk(moDevice, System.Drawing.Color.FromArgb(255, 220, 220, 220), System.Drawing.Color.FromArgb(255, 48, 48, 48))
            'End If
            myDiskMapType = CByte(MapType)
        End If

        'Store our matrices beforehand...
        matView = moDevice.Transform.View
        matProj = moDevice.Transform.Projection

        'Ok, store our original surface
        oOriginal = moDevice.GetRenderTarget(0)

        'ensure our renderstates are set correctly
        moDevice.RenderState.SourceBlend = Blend.SourceColor
        moDevice.RenderState.DestinationBlend = Blend.InvDestinationColor
        moDevice.RenderState.BlendOperation = BlendOperation.Max

        'set our material and turn off lighting
        moDevice.Material = moVisibleMat
        moDevice.SetTexture(0, Nothing)
        moDevice.SetTexture(1, Nothing)
        moDevice.RenderState.Lighting = False
        moDevice.RenderState.ZBufferWriteEnable = False
        moDevice.RenderState.ZBufferEnable = False
        moDevice.RenderState.FogEnable = False

        'Get our surface to render to
        oScene = moVisibleTex.GetSurfaceLevel(0)

        'Now, set our render target to the texture's surface
        moDevice.SetRenderTarget(0, oScene)

        'Clear out our surface
        'If gbNoTimeOfDay = True Then
        '    'If MapType = PlanetType.eBarren Then
        '    moDevice.Clear(ClearFlags.Target, System.Drawing.Color.FromArgb(255, 128, 128, 128), 1.0F, 0)
        '    'Else : moDevice.Clear(ClearFlags.Target, System.Drawing.Color.FromArgb(255, 255, 255, 255), 1.0F, 0)
        '    'End If
        'Else
        moDevice.Clear(ClearFlags.Target, System.Drawing.Color.FromArgb(255, 80, 80, 80), 1.0F, 0)
        'End If

        'Set up our matrices
        moDevice.Transform.View = Matrix.LookAtLH(New Vector3(0, 100, 0), _
          New Vector3(0, 0, 0), New Vector3(0.0F, 0.0F, 1.0F))
        moDevice.Transform.Projection = Matrix.OrthoLH(Width * Me.CellSpacing, Height * Me.CellSpacing, 0.1, 5000)     'not sure about the 1000...

        'NOTE: IT IS ASSUMED THAT GOCURRENTENVIR = the planet we are rendering... otherwise, we should not be here!!!

        matTemp = Matrix.Identity
        matWorld = Matrix.Identity
        matTemp.Scale(1250, 30, 1250)
        matWorld.Multiply(matTemp)
        matTemp = Matrix.Identity
        'Dim fTempX As Single = goCamera.mlCameraX '/ 6375.0F
        'Dim fTempZ As Single = goCamera.mlCameraZ '/ 6375.0F
        '7200, 390, 4200
        matTemp.Translate(7200, 0, 4200)
        matWorld.Multiply(matTemp)
        matTemp = Nothing
        moDevice.Transform.World = matWorld
        moDisk.DrawSubset(0)


        'matTemp = Matrix.Identity
        'matWorld = Matrix.Identity
        'matTemp.Scale(25500, 20, 25500)
        'matWorld.Multiply(matTemp)
        'matTemp = Matrix.Identity
        'matTemp.Translate(0, 0, 0)
        'matWorld.Multiply(matTemp)
        'matTemp = Nothing
        'moDevice.Transform.World = matWorld
        'moMaxRangeDisk.DrawSubset(0)


        ''Now, loop thru the entity objects in this grid...
        'For X = 0 To goCurrentEnvir.lEntityUB
        '    'Is the entity valid?
        '    If goCurrentEnvir.lEntityIdx(X) <> -1 Then
        '        'With block for slight optimization
        '        With goCurrentEnvir.oEntity(X)
        '            If .OwnerID = glPlayerID AndAlso (.ObjTypeID <> ObjectType.eFacility OrElse (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0) AndAlso (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
        '                'do rendering here...
        '                matTemp = Matrix.Identity
        '                matWorld = Matrix.Identity
        '                matTemp.Scale(.oUnitDef.FOW_OptRadarRange, 30, .oUnitDef.FOW_OptRadarRange)
        '                matWorld.Multiply(matTemp)
        '                matTemp = Matrix.Identity
        '                matTemp.Translate(.LocX, 0, .LocZ)
        '                matWorld.Multiply(matTemp)
        '                matTemp = Nothing
        '                moDevice.Transform.World = matWorld

        '                moDisk.DrawSubset(0)
        '            End If
        '        End With
        '    End If
        'Next X

        'Now, restore our original surface to the device
        moDevice.SetRenderTarget(0, oOriginal)

        'now, clear our ZBuffer and renable writing
        ' moDevice.Clear(ClearFlags.ZBuffer, 0, 1.0F, 0)
        moDevice.RenderState.ZBufferWriteEnable = True
        moDevice.RenderState.ZBufferEnable = True

        'restore our matrices
        moDevice.Transform.View = matView
        moDevice.Transform.Projection = matProj
        moDevice.Transform.World = Matrix.Identity

        'Restore our default renderstates
        moDevice.RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
        moDevice.RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
        moDevice.RenderState.BlendOperation = BlendOperation.Add
        moDevice.RenderState.Lighting = True

        'Release all our objects
        If oScene Is Nothing = False Then oScene.Dispose()
        If oOriginal Is Nothing = False Then oOriginal.Dispose()
        oScene = Nothing
        oOriginal = Nothing

        'moDevice.RenderState.FogEnable = True

        If bSaveFOWImage = True Then
            SaveFOWImage()
            bSaveFOWImage = False
        End If
    End Sub

#End Region

    Public Sub New(ByRef oDevice As Device, ByVal lSeed As Int32)
        moDevice = oDevice
        mlSeed = lSeed

        With moVisibleMat
            .Ambient = System.Drawing.Color.Black
            .Diffuse = System.Drawing.Color.White
            .Emissive = System.Drawing.Color.Black
            .Specular = System.Drawing.Color.Black
        End With
    End Sub

    'Public Function Render(ByVal lClipPlane As Int32) As Int32
    '    'Dim lModX As Int32

    '    goCamera.SetupMatrices(moDevice)

    '    Dim X As Int32
    '    Dim lQuadsRendered As Int32 = 0
    '    Dim lMaxQuadDist As Int32   'if we were viewing along an axis...

    '    'Dim fTemp As Single = Math.Abs(gfTimeOfDay - 0.25F)

    '    'If fTemp < 0.0F OrElse fTemp > 0.5F Then
    '    '    fTemp = 0.0F
    '    'Else
    '    '    fTemp = Math.Abs(0.25F - fTemp)
    '    '    fTemp = (0.25F - fTemp) / 0.25F
    '    '    If fTemp < 0 Then fTemp = 0
    '    '    If fTemp > 1.0F Then fTemp = 1
    '    'End If

    '    'X = CInt(fTemp * 255)
    '    'moTerrMat.Specular = Color.FromArgb(255, X, X, X)

    '    'need to adjust X and Z for
    '    Dim lCurX As Int32 = CInt(Math.Floor((goCamera.mlCameraAtX + (mlHalfWidth * CellSpacing)) / CellSpacing))
    '    Dim lCurZ As Int32 = CInt(Math.Floor((goCamera.mlCameraAtZ + (mlHalfHeight * CellSpacing)) / CellSpacing))

    '    Dim lFOWCenterX As Int32
    '    Dim lFOWCenterZ As Int32
    '    Dim lLastFOWX As Int32 = -1
    '    Dim lLastFOWZ As Int32 = -1
    '    'Dim lTempFOWX As Int32
    '    'Dim lTempFOWZ As Int32
    '    'Dim lFOWTexIdx As Int32

    '    If mbHMReady = False Then GenerateTerrain(mlSeed)
    '    If bReloadVertexBuffer = True Then
    '        bReloadVertexBuffer = False
    '        If moVB Is Nothing = False Then moVB.Dispose()
    '        moVB = Nothing
    '    End If
    '    If moVB Is Nothing Then LoadVertexBuffer()
    '    If moIB Is Nothing Then LoadIndexBuffer()

    '    moDevice.Transform.World = Matrix.Identity

    '    moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, TextureFilter.Linear)
    '    moDevice.SetSamplerState(0, SamplerStageStates.MagFilter, TextureFilter.Linear)
    '    moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, TextureFilter.Linear)

    '    lMaxQuadDist = CInt(lClipPlane / CellSpacing)
    '    lMaxQuadDist += (lMaxQuadDist Mod mlVertsPerQuad)

    '    If muSettings.ShowFOWTerrainShading = True Then
    '        lFOWCenterX = CInt(Math.Floor(goCamera.mlCameraAtX / 16384))
    '        lFOWCenterZ = CInt(Math.Floor(goCamera.mlCameraAtZ / 16384))

    '        'Do our terrain texture upgrades...
    '        UpdateFOWTextures(lFOWCenterX, lFOWCenterZ)

    '        'ensure our zbuffers enabled and writing
    '        moDevice.RenderState.ZBufferEnable = True
    '        moDevice.RenderState.ZBufferWriteEnable = True

    '        'Are we using volume textures or just a single texture?
    '        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = True Then
    '            moDevice.SetTexture(0, moVolTex)
    '        Else
    '            moDevice.SetTexture(0, moTex)
    '        End If

    '        'moTerrMat.Specular = moDevice.RenderState.FogColor

    '        moDevice.Material = moTerrMat
    '        'turn off culling for this process
    '        moDevice.RenderState.CullMode = Cull.None
    '        moDevice.RenderState.Lighting = True

    '        'ok, we've already set our 0 texture, now, do our multi-texture for stage 0
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.SelectArg1)
    '        moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'Stage 0 is finished

    '        'now for stage 1
    '        moDevice.SetTexture(1, moVisibleTex)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.Current)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.TextureColor)
    '        'moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.AddSmooth)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Modulate2X)
    '        moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'stage 1 is finished

    '        'now for stage 2
    '        moDevice.SetTextureStageState(2, TextureStageStates.ColorOperation, TextureOperation.Modulate)
    '        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument1, TextureArgument.Current)
    '        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
    '        moDevice.SetTextureStageState(2, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'stage 2 is finished

    '        'now for stage 3
    '        moDevice.SetTextureStageState(3, TextureStageStates.ColorOperation, TextureOperation.Add)
    '        moDevice.SetTextureStageState(3, TextureStageStates.ColorArgument1, TextureArgument.Current)
    '        moDevice.SetTextureStageState(3, TextureStageStates.ColorArgument2, TextureArgument.Specular)
    '        moDevice.SetTextureStageState(3, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'stage 3 is finished

    '        moDevice.SetStreamSource(0, moVB, 0)
    '        moDevice.VertexFormat = terVertFmt
    '        moDevice.Indices = moIB

    '        'do our rendering
    '        For X = 0 To (mlQuads * mlQuads) - 1
    '            If mlQuadX(X) >= (lCurX - lMaxQuadDist) AndAlso mlQuadX(X) <= (lCurX + lMaxQuadDist) Then
    '                If mlQuadZ(X) >= (lCurZ - lMaxQuadDist) AndAlso mlQuadZ(X) <= (lCurZ + lMaxQuadDist) Then
    '                    lQuadsRendered += 1

    '                    If mlQuadX(X) = 230 AndAlso mlQuadZ(X) = 230 Then
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 178)
    '                    ElseIf mlQuadX(X) = 230 Then
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 198)
    '                    ElseIf mlQuadZ(X) = 230 Then
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 196)
    '                    Else
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 218)
    '                    End If
    '                End If
    '            End If
    '        Next X

    '        '================ MAP WRAPPING PASS ============================
    '        Dim lTempX As Int32 = -1
    '        Dim bXGrtr As Boolean

    '        Dim lMaxQuads As Int32 = (mlQuads - 1) * 10
    '        If lCurX - lMaxQuadDist < 0 Then
    '            'Ok, on left edge
    '            lTempX = (lCurX - lMaxQuadDist) + lMaxQuads
    '            'Now, check for QuadX > lTempX
    '            bXGrtr = True
    '        ElseIf lCurX + lMaxQuadDist >= lMaxQuads Then
    '            'Ok, on right edge
    '            lTempX = (lCurX + lMaxQuadDist) - lMaxQuads
    '            'Now, check for QuadX < lTempX
    '            bXGrtr = False
    '        End If

    '        'Ok, now, go back and render the quads we didn't have before...
    '        Dim oMatWorld As Matrix = Matrix.Identity
    '        If lTempX <> -1 Then
    '            oMatWorld = Matrix.Identity
    '            'Two passes... this one is for missed Left/Right edges
    '            If lTempX <> -1 Then
    '                If bXGrtr = True Then
    '                    'On left edge, so set our world matrix to transpose all items -X where X = CellSpacing * Width
    '                    oMatWorld.Multiply(Matrix.Translation((Me.CellSpacing * TerrainClass.Width * -1) + Me.CellSpacing, 0, 0))
    '                Else
    '                    'On right edge
    '                    oMatWorld.Multiply(Matrix.Translation((Me.CellSpacing * TerrainClass.Width) - Me.CellSpacing, 0, 0))
    '                End If

    '                moDevice.Transform.World = oMatWorld

    '                For X = 0 To (mlQuads * mlQuads) - 1
    '                    'Use the old method of culling, for now...
    '                    If (bXGrtr = True AndAlso (mlQuadX(X) > lTempX)) OrElse (bXGrtr = False AndAlso (mlQuadX(X) < lTempX)) Then
    '                        If mlQuadZ(X) >= (lCurZ - lMaxQuadDist) AndAlso mlQuadZ(X) <= (lCurZ + lMaxQuadDist) Then
    '                            lQuadsRendered += 1
    '                            If mlQuadX(X) = 230 AndAlso mlQuadZ(X) = 230 Then
    '                                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 178)
    '                            ElseIf mlQuadX(X) = 230 Then
    '                                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 198)
    '                            ElseIf mlQuadZ(X) = 230 Then
    '                                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 196)
    '                            Else
    '                                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 218)
    '                            End If
    '                        End If
    '                    End If
    '                Next X
    '            End If

    '            moDevice.Transform.World = Matrix.Identity
    '        End If

    '        'glQuadsRendered = lQuadsRendered

    '        moDevice.SetTexture(1, Nothing)
    '        moDevice.SetTexture(0, Nothing)

    '        'Now, reset our texture stage states
    '        moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.Modulate)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument2, TextureArgument.Current)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.Current)
    '        moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Disable)
    '    Else
    '        'Are we using volume textures or just a single texture?
    '        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = True Then
    '            moDevice.SetTexture(0, moVolTex)
    '        Else
    '            moDevice.SetTexture(0, moTex)
    '        End If
    '        moDevice.Material = moTerrMat
    '        'turn off culling for this process
    '        moDevice.RenderState.CullMode = Cull.None
    '        moDevice.SetStreamSource(0, moVB, 0)
    '        moDevice.VertexFormat = terVertFmt
    '        moDevice.Indices = moIB

    '        For X = 0 To (mlQuads * mlQuads) - 1
    '            If mlQuadX(X) >= (lCurX - lMaxQuadDist) AndAlso mlQuadX(X) <= (lCurX + lMaxQuadDist) Then
    '                If mlQuadZ(X) >= (lCurZ - lMaxQuadDist) AndAlso mlQuadZ(X) <= (lCurZ + lMaxQuadDist) Then
    '                    lQuadsRendered += 1

    '                    If mlQuadX(X) = 230 AndAlso mlQuadZ(X) = 230 Then
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 178)
    '                    ElseIf mlQuadX(X) = 230 Then
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 198)
    '                    ElseIf mlQuadZ(X) = 230 Then
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 196)
    '                    Else
    '                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 218)
    '                    End If
    '                End If
    '            End If
    '        Next X

    '        moDevice.SetTexture(0, Nothing)
    '    End If

    '    RenderWater()

    '    'turn culling back on
    '    moDevice.RenderState.CullMode = Cull.CounterClockwise

    '    moDevice.Transform.World = Matrix.Identity
    '    'moDevice.RenderState.FogEnable = False
    '    'For X = 0 To mlVentUB
    '    '    moVents(X).Render(False)
    '    'Next X
    '    'For X = 0 To mlSlowSteamUB
    '    '    moSlowSteam(X).Render(False)
    '    'Next X
    '    'For X = 0 To mlVolcanoVentUB
    '    '    moVolcanoVent(X).Render(False)
    '    'Next X
    '    'For X = 0 To mlAcidMistUB
    '    '    moAcidMist(X).Render(False)
    '    'Next X
    '    'moDevice.RenderState.FogEnable = True
    '    'If moPFX Is Nothing = False Then
    '    '    moPFX.Render(False)
    '    'End If

    '    Return lQuadsRendered
    'End Function

    Private moOcean As OceanRender
    'Public Function Render(ByVal lClipPlane As Int32) As Int32
    '    'Dim lModX As Int32 

    '    Dim X As Int32
    '    Dim lQuadsRendered As Int32 = 0
    '    Dim lMaxQuadDist As Int32   'if we were viewing along an axis...

    '    'need to adjust X and Z for
    '    Dim lCurX As Int32 = CInt(Math.Floor((goCamera.mlCameraAtX + (mlHalfWidth * CellSpacing)) / CellSpacing))
    '    Dim lCurZ As Int32 = CInt(Math.Floor((goCamera.mlCameraAtZ + (mlHalfHeight * CellSpacing)) / CellSpacing))

    '    Dim lFOWCenterX As Int32
    '    Dim lFOWCenterZ As Int32
    '    Dim lLastFOWX As Int32 = -1
    '    Dim lLastFOWZ As Int32 = -1

    '    If mbHMReady = False Then GenerateTerrain(mlSeed)

    '    If bReloadVertexBuffer = True Then
    '        bReloadVertexBuffer = False
    '        If moVB Is Nothing = False Then moVB.Dispose()
    '        moVB = Nothing
    '    End If
    '    If moVB Is Nothing Then LoadVertexBuffer()
    '    If moIB Is Nothing Then LoadIndexBuffer()

    '    moDevice.Transform.World = Matrix.Identity

    '    moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, TextureFilter.Linear)
    '    moDevice.SetSamplerState(0, SamplerStageStates.MagFilter, TextureFilter.Linear)
    '    moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, TextureFilter.Linear)

    '    lMaxQuadDist = CInt(lClipPlane / CellSpacing)
    '    lMaxQuadDist += (lMaxQuadDist Mod mlVertsPerQuad)

    '    If muSettings.ShowFOWTerrainShading = True Then
    '        lFOWCenterX = CInt(Math.Floor(goCamera.mlCameraAtX / 16384))
    '        lFOWCenterZ = CInt(Math.Floor(goCamera.mlCameraAtZ / 16384))

    '        'Do our terrain texture upgrades...
    '        UpdateFOWTextures(lFOWCenterX, lFOWCenterZ)

    '        'ensure our zbuffers enabled and writing
    '        moDevice.RenderState.ZBufferEnable = True
    '        moDevice.RenderState.ZBufferWriteEnable = True

    '        'Are we using volume textures or just a single texture?
    '        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = True Then
    '            moDevice.SetTexture(0, moVolTex)
    '        Else
    '            moDevice.SetTexture(0, moTex)
    '        End If

    '        'moTerrMat.Specular = moDevice.RenderState.FogColor

    '        moDevice.Material = moTerrMat
    '        'turn off culling for this process
    '        moDevice.RenderState.CullMode = Cull.None
    '        moDevice.RenderState.Lighting = True

    '        'ok, we've already set our 0 texture, now, do our multi-texture for stage 0
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.SelectArg1)
    '        moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'Stage 0 is finished

    '        'now for stage 1
    '        moDevice.SetTexture(1, moVisibleTex)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.Current)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.TextureColor)
    '        'moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.AddSmooth)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Modulate2X)
    '        moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'stage 1 is finished

    '        Dim lMag As Int32 = moDevice.GetSamplerStageStateInt32(1, SamplerStageStates.MagFilter)
    '        If moDevice.DeviceCaps.TextureFilterCaps.SupportsMagnifyLinear = True AndAlso muSettings.SmoothFOW = True Then
    '            moDevice.SetSamplerState(1, SamplerStageStates.MagFilter, TextureFilter.Linear)
    '        End If

    '        'now for stage 2
    '        moDevice.SetTextureStageState(2, TextureStageStates.ColorOperation, TextureOperation.Modulate)
    '        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument1, TextureArgument.Current)
    '        moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
    '        moDevice.SetTextureStageState(2, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'stage 2 is finished

    '        'now for stage 3
    '        moDevice.SetTextureStageState(3, TextureStageStates.ColorOperation, TextureOperation.Add)
    '        moDevice.SetTextureStageState(3, TextureStageStates.ColorArgument1, TextureArgument.Current)
    '        moDevice.SetTextureStageState(3, TextureStageStates.ColorArgument2, TextureArgument.Specular)
    '        moDevice.SetTextureStageState(3, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        'stage 3 is finished

    '        moDevice.SetStreamSource(0, moVB, 0)
    '        moDevice.VertexFormat = terVertFmt
    '        moDevice.Indices = moIB

    '        goCamera.SetupMatrices(moDevice) ', CurrentView.ePlanetView)

    '        'do our rendering
    '        For X = 0 To (mlQuads * mlQuads) - 1

    '            'If goCamera.CullObject(mcbQuadBox(X)) = False Then
    '            lQuadsRendered += 1

    '            If mlQuadX(X) = 230 AndAlso mlQuadZ(X) = 230 Then
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 178)
    '            ElseIf mlQuadX(X) = 230 Then
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 198)
    '            ElseIf mlQuadZ(X) = 230 Then
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 196)
    '            Else
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 218)
    '            End If
    '            'End If
    '        Next X


    '        '================ MAP WRAPPING PASS ============================
    '        'Dim lTempX As Int32 = -1
    '        'Dim bXGrtr As Boolean

    '        'Dim lMaxQuads As Int32 = (mlQuads - 1) * 10
    '        'If lCurX - lMaxQuadDist < 0 Then
    '        '    Ok, on left edge
    '        '    lTempX = (lCurX - lMaxQuadDist) + lMaxQuads
    '        '    Now, check for QuadX > lTempX
    '        '    bXGrtr = True
    '        'ElseIf lCurX + lMaxQuadDist >= lMaxQuads Then
    '        '    Ok, on right edge
    '        '    lTempX = (lCurX + lMaxQuadDist) - lMaxQuads
    '        '    Now, check for QuadX < lTempX
    '        '    bXGrtr = False
    '        'End If

    '        'Ok, now, go back and render the quads we didn't have before...
    '        'Dim oMatWorld As Matrix = Matrix.Identity
    '        'If lTempX <> -1 Then
    '        '    oMatWorld = Matrix.Identity
    '        '    Two passes... this one is for missed Left/Right edges
    '        '    If lTempX <> -1 Then
    '        '        If bXGrtr = True Then
    '        '            On left edge, so set our world matrix to transpose all items -X where X = CellSpacing * Width
    '        '            oMatWorld.Multiply(Matrix.Translation((Me.CellSpacing * TerrainClass.Width * -1) + Me.CellSpacing, 0, 0))
    '        '        Else
    '        '            On right edge
    '        '            oMatWorld.Multiply(Matrix.Translation((Me.CellSpacing * TerrainClass.Width) - Me.CellSpacing, 0, 0))
    '        '        End If

    '        '        moDevice.Transform.World = oMatWorld

    '        '        For X = 0 To (mlQuads * mlQuads) - 1
    '        '            Use the old method of culling, for now...
    '        '            If (bXGrtr = True AndAlso goCamera.CullObject(mcbQB_East(X)) = False) OrElse (bXGrtr = False AndAlso goCamera.CullObject(mcbQB_West(X)) = False) Then
    '        '                lQuadsRendered += 1
    '        '                If mlQuadX(X) = 230 AndAlso mlQuadZ(X) = 230 Then
    '        '                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 178)
    '        '                ElseIf mlQuadX(X) = 230 Then
    '        '                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 198)
    '        '                ElseIf mlQuadZ(X) = 230 Then
    '        '                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 196)
    '        '                Else
    '        '                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 218)
    '        '                End If
    '        '            End If
    '        '        Next X
    '        '    End If

    '        '    moDevice.Transform.World = Matrix.Identity
    '        'End If
    '        '=========== END OF MAP WRAPPING PASSES =============


    '        'glQuadsRendered = lQuadsRendered

    '        moDevice.SetTexture(1, Nothing)
    '        moDevice.SetTexture(0, Nothing)

    '        If moDevice.DeviceCaps.TextureFilterCaps.SupportsMagnifyLinear = True AndAlso muSettings.SmoothFOW = True Then
    '            moDevice.SetSamplerState(1, SamplerStageStates.MagFilter, lMag)
    '        End If

    '        'Now, reset our texture stage states
    '        moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.Modulate)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
    '        moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument2, TextureArgument.Current)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.Current)
    '        moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
    '        moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Disable)
    '    Else
    '        'Are we using volume textures or just a single texture?
    '        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = True Then
    '            moDevice.SetTexture(0, moVolTex)
    '        Else
    '            moDevice.SetTexture(0, moTex)
    '        End If
    '        moDevice.Material = moTerrMat
    '        'turn off culling for this process
    '        moDevice.RenderState.CullMode = Cull.None
    '        moDevice.SetStreamSource(0, moVB, 0)
    '        moDevice.VertexFormat = terVertFmt
    '        moDevice.Indices = moIB

    '        goCamera.SetupMatrices(moDevice) ', CurrentView.ePlanetView)

    '        'do our rendering
    '        For X = 0 To (mlQuads * mlQuads) - 1

    '            'If goCamera.CullObject(mcbQuadBox(X)) = False Then
    '            lQuadsRendered += 1

    '            If mlQuadX(X) = 230 AndAlso mlQuadZ(X) = 230 Then
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 178)
    '            ElseIf mlQuadX(X) = 230 Then
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 198)
    '            ElseIf mlQuadZ(X) = 230 Then
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 196)
    '            Else
    '                moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleStrip, 0, 0, VertsTotal, mlQuadStartIndex(X), 218)
    '            End If
    '            'End If
    '        Next X

    '        moDevice.SetTexture(0, Nothing)
    '    End If

    '    'If MapType <> PlanetType.eBarren Then
    '    '    'If moDevice.DeviceCaps.VertexShaderVersion.Major >= 2 AndAlso moDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
    '    '    '    If moWaterFX Is Nothing Then moWaterFX = New OceanShader(moDevice)
    '    '    '    moWaterFX.RenderWater(CSng(Me.WaterHeight) * Me.ml_Y_Mult, Me.CellSpacing)
    '    '    'Else 
    '    If muSettings.WaterRenderMethod = 1 Then
    '        RenderWaterNew(True)
    '    Else : RenderWater()
    '    End If

    '    '    'End If
    '    'End If


    '    'turn culling back on
    '    moDevice.RenderState.CullMode = Cull.CounterClockwise

    '    'Render particle FX here
    '    'If moPFX Is Nothing = False Then
    '    '    moPFX.Render(False)
    '    'End If

    '    Return lQuadsRendered
    'End Function
    Public Function Render(ByVal lClipPlane As Int32) As Int32
        'Dim lModX As Int32 

        'Dim lQuadsRendered As Int32 = 0
        Dim lMaxQuadDist As Int32   'if we were viewing along an axis...

        'need to adjust X and Z for
        Dim lCurX As Int32 = CInt(Math.Floor((goCamera.mlCameraAtX + (mlHalfWidth * CellSpacing)) / CellSpacing))
        Dim lCurZ As Int32 = CInt(Math.Floor((goCamera.mlCameraAtZ + (mlHalfHeight * CellSpacing)) / CellSpacing))

        If mbHMReady = False Then GenerateTerrain(mlSeed)

        If bReloadVertexBuffer = True Then
            bReloadVertexBuffer = False
            LoadVertexBuffer(False)
        End If
        If moQuadVB Is Nothing OrElse mlQuadX Is Nothing OrElse mlQuadZ Is Nothing Then LoadVertexBuffer(False)
        'If moIB Is Nothing Then LoadIndexBuffer()
        If moTex Is Nothing AndAlso moVolTex Is Nothing Then SetMatsTexsAndGetWaterColor()

        moDevice.Transform.World = Matrix.Identity

        moDevice.SetSamplerState(0, SamplerStageStates.MinFilter, TextureFilter.Linear)
        moDevice.SetSamplerState(0, SamplerStageStates.MagFilter, TextureFilter.Linear)
        moDevice.SetSamplerState(0, SamplerStageStates.MipFilter, TextureFilter.Linear)

        lMaxQuadDist = CInt(lClipPlane / CellSpacing)
        lMaxQuadDist += (lMaxQuadDist Mod mlVertsPerQuad)

        If muSettings.ShowFOWTerrainShading = True Then
            Dim lFOWCenterX As Int32 = (goCamera.mlCameraAtX \ 16384I)
            Dim lFOWCenterZ As Int32 = (goCamera.mlCameraAtZ \ 16384I)
            UpdateFOWTextures(lFOWCenterX, lFOWCenterZ)

            'ensure our zbuffers enabled and writing
            moDevice.RenderState.ZBufferEnable = True
            moDevice.RenderState.ZBufferWriteEnable = True

            'turn off culling for this process
            'moDevice.RenderState.CullMode = Cull.None		'MSC
            moDevice.RenderState.Lighting = True

            Dim lMag As Int32 = moDevice.GetSamplerStageStateInt32(1, SamplerStageStates.MagFilter)
            goCamera.SetupMatrices(moDevice)

            moDevice.Material = moTerrMat

            If muSettings.BumpMapTerrain = True Then
                'moDevice.Transform.World = Matrix.Identity
                If moShader Is Nothing Then moShader = New TerrainShader(moDevice)
                If moNormalVol Is Nothing OrElse (moIllum Is Nothing AndAlso muSettings.IlluminationMapTerrain = True) Then SetMatsTexsAndGetWaterColor()
                moShader.PrepareToRender(Vector3.Normalize(moDevice.Lights(0).Direction), moDevice.Lights(0).Diffuse, moDevice.Lights(0).Ambient, moDevice.Lights(0).Specular, (TerrainClass.Width * Me.CellSpacing) / 2.0F, (TerrainClass.Width * Me.CellSpacing))
                If moNormalVol Is Nothing Then moNormalVol = moResMgr.GetVolNormalTexture("TerranNrm.dds")
                If moIllum Is Nothing Then moIllum = moResMgr.GetTexture("terran_i.dds", EpicaResourceManager.eGetTextureType.NoSpecifics, "terrdin.pak")
                moShader.SetTextures(moVolTex, moNormalVol, moIllum, moVisibleTex)
                moShader.UpdateMatrices()
            Else
                'Are we using volume textures or just a single texture?
                If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = True Then
                    moDevice.SetTexture(0, moVolTex)
                Else
                    moDevice.SetTexture(0, moTex)
                End If

                'ok, we've already set our 0 texture, now, do our multi-texture for stage 0
                moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
                moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.SelectArg1)
                moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.Disable)
                'Stage 0 is finished

                'now for stage 1
                moDevice.SetTexture(1, moVisibleTex)
                moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.Current)
                moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.TextureColor)
                'moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.AddSmooth)
                moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Modulate2X)
                moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
                'stage 1 is finished

                If moDevice.DeviceCaps.TextureFilterCaps.SupportsMagnifyLinear = True AndAlso muSettings.SmoothFOW = True Then
                    moDevice.SetSamplerState(1, SamplerStageStates.MagFilter, TextureFilter.Linear)
                End If

                'now for stage 2
                moDevice.SetTextureStageState(2, TextureStageStates.ColorOperation, TextureOperation.Modulate)
                moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument1, TextureArgument.Current)
                moDevice.SetTextureStageState(2, TextureStageStates.ColorArgument2, TextureArgument.Diffuse)
                moDevice.SetTextureStageState(2, TextureStageStates.AlphaOperation, TextureOperation.Disable)
                'stage 2 is finished

                'now for stage 3
                moDevice.SetTextureStageState(3, TextureStageStates.ColorOperation, TextureOperation.Add)
                moDevice.SetTextureStageState(3, TextureStageStates.ColorArgument1, TextureArgument.Current)
                moDevice.SetTextureStageState(3, TextureStageStates.ColorArgument2, TextureArgument.Specular)
                moDevice.SetTextureStageState(3, TextureStageStates.AlphaOperation, TextureOperation.Disable)
                'stage 3 is finished
                moDevice.VertexFormat = terVertFmt
            End If

            RenderGeometry(lCurX, lMaxQuadDist) 'lQuadsRendered

            moDevice.SetTexture(1, Nothing)
            moDevice.SetTexture(0, Nothing)

            If moDevice.DeviceCaps.TextureFilterCaps.SupportsMagnifyLinear = True AndAlso muSettings.SmoothFOW = True Then
                moDevice.SetSamplerState(1, SamplerStageStates.MagFilter, lMag)
            End If

            If muSettings.BumpMapTerrain = True Then
                moShader.EndRender()
            Else
                'Now, reset our texture stage states
                moDevice.SetTextureStageState(0, TextureStageStates.AlphaOperation, TextureOperation.SelectArg1)
                moDevice.SetTextureStageState(0, TextureStageStates.ColorOperation, TextureOperation.Modulate)
                moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                moDevice.SetTextureStageState(0, TextureStageStates.ColorArgument2, TextureArgument.Current)
                moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument1, TextureArgument.TextureColor)
                moDevice.SetTextureStageState(1, TextureStageStates.ColorArgument2, TextureArgument.Current)
                moDevice.SetTextureStageState(1, TextureStageStates.AlphaOperation, TextureOperation.Disable)
                moDevice.SetTextureStageState(1, TextureStageStates.ColorOperation, TextureOperation.Disable)
            End If

        Else
            'Are we using volume textures or just a single texture?
            If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = True Then
                moDevice.SetTexture(0, moVolTex)
            Else
                moDevice.SetTexture(0, moTex)
            End If
            moDevice.Material = moTerrMat

            goCamera.SetupMatrices(moDevice)

            RenderGeometry(lCurX, lMaxQuadDist)

            moDevice.SetTexture(0, Nothing)
        End If

        'If MapType <> PlanetType.eBarren Then

        Dim bRenderNew As Boolean = False
        If moDevice.DeviceCaps.VertexShaderVersion.Major >= 2 AndAlso moDevice.DeviceCaps.PixelShaderVersion.Major >= 2 Then
            bRenderNew = muSettings.WaterRenderMethod = 1
        End If
        If bRenderNew = True Then RenderWaterNew(True) Else RenderWater()

        'End If

        ''turn culling back on
        moDevice.RenderState.CullMode = Cull.CounterClockwise

        ''Render particle FX here
        'If moPFX Is Nothing = False Then
        '    moPFX.Render(False)
        'End If

        Return 1
    End Function
    Private Shared moTBN_VD As VertexDeclaration
    Private Shared mo_VD As VertexDeclaration

    Private Function RenderGeometry(ByVal lCurX As Int32, ByVal lMaxQuadDist As Int32) As Int32
        'renders the geometry regardless of the graphics settings used

        Dim lQuadsRendered As Int32 = 0

        Try

            moDevice.RenderState.CullMode = Cull.Clockwise
            Dim lTmpVT As Int32 = (mlVertsPerQuad + 1) * (mlVertsPerQuad + 1)
            Dim lTmpVT_XOrZ As Int32 = (mlVertsPerQuad + 1) * mlVertsPerQuad
            Dim lTmpVT_XZ As Int32 = (mlVertsPerQuad) * mlVertsPerQuad
            If muSettings.BumpMapTerrain = True Then
                'moDevice.VertexFormat = terVertFmtTBN
                If moTBN_VD Is Nothing Then
                    Device.IsUsingEventHandlers = False
                    Dim elems(5) As VertexElement
                    elems(0) = New VertexElement(0, 0, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Position, 0)
                    elems(1) = New VertexElement(0, 12, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Normal, 0)
                    elems(2) = New VertexElement(0, 24, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 0)
                    'elems(3) = New VertexElement(0, 36, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 1)
                    'elems(4) = New VertexElement(0, 48, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 2)
                    elems(3) = New VertexElement(0, 36, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Tangent, 0)
                    elems(4) = New VertexElement(0, 48, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.BiNormal, 0)
                    elems(5) = VertexElement.VertexDeclarationEnd
                    moTBN_VD = New VertexDeclaration(moDevice, elems)
                    Device.IsUsingEventHandlers = True
                End If
                moDevice.VertexDeclaration = moTBN_VD
            Else
                moDevice.VertexFormat = terVertFmt
            End If
            moDevice.Indices = moQuadIB

            For X As Int32 = 0 To (mlQuads * mlQuads) - 1
                If goCamera.CullObject(mcbQuadBox(X)) = False Then
                    lQuadsRendered += 1
                    moDevice.SetStreamSource(0, moQuadVB(X), 0)

                    If mlQuadX(X) = 230 Then
                        moDevice.Indices = moQuad_RIB
                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, lTmpVT_XOrZ, 0, mlTris_XOrZ)
                        moDevice.Indices = moQuadIB
                    ElseIf mlQuadZ(X) = 230 Then
                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, mlTris_XOrZ, 0, mlTris_XOrZ)
                    Else
                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, lTmpVT, 0, mlTris)
                    End If
                End If
            Next X

            ''=============================== MAP WRAPPING PASS ===============================
            'Dim lMaxQuads As Int32 = (mlQuads - 1) * 10
            'Dim lTempX As Int32 = -1
            'Dim bXGrtr As Boolean = False
            'If lCurX - lMaxQuadDist < 0 Then
            '    'Ok, on left edge
            '    lTempX = (lCurX - lMaxQuadDist) + lMaxQuads
            '    'Now, check for QuadX > lTempX
            '    bXGrtr = True
            'ElseIf lCurX + lMaxQuadDist >= lMaxQuads Then
            '    'Ok, on right edge
            '    lTempX = (lCurX + lMaxQuadDist) - lMaxQuads
            '    'Now, check for QuadX < lTempX
            '    bXGrtr = False
            'End If

            ''Ok, now, go back and render the quads we didn't have before...
            'Dim oMatWorld As Matrix = Matrix.Identity
            'If lTempX <> -1 Then
            '    oMatWorld = Matrix.Identity
            '    'Two passes... this one is for missed Left/Right edges
            '    If lTempX <> -1 Then
            '        If bXGrtr = True Then
            '            'On left edge, so set our world matrix to transpose all items -X where X = CellSpacing * Width
            '            oMatWorld.Multiply(Matrix.Translation((Me.CellSpacing * TerrainClass.Width * -1) + Me.CellSpacing, 0, 0))
            '        Else
            '            'On right edge
            '            oMatWorld.Multiply(Matrix.Translation((Me.CellSpacing * TerrainClass.Width) - Me.CellSpacing, 0, 0))
            '        End If

            '        moDevice.Transform.World = oMatWorld
            '        If muSettings.BumpMapTerrain = True AndAlso moShader Is Nothing = False Then moShader.UpdateMatrices()

            '        For X As Int32 = 0 To (mlQuads * mlQuads) - 1
            '            'Use the old method of culling, for now...
            '            If (bXGrtr = True AndAlso goCamera.CullObject(mcbQB_East(X)) = False) OrElse (bXGrtr = False AndAlso goCamera.CullObject(mcbQB_West(X)) = False) Then
            '                lQuadsRendered += 1
            '                myQuadRendered(X) = 2
            '                moDevice.SetStreamSource(0, moQuadVB(X), 0)

            '                If mlQuadX(X) = 230 Then
            '                    moDevice.Indices = moQuad_RIB
            '                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, lTmpVT_XOrZ, 0, mlTris_XOrZ)
            '                    moDevice.Indices = moQuadIB
            '                ElseIf mlQuadZ(X) = 230 Then
            '                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, mlTris_XOrZ, 0, mlTris_XOrZ)
            '                Else
            '                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, lTmpVT, 0, mlTris)
            '                End If
            '            End If
            '        Next X
            '    End If

            '    moDevice.Transform.World = Matrix.Identity
            'End If
            '=============================== END OF MAP WRAPPING =============================
        Catch
        End Try


        'moDevice.RenderState.FillMode = FillMode.Solid
        Return lQuadsRendered
    End Function

#Region "  New Water Rendering  "
    Private mvbWaterQuads() As VertexBuffer
    Private mibWaterQuads() As IndexBuffer
    Private mlWaterQuadVerts() As Int32
    Private mlWaterQuadPrimCnt() As Int32
    Private mbQuadRendered() As Boolean

    Private Sub GenerateWaterGeography()
        'ok, we'll use our heightmap...

        Dim yMap() As Byte
        Dim iMapVertIdx() As Int16
        ReDim yMap(HeightMap.GetUpperBound(0))
        ReDim iMapVertIdx(HeightMap.GetUpperBound(0))

        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                Dim lIdx As Int32 = Y * Width + X
                iMapVertIdx(lIdx) = -1
                If HeightMap(lIdx) < WaterHeight Then
                    yMap(lIdx) = 1
                End If
            Next X
        Next Y

        'Now, go back through and any 1's, mark the neighboring tiles that are not 1 to a 2
        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                Dim lIdx As Int32 = Y * Width + X
                If yMap(lIdx) = 1 Then
                    For lSubX As Int32 = -1 To 1
                        For lSubY As Int32 = -1 To 1
                            Dim lTempX As Int32 = X + lSubX
                            Dim lTempY As Int32 = Y + lSubY
                            If lTempX < 0 Then lTempX = Width - 1
                            If lTempX > Width - 1 Then lTempX = 0
                            If lTempY < 0 Then Continue For
                            If lTempY > Height - 1 Then Continue For

                            Dim lTempIdx As Int32 = lTempY * Width + lTempX
                            If yMap(lTempIdx) <> 1 Then yMap(lTempIdx) = 2
                        Next lSubY
                    Next lSubX
                End If
            Next X
        Next Y

        'Determine groups of water... to do so, we check where water stops... that makes a group...
        '  We also divide the water up based on QUAD for cull checking
        'The biggest a group of water can get is the size of a quad...

        'Create our vertex and index buffers...
        Device.IsUsingEventHandlers = False
        ReDim mvbWaterQuads(mlQuads * mlQuads - 1)
        ReDim mibWaterQuads(mlQuads * mlQuads - 1)
        ReDim mbQuadRendered(mlQuads * mlQuads - 1)
        ReDim mlWaterQuadVerts(mlQuads * mlQuads - 1)
        ReDim mlWaterQuadPrimCnt(mlQuads * mlQuads - 1)


        Dim bQuadHit(mlQuads * mlQuads - 1) As Boolean
        For X As Int32 = 0 To bQuadHit.GetUpperBound(0)
            bQuadHit(X) = False
        Next


        For lQuadZ As Int32 = 0 To mlQuads - 1
            For lQuadX As Int32 = 0 To mlQuads - 1
                Dim lMinX As Int32 = mlQuadX(lQuadX)
                Dim lMaxX As Int32 = mlQuadX(lQuadX) + mlVertsPerQuad '- 1
                Dim lMinZ As Int32 = lQuadZ * mlVertsPerQuad 'mlQuadZ(lQuadZ)
                Dim lMaxZ As Int32 = (lQuadZ + 1) * mlVertsPerQuad '- 1

                If lMaxX > Width - 1 Then lMaxX = Width - 1
                If lMaxZ > Height - 1 Then lMaxZ = Height - 1

                'First, determine the number of verts
                Dim lVertCnt As Int32 = 0
                For lVertX As Int32 = lMinX To lMaxX
                    For lVertZ As Int32 = lMinZ To lMaxZ
                        Dim lIdx As Int32 = lVertZ * Width + lVertX
                        If yMap(lIdx) <> 0 Then
                            'ok, a vert...
                            lVertCnt += 1
                        End If
                    Next lVertZ
                Next lVertX


                'Now, with that, let's make our verts
                Dim lQuadIndex As Int32 = lQuadZ * mlQuads + lQuadX
                mlWaterQuadPrimCnt(lQuadIndex) = 0
                If lVertCnt = 0 Then Continue For

                Dim uMattWVerts(lVertCnt - 1) As CustomVertex.PositionTextured

                Dim utempvert As CustomVertex.PositionTextured
                If bQuadHit(lQuadIndex) = True Then Stop
                bQuadHit(lQuadIndex) = True
                mvbWaterQuads(lQuadIndex) = New VertexBuffer(utempvert.GetType, lVertCnt, moDevice, Usage.WriteOnly, CustomVertex.PositionTextured.Format, Pool.Managed)
                mlWaterQuadVerts(lQuadIndex) = lVertCnt

                'Now... let's make our verts...
                Dim lVertIdx As Int32 = -1
                For lVertX As Int32 = lMinX To lMaxX
                    For lVertZ As Int32 = lMinZ To lMaxZ
                        Dim lIdx As Int32 = lVertZ * Width + lVertX
                        If yMap(lIdx) <> 0 Then
                            'ok, a vert...
                            Dim lVertLocX As Int32 = (lVertX - (Width \ 2)) * Me.CellSpacing
                            Dim lVertLocZ As Int32 = (lVertZ - (Height \ 2)) * Me.CellSpacing

                            Dim fTU As Single
                            Dim fTV As Single

                            'If lQuadX Mod 2 = 0 Then
                            fTU = CSng(lVertX - lMinX) / (lMaxX - lMinX)
                            'Else : fTU = 1 - (CSng(lVertX - lMinX) / lMaxX)
                            'End If
                            'If lQuadZ Mod 2 = 0 Then
                            fTV = CSng(lVertZ - lMinZ) / (lMaxZ - lMinZ)
                            'Else : fTV = 1 - (CSng(lVertZ - lMinZ) / lMaxZ)
                            'End If

                            lVertIdx += 1
                            uMattWVerts(lVertIdx) = New CustomVertex.PositionTextured(lVertLocX, 0, lVertLocZ, fTU, fTV)

                            iMapVertIdx(lIdx) = CShort(lVertIdx)
                        End If
                    Next lVertZ
                Next lVertX

                'Ok, we have all of our verts... now, our indices...
                'loop through from top to bottom
                Dim iWaterIndices() As Int16 = Nothing
                Dim lWaterIndicesUB As Int32 = -1
                Dim lPrimCnt As Int32 = 0
                For lVertX As Int32 = lMinX To lMaxX
                    For lVertZ As Int32 = lMinZ To lMaxZ
                        Dim lIdx As Int32 = lVertZ * Width + lVertX
                        If yMap(lIdx) <> 0 Then
                            'Ok, now, we found a valid vert, now, first, check below right... if it is not there, we are done...
                            Dim lTempX As Int32 = lVertX + 1
                            Dim lTempZ As Int32 = lVertZ + 1

                            If lTempZ > Height - 1 Then Continue For
                            If lTempX > Width - 1 Then lTempX = 0
                            If lTempX > lMaxX Then Continue For
                            If lTempZ > lMaxZ Then Continue For
                            Dim lTempIdx As Int32 = lTempZ * Width + lTempX
                            Dim lBottomRightIdx As Int32 = lTempIdx
                            If yMap(lTempIdx) = 0 Then Continue For

                            'ok, below right is good we have scenario that may required indices... we'll check our ABC first... below me
                            lTempX = lVertX
                            lTempIdx = lTempZ * Width + lTempX
                            If yMap(lTempIdx) <> 0 Then
                                'ok, got an ABC here...
                                ReDim Preserve iWaterIndices(lWaterIndicesUB + 3)
                                'A - me
                                lWaterIndicesUB += 1 : iWaterIndices(lWaterIndicesUB) = iMapVertIdx(lIdx)
                                'B - below me
                                lWaterIndicesUB += 1 : iWaterIndices(lWaterIndicesUB) = iMapVertIdx(lTempIdx)
                                'C - bottom right
                                lWaterIndicesUB += 1 : iWaterIndices(lWaterIndicesUB) = iMapVertIdx(lBottomRightIdx)

                                lPrimCnt += 1
                            End If

                            'Now, check for a ACD scenario... which is to my right
                            lTempX = lVertX + 1
                            lTempZ = lVertZ
                            lTempIdx = lTempZ * Width + lTempX
                            If yMap(lTempIdx) <> 0 Then
                                'ok, got an ACD here...
                                ReDim Preserve iWaterIndices(lWaterIndicesUB + 3)
                                'A - me
                                lWaterIndicesUB += 1 : iWaterIndices(lWaterIndicesUB) = iMapVertIdx(lIdx)
                                'C - bottom right
                                lWaterIndicesUB += 1 : iWaterIndices(lWaterIndicesUB) = iMapVertIdx(lBottomRightIdx)
                                'D - right
                                lWaterIndicesUB += 1 : iWaterIndices(lWaterIndicesUB) = iMapVertIdx(lTempIdx)

                                lPrimCnt += 1
                            End If
                        End If
                    Next lVertZ
                Next lVertX

                mlWaterQuadPrimCnt(lQuadIndex) = lPrimCnt
                If lPrimCnt = 0 Then
                    mvbWaterQuads(lQuadIndex).Dispose()
                    mvbWaterQuads(lQuadIndex) = Nothing
                    Continue For
                End If

                'Now, we have our vertices for this quad, and we have our indices to render this quad... so we're done..

                mvbWaterQuads(lQuadIndex).SetData(uMattWVerts, 0, LockFlags.None)
                Dim iTempShort As Int16 = 0
                mibWaterQuads(lQuadIndex) = New IndexBuffer(iTempShort.GetType, iWaterIndices.Length, moDevice, Usage.WriteOnly, Pool.Managed)
                mibWaterQuads(lQuadIndex).SetData(iWaterIndices, 0, LockFlags.None)

            Next lQuadX
        Next lQuadZ
        Device.IsUsingEventHandlers = True

    End Sub

    Private Sub RenderWaterNew(ByVal bShader2 As Boolean)
        If mlWaterQuadPrimCnt Is Nothing Then GenerateWaterGeography()

        If bShader2 = True Then
            If moOcean Is Nothing Then moOcean = New OceanRender(moDevice)
            If Me.MapType <> moOcean.yCurrentPlanetType OrElse Me.CellSpacing <> moOcean.lCurrentCellSpacing Then moOcean.SetAsPlanetType(CByte(Me.MapType), Me.CellSpacing)
            moOcean.fWaterHeight = CSng(Me.WaterHeight) * Me.ml_Y_Mult

            '0 to 1, at .5, the planet is mid-day
            'If Me.MapType <> PlanetType.eGeoPlastic Then
            Dim fTemp As Single = 0.5F
            fTemp = Math.Abs(0.5F - fTemp)
            fTemp /= 0.5F
            fTemp = 1.0F - fTemp
            fTemp -= 0.2F
            If fTemp < 0 Then fTemp = 0.0F Else fTemp = (fTemp / 0.8F) * 2.0F
            fTemp = Math.Max(fTemp, 0.1F)
            moOcean.SetWaterClrStr(fTemp)

            If moDevice.Lights.Count > 0 Then
                Dim uLt As Light
                uLt = moDevice.Lights(0)

                Dim vecTemp As Vector3 = uLt.Direction
                If vecTemp.Y < 0 Then
                    vecTemp.Y *= -1.0F
                    vecTemp.Z = 0
                    vecTemp.Normalize()

                    'Now...
                    vecTemp.Z = Math.Abs(vecTemp.X) + Math.Abs(vecTemp.Y)
                    vecTemp.Normalize()

                    moOcean.SetFromLightVec(vecTemp)

                    moOcean.AttenuateSpecular(Math.Abs((vecTemp.Y * 2.0F) / 2.0F))
                Else
                    vecTemp = New Vector3(0, 0, 0)
                    moOcean.SetFromLightVec(vecTemp)
                    moOcean.AttenuateSpecular(0)
                End If
            End If
            'End If

            'ok, now to do our rendering
            Dim lPasses As Int32 = moOcean.PrepareEffect()
            For lPass As Int32 = 0 To lPasses - 1
                moOcean.StartPass(lPass)

                For lQuadIdx As Int32 = 0 To (mlQuads * mlQuads - 1)
                    If mlWaterQuadPrimCnt(lQuadIdx) > 0 Then
                        moDevice.SetStreamSource(0, mvbWaterQuads(lQuadIdx), 0)
                        moDevice.Indices = mibWaterQuads(lQuadIdx)
                        moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, mlWaterQuadVerts(lQuadIdx), 0, mlWaterQuadPrimCnt(lQuadIdx))
                    End If
                Next lQuadIdx

                moOcean.StopPass()
            Next lPass
            moOcean.StopEffect()

        Else
            Dim matWorld As Matrix = Matrix.Identity
            matWorld.Multiply(Matrix.Translation(0, WaterHeight * ml_Y_Mult, 0))
            moDevice.Transform.World = matWorld

            moDevice.SetTexture(0, Nothing)
            moDevice.SetTexture(1, Nothing)
            moDevice.Material = muWaterMat
            For lQuadIdx As Int32 = 0 To (mlQuads * mlQuads - 1)
                If mbQuadRendered(lQuadIdx) = True AndAlso mlWaterQuadPrimCnt(lQuadIdx) > 0 Then
                    moDevice.SetStreamSource(0, mvbWaterQuads(lQuadIdx), 0)
                    moDevice.Indices = mibWaterQuads(lQuadIdx)
                    moDevice.DrawIndexedPrimitives(PrimitiveType.TriangleList, 0, 0, mlWaterQuadVerts(lQuadIdx), 0, mlWaterQuadPrimCnt(lQuadIdx))
                End If
            Next lQuadIdx
        End If
    End Sub

#End Region

    Public Function GetHeightAtLocation(ByVal fLocX As Single, ByVal fLocZ As Single) As Single
        Dim fTX As Single   'translated X 
        Dim fTZ As Single   'translated Z 
        Dim lCol As Int32
        Dim lRow As Int32

        Dim yA As Int32
        Dim yB As Int32
        Dim yC As Int32
        Dim yD As Int32

        Dim dZ As Single
        Dim dX As Single

        Dim lIdx As Int32

        If HeightMap Is Nothing Then Return 0

        fTX = mlHalfWidth + (fLocX / CellSpacing)
        fTZ = mlHalfHeight + (fLocZ / CellSpacing)

        lCol = CInt(Math.Floor(fTX))
        lRow = CInt(Math.Floor(fTZ))
        dX = fTX - lCol
        dZ = fTZ - lRow

        lIdx = (lRow * Width) + lCol
        If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yA = 0 Else yA = HeightMap(lIdx)
        lIdx = (lRow * Width) + lCol + 1
        If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yB = 0 Else yB = HeightMap(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol
        If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yC = 0 Else yC = HeightMap(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol + 1
        If lIdx < 0 Or lIdx > HeightMap.Length - 1 Then yD = 0 Else yD = HeightMap(lIdx)

        Dim fV1 As Single = yA * (1 - dX) + yB * dX
        Dim fV2 As Single = yC * (1 - dX) + yD * dX
        Return Math.Max((fV1 * (1 - dZ) + fV2 * dZ) * ml_Y_Mult, WaterHeight)
    End Function

    Public Function GetTerrainNormal(ByVal fLocX As Single, ByVal fLocZ As Single) As Vector3
        Dim fTX As Single   'translated X 
        Dim fTZ As Single   'translated Z 
        Dim lCol As Int32
        Dim lRow As Int32

        Dim dZ As Single
        Dim dX As Single

        Dim vN1 As Vector3
        Dim vN2 As Vector3
        Dim vN3 As Vector3
        Dim vN4 As Vector3

        Dim lIdx As Int32

        fTX = mlHalfWidth + (fLocX / CellSpacing)
        fTZ = mlHalfHeight + (fLocZ / CellSpacing)

        lCol = CInt(Math.Floor(fTX))
        lRow = CInt(Math.Floor(fTZ))
        dX = fTX - lCol
        dZ = fTZ - lRow

        lIdx = (lRow * Width) + lCol
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN1 = Vector3.Empty Else vN1 = HMNormals(lIdx)
        lIdx = (lRow * Width) + lCol + 1
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN2 = Vector3.Empty Else vN2 = HMNormals(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN3 = Vector3.Empty Else vN3 = HMNormals(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol + 1
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN4 = Vector3.Empty Else vN4 = HMNormals(lIdx)

        vN1.Multiply(1 - dX)
        vN2.Multiply(dX)
        Dim vA As Vector3 = Vector3.Add(vN1, vN2)
        vN3.Multiply(1 - dX)
        vN4.Multiply(dX)
        Dim vB As Vector3 = Vector3.Add(vN3, vN4)
        vA.Multiply(1 - dZ)
        vB.Multiply(dZ)
        Dim vF As Vector3 = Vector3.Add(vA, vB)
        vF.Normalize()

        Return vF
    End Function


    Public Sub CleanResources()
        'A call to this sub means that the program wants to release all but the bare essential data...
        mbHMReady = False
        Erase HeightMap

        'MSC - 06/05/08 - remarked out for bumpmapping
        'moVB = Nothing
        ClearShader()

        Erase moVerts
        Erase moVertsTBN

        mvbWater = Nothing
        Erase muWaterVerts

        Erase HMNormals
        Erase mlQuadFOWZIndex
        Erase mlQuadFOWXIndex
        Erase mlQuadStartIndex
        Erase mlQuadX
        Erase mlQuadZ
        If moVolTex Is Nothing = False Then moVolTex.Dispose() 'we do this because it is a SCRATCH surface
        moVolTex = Nothing
        If moTex Is Nothing = False Then moTex.Dispose() 'we do this becuase it is a SCRATCH surface
        moTex = Nothing
        If moNormalVol Is Nothing = False Then moNormalVol.Dispose()
        moNormalVol = Nothing
        If moIllum Is Nothing = False Then moIllum.Dispose()
        moIllum = Nothing
    End Sub

    Public Sub CreateTerrainPlanetTexture(ByRef oTex As Texture)
        'render it with older (uglier) technique
        oTex = New Texture(moDevice, 256, 256, 1, Usage.None, Format.A8R8G8B8, Pool.Managed)
        Dim pitch As Int32
        Dim yDest As Byte() = CType(oTex.LockRectangle(GetType(Byte), 0, 0, pitch, 256 * 256 * 4), Byte())
        Dim X As Int32
        Dim Y As Int32
        Dim lVal As Int32
        Dim lR As Int32
        Dim lG As Int32
        Dim lB As Int32

        Dim lMapX As Int32
        Dim lMapY As Int32

        Dim lWaterColor As Int32

        'Regardless of method, we must generate the terrain
        If mbHMReady = False Then GenerateTerrain(mlSeed)

        'set our water color...
        'Select Case MapType
        '    Case PlanetType.eAcidic
        '        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 128, 48).ToArgb
        '    Case PlanetType.eAdaptable
        '        lWaterColor = System.Drawing.Color.FromArgb(120, 0, 192, 255).ToArgb
        '    Case PlanetType.eBarren
        '        lWaterColor = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb
        '    Case PlanetType.eDesert
        '        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 64, 128).ToArgb
        '    Case PlanetType.eGeoPlastic
        '        lWaterColor = System.Drawing.Color.FromArgb(200, 255, 48, 0).ToArgb
        '    Case PlanetType.eTerran
        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 64, 128).ToArgb
        '    Case PlanetType.eTundra
        '        lWaterColor = System.Drawing.Color.FromArgb(150, 0, 192, 192).ToArgb
        '    Case PlanetType.eWaterWorld
        '        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 128, 196).ToArgb
        'End Select

        'Use 255 for purposes of texture mapping
        For Y = 0 To 255
            lMapY = CInt(Math.Floor((Y / 256.0F) * 240.0F))
            For X = 0 To 255
                lMapX = CInt(Math.Floor((X / 256.0F) * 240.0F))
                lVal = HeightMap(lMapY * Width + lMapX)

                If lVal < WaterHeight Then
                    lR = System.Drawing.Color.FromArgb(lWaterColor).R + CInt((lVal / WaterHeight) * 10) - CInt(Rnd() * 5)
                    lG = System.Drawing.Color.FromArgb(lWaterColor).G + CInt((lVal / WaterHeight) * 10) - CInt(Rnd() * 5)
                    lB = System.Drawing.Color.FromArgb(lWaterColor).B + CInt((lVal / WaterHeight) * 10) - CInt(Rnd() * 5)

                    If lR > 255 Then lR = 255
                    If lR < 0 Then lR = 0
                    If lG > 255 Then lG = 255
                    If lG < 0 Then lG = 0
                    If lB > 255 Then lB = 255
                    If lB < 0 Then lB = 0
                Else
                    Call GetHeightMapColor(lVal, lR, lG, lB)
                End If

                yDest((Y * 256 + X) * 4 + 3) = 255  'A
                yDest((Y * 256 + X) * 4 + 2) = CByte(lR) 'R
                yDest((Y * 256 + X) * 4 + 1) = CByte(lG) 'G
                yDest((Y * 256 + X) * 4 + 0) = CByte(lB) 'B
            Next X
        Next Y

        oTex.UnlockRectangle(0)

        Erase yDest

    End Sub

    Private Sub SetMatsTexsAndGetWaterColor()
        Dim lWaterColor As Int32
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        Dim lWaterA As Int32
        Dim sTerrName As String = ""    'volume texture terrain name
        Dim sNVTerrName As String = ""  'no volume texture terrain name
        Dim lWaterS As Int32 = System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb

        'Select Case MapType
        '    Case PlanetType.eAcidic
        '        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 64, 20).ToArgb
        '        lWaterA = System.Drawing.Color.FromArgb(230, 0, 64, 0).ToArgb
        '        sTerrName = "AcidSet.dds"
        '        sNVTerrName = "AcidNV.bmp"
        '    Case PlanetType.eAdaptable
        '        lWaterColor = System.Drawing.Color.FromArgb(120, 0, 64, 255).ToArgb
        '        lWaterA = System.Drawing.Color.FromArgb(120, 0, 32, 72).ToArgb
        '        sTerrName = "AdaptableSet.dds"
        '        sNVTerrName = "AdaptableNV.bmp"
        '    Case PlanetType.eBarren
        '        lWaterColor = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb
        '        lWaterA = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb
        '        lWaterS = System.Drawing.Color.FromArgb(0, 0, 0, 0).ToArgb
        '        sTerrName = "BarrenSet.dds"
        '        sNVTerrName = "BarrenNV.bmp"
        '    Case PlanetType.eDesert
        '        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 0, 48).ToArgb
        '        lWaterA = System.Drawing.Color.FromArgb(230, 0, 0, 24).ToArgb
        '        sTerrName = "DesertSet.dds"
        '        sNVTerrName = "DesertNV.bmp"
        '    Case PlanetType.eGeoPlastic
        '        lWaterColor = System.Drawing.Color.FromArgb(230, 255, 0, 0).ToArgb
        '        lWaterA = System.Drawing.Color.FromArgb(230, 96, 48, 0).ToArgb      '???
        '        lWaterS = System.Drawing.Color.FromArgb(255, 64, 0, 0).ToArgb
        '        sTerrName = "GeoPlasticSet.dds"
        '        sNVTerrName = "GeoPlasticNV.bmp"
        '    Case PlanetType.eTerran
        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 24, 128).ToArgb
        lWaterA = System.Drawing.Color.FromArgb(230, 0, 0, 24).ToArgb
        sTerrName = "TerranSet.dds"
        sNVTerrName = "TerranNV.bmp"
        '    Case PlanetType.eTundra
        '        lWaterColor = System.Drawing.Color.FromArgb(120, 0, 192, 220).ToArgb
        '        lWaterA = System.Drawing.Color.FromArgb(120, 0, 64, 72).ToArgb
        '        sTerrName = "TundraSet.dds"
        '        sNVTerrName = "TundraNV.bmp"
        '    Case PlanetType.eWaterWorld
        '        lWaterColor = System.Drawing.Color.FromArgb(230, 0, 48, 96).ToArgb
        '        lWaterA = System.Drawing.Color.FromArgb(230, 0, 24, 48).ToArgb
        '        lWaterS = System.Drawing.Color.FromArgb(255, 0, 192, 255).ToArgb
        '        sTerrName = "WaterworldSet.dds"
        '        sNVTerrName = "WaterworldNV.bmp"
        'End Select

        With muWaterMat
            .Diffuse = System.Drawing.Color.FromArgb(lWaterColor)
            .Ambient = System.Drawing.Color.FromArgb(lWaterA)
            .Specular = System.Drawing.Color.FromArgb(lWaterS)

            'If MapType <> PlanetType.eGeoPlastic Then
            .Emissive = System.Drawing.Color.Black
            'Else
            '    '.Emissive = System.Drawing.Color.FromArgb(lWaterA)
            '    .Emissive = .Specular
            'End If

            .SpecularSharpness = 10
        End With

        With moTerrMat
            .Diffuse = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            'If MapType <> PlanetType.eGeoPlastic Then
            .Ambient = System.Drawing.Color.FromArgb(255, 192, 192, 192)
            'Else
            '    .Ambient = System.Drawing.Color.FromArgb(255, 255, 128, 128)
            'End If
            .Emissive = System.Drawing.Color.Black

            'If MapType = PlanetType.eTundra Then
            '    .Specular = Color.White
            '    .SpecularSharpness = 2
            'Else
            .Specular = System.Drawing.Color.Black
            .SpecularSharpness = 0
            'End If
        End With

        'The resource manager checks this too, but I want to be efficient, so I'll check it here to be sure...
        'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        'If sPath.EndsWith("\") = False Then sPath &= "\"
        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap Then
            '    moVolTex = TextureLoader.FromVolumeFile(moDevice, sPath & sTerrName, 0, 0, 0, 1, Usage.None, Format.X8R8G8B8, Pool.Managed, Filter.Linear, Filter.Linear, 0)
            moVolTex = moResMgr.GetVolTerrainTexture(sTerrName)
        Else
            moTex = moResMgr.GetNVTerrainTexture(sNVTerrName)
        End If

    End Sub

    Private Function GetTWTextureCoord(ByVal lHeight As Int32) As Single
        Dim fTw As Single
        Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

        'Select Case MapType
        '    Case PlanetType.eAcidic
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.01)
        '                fTw = CSng(lHeight / (WaterHeight + (lAboveWaterHeight * 0.01)))
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.1)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.01), WaterHeight + CInt(lAboveWaterHeight * 0.1), lHeight)
        '                fTw += 1
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.1), 255, lHeight)
        '                fTw *= 6
        '                fTw += 2
        '        End Select
        '    Case PlanetType.eAdaptable
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.08)
        '                fTw = CSng(lHeight / (WaterHeight + (lAboveWaterHeight * 0.08)))
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.08), 255, lHeight)
        '                fTw *= 7
        '                fTw += 1
        '        End Select
        '    Case PlanetType.eBarren
        '        fTw = lHeight / 255.0F
        '        fTw *= 8
        '    Case PlanetType.eDesert
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.08)
        '                fTw = CSng(lHeight / (WaterHeight + (lAboveWaterHeight * 0.08)))
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.08), 255, lHeight)
        '                fTw *= 7
        '                fTw += 1
        '        End Select
        '    Case PlanetType.eGeoPlastic
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.01)
        '                fTw = CSng(lHeight / (WaterHeight + (lAboveWaterHeight * 0.01)))
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.15)
        '                fTw = GetRangedTexCoord((WaterHeight + CInt(lAboveWaterHeight * 0.01)), (WaterHeight + CInt(lAboveWaterHeight * 0.15)), lHeight)
        '                fTw += 1
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.3)
        '                fTw = GetRangedTexCoord((WaterHeight + CInt(lAboveWaterHeight * 0.15)), (WaterHeight + CInt(lAboveWaterHeight * 0.3)), lHeight)
        '                fTw += 2
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.45)
        '                fTw = GetRangedTexCoord((WaterHeight + CInt(lAboveWaterHeight * 0.3)), (WaterHeight + CInt(lAboveWaterHeight * 0.45)), lHeight)
        '                fTw += 3
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.6)
        '                fTw = GetRangedTexCoord((WaterHeight + CInt(lAboveWaterHeight * 0.45)), (WaterHeight + CInt(lAboveWaterHeight * 0.6)), lHeight)
        '                fTw += 4
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.75)
        '                fTw = GetRangedTexCoord((WaterHeight + CInt(lAboveWaterHeight * 0.6)), (WaterHeight + CInt(lAboveWaterHeight * 0.75)), lHeight)
        '                fTw += 5
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.9)
        '                fTw = GetRangedTexCoord((WaterHeight + CInt(lAboveWaterHeight * 0.75)), (WaterHeight + CInt(lAboveWaterHeight * 0.9)), lHeight)
        '                fTw += 6
        '            Case Else
        '                fTw = GetRangedTexCoord((WaterHeight + CInt(lAboveWaterHeight * 0.9)), 255, lHeight)
        '                fTw += 7
        '        End Select
        '    Case PlanetType.eTerran
        Select Case lHeight
            Case Is < WaterHeight
                'sand (underwater)
                fTw = CSng(lHeight / WaterHeight)
            Case Is < CInt(lAboveWaterHeight * 0.05F) + WaterHeight
                'sand
                fTw = GetRangedTexCoord(WaterHeight, CInt(lAboveWaterHeight * 0.05F) + WaterHeight, lHeight)
                fTw += 1
            Case Is < CInt(lAboveWaterHeight * 0.2F) + WaterHeight
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.05F) + WaterHeight, CInt(lAboveWaterHeight * 0.2F) + WaterHeight, lHeight)
                fTw += 2
            Case Is < CInt(lAboveWaterHeight * 0.45F) + WaterHeight
                'grass                        
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.2F) + WaterHeight, CInt(lAboveWaterHeight * 0.45F) + WaterHeight, lHeight)
                fTw += 3
            Case Is < CInt(lAboveWaterHeight * 0.53F) + WaterHeight
                'grass red rock
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.45F) + WaterHeight, CInt(lAboveWaterHeight * 0.53F) + WaterHeight, lHeight)
                fTw += 4
            Case Is < CInt(lAboveWaterHeight * 0.61F) + WaterHeight
                'red rock rock
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.53F) + WaterHeight, CInt(lAboveWaterHeight * 0.61F) + WaterHeight, lHeight)
                fTw += 5
            Case Is < CInt(lAboveWaterHeight * 0.85F) + WaterHeight
                'rock
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.61F) + WaterHeight, CInt(lAboveWaterHeight * 0.85F) + WaterHeight, lHeight)
                fTw += 6
            Case Else
                'snow
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.85F) + WaterHeight, 255, lHeight)
                fTw += 7
        End Select
        '    Case PlanetType.eTundra
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.1)
        '                fTw = CSng(lHeight / (WaterHeight + (lAboveWaterHeight * 0.1)))
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.1), 255, lHeight)
        '                fTw *= 7
        '                fTw += 1
        '        End Select
        '    Case PlanetType.eWaterWorld
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.1)
        '                fTw = CSng(lHeight / (WaterHeight + (lAboveWaterHeight * 0.1)))
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.1), 255, lHeight)
        '                fTw *= 7
        '                fTw += 1
        '        End Select
        'End Select

        Return fTw / 8
    End Function

    Private Sub GenerateTerrain(ByVal lSeed As Int32)
        Dim lWidthSpans() As Int32
        Dim lHeightSpans() As Int32
        Dim X As Int32
        Dim lVal As Int32
        Dim lWaterPerc As Int32
        Dim bDone As Boolean
        Dim bLandmassLoop As Boolean
        Dim lStartX As Int32
        Dim lStartY As Int32
        Dim Y As Int32
        Dim lIdx As Int32
        Dim lPass As Int32

        Dim lSubXMax As Int32
        Dim lSubYMax As Int32
        Dim lSubX As Int32
        Dim lSubY As Int32

        Dim lTemp As Int32

        Dim lTotalLand As Int32

        Dim HM_Type() As Byte

        Call Rnd(-1)
        Randomize(lSeed)

        'fill our spans
        ReDim lWidthSpans(mlPASSES - 1)
        For X = 0 To mlPASSES - 1
            lVal = CInt(2 ^ X)
            If lVal > (Width / mlPASSES) Then
                lVal = lWidthSpans(X - 1) + CInt((Width / mlPASSES) * (X / mlPASSES))
            End If
            lWidthSpans(X) = lVal
        Next X
        ReDim lHeightSpans(mlPASSES - 1)
        For X = 0 To mlPASSES - 1
            lVal = CInt(2 ^ X)
            If lVal > (Height / mlPASSES) Then
                lVal = lHeightSpans(X - 1) + CInt((Height / mlPASSES) * (X / mlPASSES))
            End If
            lHeightSpans(X) = lVal
        Next X

        ReDim HeightMap(Width * Height)
        ReDim HM_Type(Width * Height)

        'Get our map type
        'MapType = -1
        'lTemp = Int(Rnd() * 100) + 1
        'Select Case lTemp
        '    Case Is < me_Planet_Type.ePT_Acid
        '        MapType = me_Planet_Type.ePT_Acid
        '    Case Is < me_Planet_Type.ePT_Barren
        '        MapType = me_Planet_Type.ePT_Barren
        '    Case Is < me_Planet_Type.ePT_GeoPlastic
        '        MapType = me_Planet_Type.ePT_GeoPlastic
        '    Case Is < me_Planet_Type.ePT_Desert
        '        MapType = me_Planet_Type.ePT_Desert
        '    Case Is < me_Planet_Type.ePT_Adaptable
        '        MapType = me_Planet_Type.ePT_Adaptable
        '    Case Is < me_Planet_Type.ePT_Tundra
        '        MapType = me_Planet_Type.ePT_Tundra
        '    Case Is < me_Planet_Type.ePT_Terran
        '        MapType = me_Planet_Type.ePT_Terran
        '    Case Else
        '        MapType = me_Planet_Type.ePT_Waterworld
        'End Select

        lTemp = CInt(Int(Rnd() * 100) + 1)


        'Get our water percentage based on map type
        lWaterPerc = GetWaterPerc(MapType)

        'Now, set our values to water unless there is none
        If lWaterPerc = 0 Then
            'set em all to land
            WaterHeight = 0
            For X = 0 To HeightMap.Length - 1
                HeightMap(X) = 100
                HM_Type(X) = 1
            Next X
        Else
            For X = 0 To HeightMap.Length - 1
                HeightMap(X) = 0
                HM_Type(X) = 0
            Next X

            'Base the waterheight off of the waterperc, if our water perc is practically nothing, then
            '  we don't want waterheight to be 256 :)
            'WaterHeight = CByte(Int(Rnd() * (lWaterPerc + 20)) + 1)

            'If MapType = PlanetType.eTerran Then
            '    WaterHeight = 70
            'ElseIf MapType = PlanetType.eWaterWorld Then
            '    WaterHeight = 160
            WaterHeight = 40
            'End If

            'Generate landmasses
            bDone = False
            While bDone = False
                lStartX = CInt(Int(Rnd() * Width) + 1)
                lStartY = CInt(Int(Rnd() * Height) + 1)
                X = lStartX
                Y = lStartY

                bLandmassLoop = False
                While bLandmassLoop = False
                    'Get our next movement...
                    Select Case Int(Rnd() * 8) + 1
                        Case 1 : Y -= 1
                        Case 2 : Y -= 1 : X += 1
                        Case 3 : X += 1
                        Case 4 : Y += 1 : X += 1
                        Case 5 : Y += 1
                        Case 6 : Y += 1 : X -= 1
                        Case 7 : X -= 1
                        Case 8 : X -= 1 : Y -= 1
                    End Select
                    'Validate ranges
                    If X < 0 Then X = Width - 1
                    If X >= Width Then X = 0
                    If Y < 0 Then Y = 1
                    If Y >= Height Then Y = Height - 2

                    lIdx = (Y * Width) + X
                    If HM_Type(lIdx) = 0 Then
                        HeightMap(lIdx) = CByte(WaterHeight + 32)      'to ensure that we are way above water
                        HM_Type(lIdx) = 1
                        lTotalLand += 1
                    End If
                    If X = lStartX AndAlso Y = lStartY Then
                        bLandmassLoop = True
                    ElseIf (lTotalLand / (Width * Height)) * 100 > 100 - lWaterPerc Then
                        bLandmassLoop = True
                        bDone = True
                    End If
                End While
            End While
        End If

        'Generate the terrain on the map...
        For lPass = mlPASSES - 1 To 0 Step -1
            For X = 0 To Width - 1 Step lWidthSpans(lPass)
                For Y = 0 To Height - 1 Step lHeightSpans(lPass)
                    lVal = CInt(Int(Rnd() * 255) + 1)

                    lSubXMax = X + lWidthSpans(lPass) - 1
                    If lSubXMax > Width Then lSubXMax = Width - 1
                    lSubYMax = Y + lHeightSpans(lPass) - 1
                    If lSubYMax > Height Then lSubYMax = Height - 1

                    For lSubX = X To lSubXMax
                        For lSubY = Y To lSubYMax
                            lIdx = lSubY * Width + lSubX
                            If HM_Type(lIdx) = 0 Then
                                lTemp = CInt((lVal / 255) * WaterHeight)
                            Else
                                lTemp = lVal
                            End If
                            HeightMap(lIdx) = SmoothTerrainVal(lSubX, lSubY, lTemp)
                        Next lSubY
                    Next lSubX
                Next Y
            Next X
        Next lPass

        'Now, assuming we are lunar...
        'If MapType = PlanetType.eBarren Then
        '    DoLunar()
        'End If

        'Now, apply final filter to smooth everything over
        For lPass = 1 To 0 Step -1
            For X = 0 To Width - 1 Step lWidthSpans(lPass)
                For Y = 0 To Height - 1 Step lHeightSpans(lPass)
                    lSubXMax = X + lWidthSpans(lPass) - 1
                    If lSubXMax > Width Then lSubXMax = Width - 1
                    lSubYMax = Y + lHeightSpans(lPass) - 1
                    If lSubYMax > Height Then lSubYMax = Height - 1

                    For lSubX = X To lSubXMax
                        For lSubY = Y To lSubYMax
                            lIdx = lSubY * Width + lSubX

                            lTemp = HeightMap(lIdx)
                            HeightMap(lIdx) = SmoothTerrainVal(lSubX, lSubY, lTemp)
                        Next lSubY
                    Next lSubX
                Next Y
            Next X
        Next lPass


        'If MapType = PlanetType.eGeoPlastic OrElse MapType = PlanetType.eWaterWorld Then
        '    CreateVolcanoes()
        'End If

        'Now, accentuate
        lTemp = 0 : lVal = 0
        For X = 0 To Width - 1
            For Y = 0 To Height - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lTemp Then lTemp = HeightMap(lIdx)
                If HeightMap(lIdx) > WaterHeight Then
                    lVal += HeightMap(lIdx)
                End If
            Next Y
        Next X
        lVal = (lVal \ (Width * Height))
        lVal = lVal + CInt((lTemp - lVal) / 1.8F)
        For X = 0 To Width - 1
            For Y = 0 To Height - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lVal Then HeightMap(lIdx) = CByte(HeightMap(lIdx) * (255 / lTemp))
            Next Y
        Next X

        'If MapType = PlanetType.eBarren Then
        '    For Y = 0 To Height - 1
        '        For X = 0 To Width - 1
        '            lIdx = Y * Width + X
        '            lTemp = HeightMap(lIdx)
        '            lTemp += (CInt(Rnd() * 20) - 10)
        '            If lTemp < 0 Then lTemp = 0
        '            If lTemp > 255 Then lTemp = 255
        '            HeightMap(lIdx) = CByte(lTemp)
        '        Next X
        '    Next Y
        '    MaximizeTerrain()
        'ElseIf MapType = PlanetType.eDesert Then
        '    DoDesert()

        '    'and an additional smooth over
        '    For X = 0 To Width - 1
        '        For Y = 0 To Height - 1
        '            lIdx = Y * Width + X
        '            HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
        '        Next Y
        '    Next X
        'ElseIf MapType = PlanetType.eAcidic Then
        '    DoAcidPlateaus()
        'ElseIf MapType = PlanetType.eTerran Then
        DoPeakAccents()
        'End If

        'One last soften...
        For X = 0 To Width - 1
            For Y = 0 To Height - 1
                lIdx = Y * Width + X
                HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
            Next Y
        Next X

        'If MapType = PlanetType.eGeoPlastic Then
        '    DoGeoPlastic()
        'ElseIf MapType = PlanetType.eAcidic Then
        '    DoAcidic()
        'ElseIf MapType = PlanetType.eWaterWorld Then
        '    If mptVolcanoes Is Nothing = False Then
        '        For X = 0 To mptVolcanoes.GetUpperBound(0)
        '            lIdx = mptVolcanoes(X).Y * Width + mptVolcanoes(X).X
        '            HeightMap(lIdx) = 0

        '            Dim vecVent As Vector3
        '            vecVent.X = (mptVolcanoes(X).X - (Width \ 2)) * CellSpacing
        '            vecVent.Y = WaterHeight * ml_Y_Mult
        '            vecVent.Z = (mptVolcanoes(X).Y - (Height \ 2)) * CellSpacing
        '            If moPFX Is Nothing Then moPFX = New PlanetFX.ParticleEngine(moDevice, Me)
        '            moPFX.AddEmitter(PlanetFX.ParticleEngine.EmitterType.eVolcano, vecVent, 400)
        '        Next
        '    End If
        '    MaximizeTerrain()
        'ElseIf MapType = PlanetType.eTerran Then
        MaximizeTerrain()
        'End If

        'If MapType = PlanetType.eTundra Then
        '    If moPFX Is Nothing Then moPFX = New PlanetFX.ParticleEngine(moDevice, Me)

        '    'For X = 0 To Width - 1
        '    '    For Y = 0 To Height - 1
        '    '        lIdx = Y * Width + X
        '    '        If HeightMap(lIdx) < WaterHeight Then
        '    '            Dim bGood As Boolean = False
        '    '            For lTmpX As Int32 = X - 1 To X + 1
        '    '                For lTmpY As Int32 = Y - 1 To Y + 1
        '    '                    If lTmpX > -1 AndAlso lTmpX < Width AndAlso lTmpY > -1 AndAlso lTmpY < Height Then
        '    '                        Dim lTmpIdx As Int32 = lTmpY * Width + lTmpX
        '    '                        If HeightMap(lTmpIdx) > WaterHeight Then
        '    '                            bGood = True
        '    '                            Exit For
        '    '                        End If
        '    '                    End If
        '    '                Next lTmpY
        '    '                If bGood = True Then Exit For
        '    '            Next lTmpX

        '    '            If bGood = True Then
        '    '                Dim vecVent As Vector3

        '    '                vecVent.X = (X - (Width \ 2)) * CellSpacing
        '    '                vecVent.Z = (Y - (Height \ 2)) * CellSpacing
        '    '                vecVent.Y = Me.GetHeightAtLocation(vecVent.X, vecVent.Z) + (3 * ml_Y_Mult)
        '    '                If vecVent.Y < WaterHeight * ml_Y_Mult Then vecVent.Y = WaterHeight * ml_Y_Mult 

        '    '                moPFX.AddEmitter(PlanetFX.ParticleEngine.EmitterType.eSnowDrift, vecVent, 40)
        '    '            End If
        '    '        End If
        '    '    Next Y
        '    'Next X

        '    Dim XMax As Int32 = CInt(Rnd() * 50 + 20)
        '    For lIdx = 0 To XMax
        '        Y = CInt(Rnd() * (Height - 4)) + 2
        '        X = CInt(Rnd() * (Width - 4)) + 2
        '        Dim vecVent As Vector3

        '        'For lTmpX As Int32 = X - 1 To X + 1
        '        '    For lTmpY As Int32 = Y - 1 To Y + 1
        '        '        Dim lTmpIdx As Int32 = (((lTmpY - Y) + 1) * 3) + (lTmpX - X) + 1
        '        vecVent.X = (X - (Width \ 2)) * CellSpacing
        '        vecVent.Z = (Y - (Height \ 2)) * CellSpacing
        '        vecVent.Y = Me.GetHeightAtLocation(vecVent.X, vecVent.Z)
        '        If vecVent.Y < WaterHeight * ml_Y_Mult Then vecVent.Y = WaterHeight * ml_Y_Mult
        '        'Next lTmpY
        '        '    Next lTmpX

        '        moPFX.AddEmitter(PlanetFX.ParticleEngine.EmitterType.eSnowDrift, vecVent, 400)
        '    Next lIdx

        'End If


        For Y = 0 To TerrainClass.Height - 1
            lIdx = (Y * TerrainClass.Width) '+ (X * (TerrainClass.Width - 1))
            Dim lOtherIdx As Int32 = (Y * TerrainClass.Width) + (TerrainClass.Width - 1)

            HeightMap(lOtherIdx) = HeightMap(lIdx)
        Next Y

        mbHMReady = True

    End Sub

    Private mptVolcanoes() As Point
    Private Sub CreateVolcanoes()
        Dim lCnt As Int32 '= CInt(Rnd() * 5) + 1

        'If MapType = PlanetType.eWaterWorld Then
        '    If Rnd() * 100 > 60 Then Return

        '    ReDim mptVolcanoes(0)

        '    Dim lYRng As Int32 = Height \ 2
        '    Dim lYOffset As Int32 = Height \ 4
        '    Dim lIdx As Int32

        '    Dim lVolX As Int32
        '    Dim lVolY As Int32

        '    'Get our max height
        '    Dim lVolcanoVal As Int32
        '    For Y As Int32 = 0 To Height - 1
        '        For X As Int32 = 0 To Width - 1
        '            lIdx = Y * Width + X
        '            If HeightMap(lIdx) > lVolcanoVal Then
        '                lVolcanoVal = HeightMap(lIdx)
        '                lVolX = X
        '                lVolY = Y
        '            End If
        '        Next X
        '    Next Y
        '    lVolcanoVal += 5
        '    If lVolcanoVal > 255 Then lVolcanoVal = 255

        '    mptVolcanoes(0).X = lVolX
        '    mptVolcanoes(0).Y = lVolY

        '    Dim lSize As Int32 = CInt(Rnd() * 4) + 6
        '    For lLocY As Int32 = lVolY - lSize To lVolY + lSize
        '        For lLocX As Int32 = lVolX - lSize To lVolX + lSize

        '            Dim lTempX As Int32 = lLocX
        '            Dim lTempY As Int32 = lLocY
        '            If lTempX > Width - 1 Then lTempX -= Width
        '            If lTempX < 0 Then lTempX += Width
        '            If lTempY > Height - 1 Then lTempY -= Height
        '            If lTempY < 0 Then lTempY += Height

        '            lIdx = lTempY * Width + lTempX

        '            If lLocX <> lVolX OrElse lLocY <> lVolY Then
        '                If HeightMap(lIdx) > WaterHeight Then
        '                    If (Math.Abs(lLocX - lVolX) = 1 AndAlso Math.Abs(lLocY - lVolY) < 2) OrElse (Math.Abs(lLocY - lVolY) = 1 AndAlso Math.Abs(lLocX - lVolX) < 2) Then
        '                        HeightMap(lIdx) = CByte(lVolcanoVal)
        '                    Else
        '                        Dim lXVal As Int32 = lLocX - lVolX
        '                        lXVal *= lXVal
        '                        Dim lYVal As Int32 = lLocY - lVolY
        '                        lYVal *= lYVal

        '                        Dim fTotalVal As Single = CSng(Math.Sqrt(lXVal + lYVal))
        '                        Dim lOriginal As Int32 = HeightMap(lIdx)
        '                        Dim lVal As Int32 = WaterHeight + CInt((lVolcanoVal - WaterHeight) * (1.0F - (fTotalVal / (lSize * 3))))

        '                        If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
        '                        If lVal > lVolcanoVal Then lVal = lVolcanoVal
        '                        If lVal < lOriginal Then lVal = lOriginal

        '                        HeightMap(lIdx) = CByte(lVal)
        '                    End If
        '                End If
        '            Else : HeightMap(lIdx) = CByte(lVolcanoVal)
        '            End If
        '        Next lLocX
        '    Next lLocY

        'Else
        lCnt = CInt(Rnd() * 5) + 1
        ReDim mptVolcanoes(lCnt - 1)

        Dim lYRng As Int32 = Height \ 2
        Dim lYOffset As Int32 = Height \ 4
        Dim lIdx As Int32

        'Get our max height
        Dim lVolcanoVal As Int32
        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lVolcanoVal Then lVolcanoVal = HeightMap(lIdx)
            Next X
        Next Y
        lVolcanoVal += 5
        If lVolcanoVal > 255 Then lVolcanoVal = 255

        While lCnt > 0
            Dim Y As Int32 = CInt(Rnd() * lYRng) + lYOffset
            Dim X As Int32 = CInt(Rnd() * (Width - 1))

            If HeightMap(Y * Width + X) > WaterHeight Then
                mptVolcanoes(lCnt - 1).X = X
                mptVolcanoes(lCnt - 1).Y = Y
                lCnt -= 1

                Dim lSize As Int32 = CInt(Rnd() * 4) + 6
                For lLocY As Int32 = Y - lSize To Y + lSize
                    For lLocX As Int32 = X - lSize To X + lSize

                        Dim lTempX As Int32 = lLocX
                        Dim lTempY As Int32 = lLocY
                        If lTempX > Width - 1 Then lTempX -= Width
                        If lTempX < 0 Then lTempX += Width
                        If lTempY > Height - 1 Then lTempY -= Height
                        If lTempY < 0 Then lTempY += Height

                        lIdx = lTempY * Width + lTempX

                        If lLocX <> X OrElse lLocY <> Y Then
                            If HeightMap(lIdx) > WaterHeight Then
                                If (Math.Abs(lLocX - X) = 1 AndAlso Math.Abs(lLocY - Y) < 2) OrElse (Math.Abs(lLocY - Y) = 1 AndAlso Math.Abs(lLocX - X) < 2) Then
                                    HeightMap(lIdx) = CByte(lVolcanoVal)
                                Else
                                    Dim lXVal As Int32 = lLocX - X
                                    lXVal *= lXVal
                                    Dim lYVal As Int32 = lLocY - Y
                                    lYVal *= lYVal

                                    Dim fTotalVal As Single = CSng(Math.Sqrt(lXVal + lYVal))
                                    Dim lOriginal As Int32 = HeightMap(lIdx)
                                    Dim lVal As Int32 = WaterHeight + CInt((lVolcanoVal - WaterHeight) * (1.0F - (fTotalVal / (lSize * 3))))

                                    If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                    If lVal > lVolcanoVal Then lVal = lVolcanoVal
                                    If lVal < lOriginal Then lVal = lOriginal

                                    HeightMap(lIdx) = CByte(lVal)
                                End If
                            End If
                        Else : HeightMap(lIdx) = CByte(lVolcanoVal)
                        End If
                    Next lLocX
                Next lLocY
            End If

        End While
        'End If

    End Sub

    Private Sub DoPeakAccents()
        Dim X As Int32
        Dim Y As Int32
        Dim lIdx As Int32
        Dim ptPeaks() As Point = Nothing
        Dim lPeakHt() As Int32 = Nothing
        Dim lPeakUB As Int32 = -1
        Dim lMinPeakHt As Int32 = 0

        For Y = 0 To Height - 1
            For X = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lMinPeakHt Then lMinPeakHt = HeightMap(lIdx)
            Next X
        Next Y
        lMinPeakHt -= CInt((Rnd() * 15) + 5)

        For Y = 0 To Height - 1
            For X = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > lMinPeakHt Then
                    Dim bSkip As Boolean = False
                    For lTmpIdx As Int32 = 0 To lPeakUB
                        If Math.Abs(ptPeaks(lTmpIdx).X - X) < 3 AndAlso Math.Abs(ptPeaks(lTmpIdx).Y - Y) < 3 Then
                            If lPeakHt(lTmpIdx) < HeightMap(lIdx) Then
                                ptPeaks(lTmpIdx).X = X
                                ptPeaks(lTmpIdx).Y = Y
                                lPeakHt(lTmpIdx) = HeightMap(lIdx)
                            End If
                            bSkip = True
                        End If
                    Next lTmpIdx

                    If bSkip = False Then
                        lPeakUB += 1
                        ReDim Preserve ptPeaks(lPeakUB)
                        ReDim Preserve lPeakHt(lPeakUB)
                        ptPeaks(lPeakUB).X = X
                        ptPeaks(lPeakUB).Y = Y
                        lPeakHt(lPeakUB) = HeightMap(lIdx)
                    End If
                End If
            Next X
        Next Y

        Dim mfMults(5) As Single
        mfMults(0) = 1.0F
        mfMults(1) = 0.667F
        mfMults(2) = 0.475F
        mfMults(3) = 0.365F
        mfMults(4) = 0.304F
        mfMults(5) = 0.276F

        'Now, check our UB
        If lPeakUB <> -1 Then
            For lPkIdx As Int32 = 0 To lPeakUB

                Dim lVolX As Int32 = ptPeaks(lPkIdx).X
                Dim lVolY As Int32 = ptPeaks(lPkIdx).Y

                For lLocY As Int32 = lVolY - 6 To lVolY + 6
                    For lLocX As Int32 = lVolX - 6 To lVolX + 6

                        Dim lTempX As Int32 = lLocX
                        Dim lTempY As Int32 = lLocY
                        If lTempX > Width - 1 Then lTempX -= Width
                        If lTempX < 0 Then lTempX += Width
                        If lTempY > Height - 1 Then lTempY -= Height
                        If lTempY < 0 Then lTempY += Height

                        lIdx = lTempY * Width + lTempX

                        If lLocX <> lVolX OrElse lLocY <> lVolY Then
                            If HeightMap(lIdx) > WaterHeight Then
                                If (Math.Abs(lLocX - lVolX) = 1 AndAlso Math.Abs(lLocY - lVolY) < 2) OrElse (Math.Abs(lLocY - lVolY) = 1 AndAlso Math.Abs(lLocX - lVolX) < 2) Then
                                    HeightMap(lIdx) = 255
                                Else
                                    Dim lXVal As Int32 = lLocX - lVolX
                                    lXVal *= lXVal
                                    Dim lYVal As Int32 = lLocY - lVolY
                                    lYVal *= lYVal

                                    Dim fTotalVal As Single = lXVal + lYVal 'CSng(Math.Sqrt(lXVal + lYVal))
                                    Dim lOriginal As Int32 = HeightMap(lIdx)
                                    Dim lVal As Int32 = WaterHeight + CInt((255 - WaterHeight) * (1.0F - (fTotalVal / 18.0F)))
                                    If lVal > lOriginal Then lVal = (lVal + lOriginal + lVal) \ 3 Else lVal = lOriginal
                                    If lVal > 255 Then lVal = 255
                                    If lVal < lOriginal \ 2 Then lVal = lOriginal \ 2

                                    HeightMap(lIdx) = CByte(lVal)
                                End If
                            End If
                        Else : HeightMap(lIdx) = CByte(255)
                        End If
                    Next lLocX
                Next lLocY
            Next lPkIdx

        End If
    End Sub

    Private Sub DoGeoPlastic()
        'GeoPlastic - 
        Dim lIdx As Int32
        Dim lTemp As Int32
        Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

        'short dropoff cliffs where rivers of lava flow. 
        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight + 10 Then
                    lTemp = CInt(HeightMap(lIdx)) + 10
                Else : lTemp = 0
                End If
                If lTemp > 255 Then lTemp = 255
                HeightMap(lIdx) = CByte(lTemp)
            Next X
        Next Y

        'Magma regions (water's beach area). 
        'Mountainous regions with no sharp inclines (rolling but with peaks). 

        '"Rips" in the ground for steam vents.
        Dim lCnt As Int32 = CInt(Rnd() * 10) + 15
        While lCnt <> 0
            Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
            Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight AndAlso HeightMap(lIdx) < CInt(lAboveWaterHeight * 0.6F) Then
                Dim lSize As Int32 = CInt(Rnd() * 10) + 10
                Dim lDirX As Int32
                Dim lDirZ As Int32

                If Rnd() * 100 < 50 Then lDirX = -1 Else lDirX = 1
                If Rnd() * 100 < 50 Then lDirZ = -1 Else lDirZ = 1

                lCnt -= 1

                Dim vecSteam(lSize - 1) As Vector3

                While lSize > 0
                    If Rnd() * 100 < 50 Then
                        lSX += lDirX
                    Else : lSZ += lDirZ
                    End If

                    If lSX < 0 Then
                        lSX += Width
                    ElseIf lSX > Width - 1 Then
                        lSX -= Width
                    End If
                    If lSZ < 0 Then
                        lSZ += Height
                    ElseIf lSZ > Height - 1 Then
                        lSZ -= Height
                    End If



                    'If Int(Rnd() * 100) < 7 Then
                    '    mlVentUB += 1
                    '    ReDim Preserve moVents(mlVentUB)
                    '    Dim vecVent As Vector3
                    '    vecVent.X = (lSX - (Width \ 2)) * CellSpacing
                    '    vecVent.Y = WaterHeight * ml_Y_Mult
                    '    vecVent.Z = (lSZ - (Height \ 2)) * CellSpacing
                    '    moVents(mlVentUB) = New LavaColumn(moDevice, vecVent, 150, CellSpacing, -10.0F)
                    'End If
                    lIdx = lSZ * Width + lSX

                    vecSteam(lSize - 1).X = (lSX - (Width \ 2)) * CellSpacing
                    vecSteam(lSize - 1).Y = (HeightMap(lIdx) - 10) * ml_Y_Mult
                    vecSteam(lSize - 1).Z = (lSZ - (Height \ 2)) * CellSpacing

                    HeightMap(lIdx) = 0

                    lSize -= 1
                End While

                'If moPFX Is Nothing Then moPFX = New PlanetFX.ParticleEngine(moDevice, Me)
                'moPFX.AddEmitter(PlanetFX.ParticleEngine.EmitterType.eSlowSteam, vecSteam, 300)

            End If
        End While


        'Reset the volcano center's to 0
        If mptVolcanoes Is Nothing = False Then
            For X As Int32 = 0 To mptVolcanoes.GetUpperBound(0)
                lIdx = mptVolcanoes(X).Y * Width + mptVolcanoes(X).X
                HeightMap(lIdx) = 0

                Dim vecVent As Vector3
                vecVent.X = (mptVolcanoes(X).X - (Width \ 2)) * CellSpacing
                vecVent.Y = WaterHeight * ml_Y_Mult
                vecVent.Z = (mptVolcanoes(X).Y - (Height \ 2)) * CellSpacing

                'If moPFX Is Nothing Then moPFX = New PlanetFX.ParticleEngine(moDevice, Me)
                'moPFX.AddEmitter(PlanetFX.ParticleEngine.EmitterType.eVolcano, vecVent, 400)
            Next
        End If

        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) < WaterHeight Then
                    If Rnd() * 100 < 1 Then
                        'If moPFX Is Nothing Then moPFX = New PlanetFX.ParticleEngine(moDevice, Me)

                        Dim vecVent As Vector3
                        vecVent.X = (X - (Width \ 2)) * CellSpacing
                        vecVent.Y = WaterHeight * ml_Y_Mult
                        vecVent.Z = (Y - (Height \ 2)) * CellSpacing
                        'moPFX.AddEmitter(PlanetFX.ParticleEngine.EmitterType.eLavaColumn, vecVent, 30)

                    End If
                End If
            Next X
        Next Y

    End Sub

    Private Sub DoAcidPlateaus()
        'Ok, at this point, everything is accentuated... we are almost done
        Dim fPlateauStrength() As Single
        ReDim fPlateauStrength((Width * Height) - 1)

        'Ok, now... let's create our plateaus
        Dim lCnt As Int32 = CInt(Rnd() * 3) + 3
        Dim lIdx As Int32 ''

        While lCnt > 0
            Dim lSX As Int32 = CInt(Rnd() * Width - 1)
            Dim lSZ As Int32 = CInt(Rnd() * Height - 1)

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight Then
                lCnt -= 1

                Select Case CInt(Rnd() * 100)
                    Case Is < 50
                        'EAST / WEST
                        'determine which side is the high side
                        If CInt(Rnd() * 100) < 50 Then
                            'East
                            Dim lEndVal As Int32 = Math.Max(lSX - (CInt(Rnd() * 30) + 30), 0)
                            For X As Int32 = lSX To lEndVal Step -1    '0 step -1

                                Dim fHtMult As Single = CSng((X - lEndVal) / (lSX - lEndVal))

                                'For Y As Int32 = 0 To Height - 1
                                For Y As Int32 = Math.Max(lSZ - CInt(Rnd() * 30 + 15), 0) To Math.Min(Height - 1, lSZ + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next Y
                            Next X
                        Else
                            'West
                            Dim lEndVal As Int32 = Math.Min(Width - 1, lSX + (CInt(Rnd() * 30) + 30))
                            For X As Int32 = lSX To lEndVal 'Width - 1

                                Dim fHtMult As Single = CSng((X - lSX) / (lEndVal - lSX))
                                'For Y As Int32 = 0 To Height - 1
                                For Y As Int32 = Math.Max(lSZ - CInt(Rnd() * 30 + 15), 0) To Math.Min(Height - 1, lSZ + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next Y
                            Next X
                        End If
                    Case Else ' Is < 40
                        'NORTH / SOUTH
                        'determine which side is the high side
                        If CInt(Rnd() * 100) < 50 Then
                            'NORTH
                            Dim lEndVal As Int32 = Math.Max(lSZ - (CInt(Rnd() * 30) + 30), 0)
                            For Y As Int32 = lSZ To lEndVal Step -1 '0 Step -1
                                Dim fHtMult As Single = CSng((Y - lEndVal) / (lSZ - lEndVal))
                                'For X As Int32 = 0 To Width - 1
                                For X As Int32 = Math.Max(lSX - CInt(Rnd() * 30 + 15), 0) To Math.Min(Width - 1, lSX + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next X
                            Next Y
                        Else
                            'SOUTH
                            Dim lEndVal As Int32 = Math.Min(Height - 1, lSZ + (CInt(Rnd() * 30) + 30))
                            For Y As Int32 = lSZ To lEndVal 'Height - 1
                                Dim fHtMult As Single = CSng((Y - lSZ) / (lEndVal - lSZ))
                                'For X As Int32 = 0 To Width - 1
                                For X As Int32 = Math.Max(lSX - CInt(Rnd() * 30 + 15), 0) To Math.Min(Width - 1, lSX + CInt(Rnd() * 30 + 15))
                                    lIdx = Y * Width + X
                                    Dim lTemp As Int32 = HeightMap(lIdx)
                                    If lTemp > WaterHeight Then
                                        Dim fTemp As Single = fPlateauStrength(lIdx)
                                        If fTemp = 0 Then
                                            fTemp = fHtMult
                                        Else : fTemp = (fTemp + fHtMult) / 2.0F
                                        End If
                                        fPlateauStrength(lIdx) = fTemp
                                    End If
                                Next X
                            Next Y
                        End If
                End Select

            End If
        End While

        'Now, go through all of our values
        Dim lTmpValWH As Int32 = 255 - WaterHeight
        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                Dim lVal As Int32 = HeightMap(lIdx)

                If lVal > WaterHeight Then
                    Dim fVal As Single = fPlateauStrength(lIdx) * lTmpValWH

                    'Now... I want to affect this by fVal
                    fVal = ((lVal + fVal) / 2.0F) + WaterHeight
                    If fVal > 255 Then
                        HeightMap(lIdx) = 255
                    ElseIf fVal < WaterHeight + 1 Then
                        HeightMap(lIdx) = CByte(WaterHeight + 1)
                    Else
                        HeightMap(lIdx) = CByte(fVal)
                    End If

                End If
            Next X
        Next Y

        MaximizeTerrain()

    End Sub

    Private Sub MaximizeTerrain()
        Dim lIdx As Int32

        Dim lMinVal As Int32 = 255
        Dim lMaxVal As Int32 = 0

        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight Then
                    If HeightMap(lIdx) > lMaxVal Then lMaxVal = HeightMap(lIdx)
                    If HeightMap(lIdx) < lMinVal Then lMinVal = HeightMap(lIdx)
                End If
            Next X
        Next Y

        Dim lDiff As Int32 = lMaxVal - lMinVal
        Dim lDesiredDiff As Int32 = 255 - (WaterHeight + 1)

        For Y As Int32 = 0 To Height - 1
            For X As Int32 = 0 To Width - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight Then
                    Dim lVal As Int32 = HeightMap(lIdx)
                    lVal = CInt(((lVal - lMinVal) / lDiff) * lDesiredDiff) + WaterHeight + 1
                    If lVal < WaterHeight + 1 Then lVal = WaterHeight + 1
                    If lVal > 255 Then lVal = 255
                    HeightMap(lIdx) = CByte(lVal)
                End If
            Next X
        Next Y
    End Sub

    Private Sub DoAcidic()
        Dim lIdx As Int32
        Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

        'Deep chasms with acidic rivers flowing through them. 
        Dim lCnt As Int32 = CInt(Rnd() * 10) + 15
        While lCnt <> 0
            Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
            Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight AndAlso HeightMap(lIdx) < CInt(lAboveWaterHeight * 0.6F) Then
                Dim lSize As Int32 = CInt(Rnd() * 10) + 10
                Dim lDirX As Int32
                Dim lDirZ As Int32

                If Rnd() * 100 < 50 Then lDirX = -1 Else lDirX = 1
                If Rnd() * 100 < 50 Then lDirZ = -1 Else lDirZ = 1

                lCnt -= 1

                Dim vecSteam() As Vector3 = Nothing
                Dim lVecUB As Int32 = -1

                While lSize > 0

                    Dim bXChg As Boolean = False
                    If Rnd() * 100 < 50 Then
                        lSX += lDirX
                        bXChg = True
                    Else : lSZ += lDirZ
                    End If


                    If lSX < 0 Then
                        lSX += Width
                    ElseIf lSX > Width - 1 Then
                        lSX -= Width
                    End If
                    If lSZ < 0 Then
                        lSZ += Height
                    ElseIf lSZ > Height - 1 Then
                        lSZ -= Height
                    End If

                    lIdx = lSZ * Width + lSX
                    HeightMap(lIdx) = 0

                    lVecUB += 1
                    ReDim Preserve vecSteam(lVecUB)
                    With vecSteam(lVecUB)
                        .X = (lSX - (Width \ 2)) * CellSpacing
                        .Y = WaterHeight * ml_Y_Mult
                        .Z = (lSZ - (Height \ 2)) * CellSpacing
                    End With

                    If Rnd() * 100 < 35 Then
                        Dim lTempX As Int32 = lSX
                        Dim lTempZ As Int32 = lSZ

                        If bXChg = True Then
                            lTempZ += lDirZ
                            If lTempZ < 0 Then
                                lTempZ += Height
                            ElseIf lTempZ > Height - 1 Then
                                lTempZ -= Height
                            End If
                        Else
                            lTempX += lDirX
                            If lTempX < 0 Then
                                lTempX += Width
                            ElseIf lTempX > Width - 1 Then
                                lTempX -= Width
                            End If
                        End If
                        lIdx = lTempZ * Width + lTempX
                        HeightMap(lIdx) = 0

                        lVecUB += 1
                        ReDim Preserve vecSteam(lVecUB)
                        With vecSteam(lVecUB)
                            .X = (lTempX - (Width \ 2)) * CellSpacing
                            .Y = WaterHeight * ml_Y_Mult
                            .Z = (lTempZ - (Height \ 2)) * CellSpacing
                        End With
                    End If

                    lSize -= 1
                End While

                'If moPFX Is Nothing Then moPFX = New PlanetFX.ParticleEngine(moDevice, Me)
                'moPFX.AddEmitter(PlanetFX.ParticleEngine.EmitterType.eAcidMist, vecSteam, 300)

            End If
        End While

        'One last soften...
        For X As Int32 = 0 To Width - 1
            For Y As Int32 = 0 To Height - 1
                lIdx = Y * Width + X
                If HeightMap(lIdx) > WaterHeight Then
                    HeightMap(lIdx) = SmoothTerrainVal(X, Y, HeightMap(lIdx))
                End If
            Next Y
        Next X
    End Sub

    Private Sub DoDesert()
        Dim lIdx As Int32

        'smooth mountain regions
        Dim lCnt As Int32 = CInt(Rnd() * 5) + 3
        While lCnt > 0
            Dim lSX As Int32 = CInt(Rnd() * (Width - 1))
            Dim lSZ As Int32 = CInt(Rnd() * (Height - 1))

            lIdx = lSZ * Width + lSX
            If HeightMap(lIdx) > WaterHeight Then
                lCnt -= 1

                Dim lSizeW As Int32 = CInt(Rnd() * 40) + 15
                Dim lHalfSizeW As Int32 = lSizeW \ 2
                Dim lSizeH As Int32 = CInt(Rnd() * 40) + 15
                Dim lHalfSizeH As Int32 = lSizeH \ 2

                For Y As Int32 = -(CInt(lHalfSizeH * Math.Max(Rnd, 0.3))) To (CInt(lHalfSizeH * Math.Max(Rnd, 0.3)))
                    For X As Int32 = -(CInt(lHalfSizeW * Math.Max(Rnd, 0.3))) To (CInt(lHalfSizeW * Math.Max(Rnd, 0.3)))

                        Dim lMapX As Int32 = lSX + X
                        Dim lMapY As Int32 = lSZ + Y

                        If lMapX < 0 Then lMapX += Width
                        If lMapX > Width - 1 Then lMapX -= Width
                        If lMapY < 0 Then lMapY += Height
                        If lMapY > Height - 1 Then lMapY -= Height

                        lIdx = lMapY * Width + lMapX

                        If HeightMap(lIdx) > WaterHeight Then
                            Dim lTemp As Int32 = HeightMap(lIdx)
                            lTemp += CInt(Rnd() * 20) + 100
                            If lTemp > 255 Then lTemp = 255
                            If lTemp < 0 Then lTemp = 0

                            HeightMap(lIdx) = CByte(lTemp)
                        End If
                    Next X
                Next Y

                'Now, add nobs to the top and bottoms
                For X As Int32 = -lHalfSizeW To lHalfSizeW
                    For Y As Int32 = CInt(Rnd() * 10) To 0 Step -1
                        Dim lMapX As Int32 = lSX + X
                        Dim lMapY As Int32 = lSZ - Y - lHalfSizeH

                        If lMapX < 0 Then lMapX += Width
                        If lMapX > Width - 1 Then lMapX -= Width
                        If lMapY < 0 Then lMapY += Height
                        If lMapY > Height - 1 Then lMapY -= Height

                        lIdx = lMapY * Width + lMapX

                        If HeightMap(lIdx) > WaterHeight Then
                            Dim lTemp As Int32 = HeightMap(lIdx)
                            lTemp += CInt(Rnd() * 20) + 100
                            If lTemp > 255 Then lTemp = 255
                            If lTemp < 0 Then lTemp = 0

                            HeightMap(lIdx) = CByte(lTemp)
                        End If
                    Next Y

                    For Y As Int32 = CInt(Rnd() * 10) To 0 Step -1
                        Dim lMapX As Int32 = lSX + X
                        Dim lMapY As Int32 = lSZ + Y + lHalfSizeH

                        If lMapX < 0 Then lMapX += Width
                        If lMapX > Width - 1 Then lMapX -= Width
                        If lMapY < 0 Then lMapY += Height
                        If lMapY > Height - 1 Then lMapY -= Height

                        lIdx = lMapY * Width + lMapX

                        If HeightMap(lIdx) > WaterHeight Then
                            Dim lTemp As Int32 = HeightMap(lIdx)
                            lTemp += CInt(Rnd() * 20) + 100
                            If lTemp > 255 Then lTemp = 255
                            If lTemp < 0 Then lTemp = 0

                            HeightMap(lIdx) = CByte(lTemp)
                        End If
                    Next Y
                Next X
            End If
        End While

    End Sub

    Private Sub DoLunar()
        Dim lPass As Int32
        Dim lPassMax As Int32
        Dim lStartX As Int32
        Dim lStartY As Int32
        Dim lSubX As Int32
        Dim lSubY As Int32
        Dim lSubXTmp As Int32
        Dim lSubYTmp As Int32
        Dim lSubYMax As Int32

        Dim lRadius As Int32
        Dim lImpactPower As Int32
        Dim lCraterBasinHeight As Int32
        Dim lDist As Int32

        'alrighty, let's do some crazy stuff...
        lPassMax = CInt(Int(Rnd() * 10) + 20)     'lpass is the number
        For lPass = 0 To lPassMax
            'ok, get ground-zero
            lStartX = CInt(Int(Rnd() * (Width - 1)))
            lStartY = CInt(Int(Rnd() * (Height - 1)))
            'and our radius
            lRadius = CInt(Int(Rnd() * 32) + 1)
            lImpactPower = CInt(Int(Rnd() * 10) + 1) 'impactpower
            lCraterBasinHeight = HeightMap(lStartY * Width + lStartX) - 60       'crater basin height
            If lCraterBasinHeight < 0 Then lCraterBasinHeight = 0

            lSubYMax = (lStartY + (lRadius + lImpactPower))
            If lSubYMax > Height - 1 Then lSubYMax = Height - 1

            For lSubY = (lStartY - (lRadius + lImpactPower)) To lSubYMax
                For lSubX = (lStartX - (lRadius + lImpactPower)) To (lStartX + (lRadius + lImpactPower))
                    lSubXTmp = lSubX
                    lSubYTmp = lSubY

                    If lSubYTmp < 0 Then lSubYTmp = 0
                    If lSubXTmp > Width - 1 Then lSubXTmp -= Width
                    If lSubXTmp < 0 Then lSubXTmp += Width

                    lDist = CInt(Math.Floor(Distance(lStartX, lStartY, lSubX, lSubY)))
                    Dim lTemp As Int32 = HeightMap(lSubYTmp * Width + lSubXTmp)

                    If lDist < lRadius Then
                        lTemp = CInt(lCraterBasinHeight + Int(Rnd() * 5) + 1)
                    ElseIf lDist = lRadius Then
                        lTemp = CInt(lCraterBasinHeight + 128)
                    ElseIf lDist - lRadius < lImpactPower Then
                        lTemp = CInt((lCraterBasinHeight + (12 * lImpactPower)) - (((lDist - lRadius) / lImpactPower) * lCraterBasinHeight))
                    End If
                    If lTemp < 0 Then lTemp = 0
                    If lTemp > 255 Then lTemp = 255
                    HeightMap(lSubYTmp * Width + lSubXTmp) = CByte(lTemp)
                Next lSubX
            Next lSubY
        Next lPass
    End Sub

    Private Function GetWaterPerc(ByVal lType As Int32) As Int32
        'Select Case lType
        '    Case PlanetType.eBarren
        '        Return 0
        '    Case PlanetType.eDesert
        '        Return CInt(Int(Rnd() * 15) + 1)
        '    Case PlanetType.eGeoPlastic
        '        Return CInt(Int(Rnd() * 40) + 15)
        '    Case PlanetType.eWaterWorld
        '        Return CInt(80 + (Int(Rnd() * 20) + 1))
        '    Case PlanetType.eAdaptable
        '        Return CInt(5 + CInt(Rnd() * 40))
        '    Case PlanetType.eTundra
        '        Return 20 + CInt(Rnd() * 40)
        '    Case PlanetType.eTerran
        Return 30 + CInt(Rnd() * 20)
        '    Case Else 'PlanetType.eTerran, PlanetType.eAcidic
        '        Return CInt(30 + (Int(Rnd() * 50) + 1))

        'End Select
    End Function

    Private Function SmoothTerrainVal(ByVal X As Int32, ByVal Y As Int32, ByVal lVal As Int32) As Byte
        Dim fCorners As Single
        Dim fSides As Single
        Dim fCenter As Single
        Dim fTotal As Single

        Dim lBackX As Int32
        Dim lBackY As Int32
        Dim lForeX As Int32
        Dim lForeY As Int32

        If X = 0 Then
            lBackX = Width - 1
        Else : lBackX = X - 1
        End If
        If X = Width - 1 Then
            lForeX = 0
        Else : lForeX = X + 1
        End If
        If Y = 0 Then
            lBackY = 0
        Else : lBackY = Y - 1
        End If
        If Y = Height - 1 Then
            lForeY = Height - 2
        Else : lForeY = Y + 1
        End If

        fCorners = 0
        fCorners = fCorners + HeightMap((lBackY * Width) + lBackX) 'muTiles(lBackX, lBackY).lHeight
        fCorners = fCorners + HeightMap((lBackY * Width) + lForeX) 'muTiles(lForeX, lBackY).lHeight
        fCorners = fCorners + HeightMap((lForeY * Width) + lBackX) 'muTiles(lBackX, lForeY).lHeight
        fCorners = fCorners + HeightMap((lForeY * Width) + lForeX) 'muTiles(lForeX, lForeY).lHeight
        fCorners = fCorners / 16

        fSides = 0
        fSides = fSides + HeightMap((Y * Width) + lBackX) 'muTiles(lBackX, Y).lHeight
        fSides = fSides + HeightMap((Y * Width) + lForeX) 'muTiles(lForeX, Y).lHeight
        fSides = fSides + HeightMap((lBackY * Width) + X) 'muTiles(X, lBackY).lHeight
        fSides = fSides + HeightMap((lForeY * Width) + X) 'muTiles(X, lForeY).lHeight
        fSides = fSides / 8

        fCenter = lVal / 4.0F

        fTotal = fCorners + fSides + fCenter
        If fTotal < 0 Then fTotal = 0
        If fTotal > 255 Then fTotal = 255

        Return CByte(fTotal)
    End Function

    Public Shared bReloadVertexBuffer As Boolean = False
    Private Sub LoadVertexBuffer(ByVal bPlanetMap As Boolean)
        Dim Z As Int32
        Dim X As Int32

        Dim lHeight As Int32
        Dim Nx As Single
        Dim Ny As Single
        Dim Nz As Single
        Dim fTu As Single
        Dim fTv As Single
        Dim fTw As Single

        Dim lQX As Int32
        Dim lQZ As Int32

        ReDim moVerts(VertsTotal - 1)

        'MSC 06/05/08 - remarked this
        'Device.IsUsingEventHandlers = False
        'moVB = New VertexBuffer(moVerts(0).GetType, moVerts.Length, moDevice, Usage.WriteOnly, terVertFmt, Pool.Managed)
        'Device.IsUsingEventHandlers = True

        'Ok, 0,0 is the FOW Grid's center (305,305)
        Dim lQuadsPerFOWQuad As Int32 = CInt(Math.Floor(16384 / (mlVertsPerQuad * CellSpacing)))
        'lQuadsPerFOWQuad should be equal to 8, 5, 4, 3 or 2 depending on CellSpacing 
        Dim fFOWTexUsage As Single = (lQuadsPerFOWQuad * mlVertsPerQuad * CellSpacing) / 16384.0F

        For Z = 0 To Height - 1
            For X = 0 To Width - 1
                lHeight = HeightMap((Z * Height) + X)

                lQX = CInt(Math.Floor(X / mlVertsPerQuad))
                lQZ = CInt(Math.Floor(Z / mlVertsPerQuad))
                fTu = CSng(X / mlVertsPerQuad)
                fTv = CSng(Z / mlVertsPerQuad)
                If lQX Mod 2 = 0 Then
                    fTu = fTu - lQX
                Else
                    fTu = 1.0F - (fTu - lQX)
                End If
                If lQZ Mod 2 = 0 Then
                    fTv = fTv - lQZ
                Else
                    fTv = 1.0F - (fTv - lQZ)
                End If
 
                If muSettings.TerrainTextureResolution < 1 Then muSettings.TerrainTextureResolution = 1
                fTu *= muSettings.TerrainTextureResolution
                fTv *= muSettings.TerrainTextureResolution

                If bPlanetMap = True Then
                    fTu = CSng(X / Width)
                    fTv = CSng(Z / Height)

                    fTu *= 3
                    fTv *= 3
                End If

                Nx = 0
                Ny = 1
                Nz = 0

                fTw = GetTWTextureCoord(lHeight)
               
                If fTw < 0.0F Then fTw = 0.0F

                With moVerts((Z * Height) + X)
                    .p = New Vector3((X - (mlHalfWidth)) * CellSpacing, lHeight * ml_Y_Mult, (Z - (mlHalfHeight)) * CellSpacing)
                    .n = New Vector3(Nx, Ny, Nz)
                    .tu = fTu
                    .tv = fTv
                    .tw = fTw

                    .tu2 = CSng(X / Width)
                    .tv2 = 1 - CSng(Z / Height)
                    .tw2 = 0
                End With

            Next X
        Next Z

        ComputeNormals(moVerts)

        SetMatsTexsAndGetWaterColor()

        'MSC 06/05/08 - remarked this
        'moVB.SetData(moVerts, 0, LockFlags.Discard)

        'MSC 06/05/08 - added this code for the terrain bump mapping
        If muSettings.BumpMapTerrain = True Then
            'k, we need to create our new verts...
            moVertsTBN = DoTangentCalcs(moVerts, LoadIndexBuffer())
        Else
            LoadIndexBuffer()
        End If
        'Now, do our loadvertexbuffer2() code here...
        SetVertexBufferData()
    End Sub


    Private Sub ComputeNormals(ByRef oVerts() As terVert)
        Dim Z As Int32
        Dim X As Int32

        Dim vecX As Vector3 = New Vector3()
        Dim vecZ As Vector3 = New Vector3()
        Dim vecN As Vector3

        'Dim fA As Single
        'Dim fB As Single
        'Dim fC As Single

        'Dim lIdx As Int32

        ReDim HMNormals(oVerts.Length - 1)

        For Z = 1 To QuadsZ - 1
            For X = 1 To QuadsX - 1
                vecX = Vector3.Subtract(oVerts(Z * Width + X + 1).p, oVerts(Z * Width + X - 1).p)
                vecZ = Vector3.Subtract(oVerts((Z + 1) * Width + X).p, oVerts((Z - 1) * Width + X).p)
                vecN = Vector3.Cross(vecZ, vecX)
                vecN.Normalize()
                oVerts(Z * Height + X).n = vecN

                HMNormals(Z * Width + X) = vecN
            Next X
        Next Z

        vecX = Nothing
        vecZ = Nothing
        vecN = Nothing

        NormalsReady = True
    End Sub

    Private Sub ComputeNormals2(ByRef oVerts() As terVert, ByVal lOffsetVertX As Int32, ByVal lOffsetVertZ As Int32)
        Dim Z As Int32
        Dim X As Int32

        Dim vecX As Vector3 = New Vector3()
        Dim vecZ As Vector3 = New Vector3()
        Dim vecN As Vector3

        Dim XMax As Int32 = mlVertsPerQuad '- 1
        Dim ZMax As Int32 = mlVertsPerQuad '- 1
        Dim lOffsetVal As Int32 = mlVertsPerQuad + 1

        If lOffsetVertZ = 230 Then ZMax -= 1
        If lOffsetVertX = 230 Then lOffsetVal -= 1

        For Z = 0 To ZMax
            For X = 0 To XMax
                Dim lZVal As Int32 = lOffsetVertZ + Z
                Dim lXVal As Int32 = lOffsetVertX + X

                If lZVal = 240 Then lZVal = 0
                If lXVal = 240 Then lXVal = 0
                If lXVal = 0 OrElse lXVal = 239 Then
                    lXVal = 0
                    Dim vecRight, vecLeft, vecUp, vecDown As Vector3
                    With vecRight : .X = CellSpacing : .Z = 0 : .Y = HeightMap(lZVal * Width + 1) * ml_Y_Mult : End With
                    With vecLeft : .X = -CellSpacing : .Z = 0 : .Y = HeightMap(lZVal * Width + Width - 1) * ml_Y_Mult : End With
                    If lZVal = 0 Then
                        With vecDown : .X = 0 : .Z = CellSpacing : .Y = HeightMap((lZVal + 1) * Width) * ml_Y_Mult : End With
                        With vecUp : .X = 0 : .Z = -CellSpacing : .Y = vecDown.Y : End With
                    ElseIf lZVal = 239 Then
                        With vecUp : .X = 0 : .Z = -CellSpacing : .Y = HeightMap((lZVal - 1) * Width) * ml_Y_Mult : End With
                        With vecDown : .X = 0 : .Z = CellSpacing : .Y = vecUp.Y : End With
                    Else
                        With vecDown : .X = 0 : .Z = CellSpacing : .Y = HeightMap((lZVal + 1) * Width) * ml_Y_Mult : End With
                        With vecUp : .X = 0 : .Z = -CellSpacing : .Y = HeightMap((lZVal - 1) * Width) * ml_Y_Mult : End With
                    End If

                    vecX = Vector3.Subtract(vecRight, vecLeft)
                    vecZ = Vector3.Subtract(vecUp, vecDown)
                    vecN = Vector3.Cross(vecZ, vecX)
                    vecN.Multiply(-1)
                    vecN.Normalize()
                    oVerts(Z * lOffsetVal + X).n = vecN
                Else
                    Dim lHMIdx As Int32 = (lZVal * Width) + lXVal
                    oVerts(Z * lOffsetVal + X).n = HMNormals(lHMIdx)
                End If

            Next X
        Next Z

        vecX = Nothing
        vecZ = Nothing
        vecN = Nothing

        NormalsReady = True
    End Sub


    Private Shared Function GetRangedTexCoord(ByVal lMinVal As Int32, ByVal lMaxVal As Int32, ByVal lVal As Int32) As Single
        Dim lRange As Int32 = lMaxVal - lMinVal
        Dim lValInRange As Int32 = lVal - lMinVal
        Return CSng(lValInRange / lRange)
    End Function
 
    Private Function LoadIndexBuffer() As Int32()
        Dim lCnt As Int32 = 150962      '(for 24x24 quads)
        Dim lIndex As Int32
        Dim Z As Int32
        Dim X As Int32
        Dim mlIndices() As Int32

        ReDim mlIndices(lCnt)
        'Device.IsUsingEventHandlers = False
        'moIB = New IndexBuffer(lCnt.GetType, mlIndices.Length, moDevice, Usage.WriteOnly, Pool.Managed)
        'Device.IsUsingEventHandlers = True

        'Ok, so Z from QuadBaseZ to QuadBaseZ + QuadExtZ (+1 if this is NOT the end)
        '  X From QuadBaseX to QuadBaseX + QuadExtX (+1 if this is not the end)
        Dim lQuadBaseX As Int32
        Dim lQuadBaseZ As Int32
        Dim lSubZ As Int32
        Dim lSubX As Int32

        Dim lTmpX As Int32
        Dim lTmpZ As Int32
        Dim lQuadIdx As Int32

        Dim lSubZMax As Int32
        Dim lSubXMax As Int32

        Dim lQuadsPerFOWQuad As Int32 = CInt(Math.Floor(16384 / (mlVertsPerQuad * CellSpacing)))
        Dim fFOWTexUsage As Single = CSng(lQuadsPerFOWQuad / Math.Ceiling(16384 / (mlVertsPerQuad * CellSpacing)))
        Dim lMidQuad As Int32 = CInt(Math.Floor(mlQuads / 2))

        ReDim mlQuadFOWXIndex((mlQuads * mlQuads) - 1)
        ReDim mlQuadFOWZIndex((mlQuads * mlQuads) - 1)

        ReDim mlQuadStartIndex((mlQuads * mlQuads) - 1)
        ReDim mlQuadX((mlQuads * mlQuads) - 1)
        ReDim mlQuadZ((mlQuads * mlQuads) - 1)

        ReDim mcbQuadBox((mlQuads * mlQuads) - 1)
        ReDim mcbQB_East((mlQuads * mlQuads) - 1)
        ReDim mcbQB_West((mlQuads * mlQuads) - 1)

        Dim lHalfMod As Int32 = mlHalfWidth * CellSpacing

        lIndex = 0

        For Z = 0 To (mlQuads - 1)
            For X = 0 To (mlQuads - 1)
                'ok, in a quad... find our quad bases
                lQuadBaseX = X * mlVertsPerQuad
                lQuadBaseZ = Z * mlVertsPerQuad

                lQuadIdx = (Z * mlQuads) + X
                mlQuadStartIndex(lQuadIdx) = lIndex
                mlQuadX(lQuadIdx) = lQuadBaseX
                mlQuadZ(lQuadIdx) = lQuadBaseZ

                mcbQuadBox(lQuadIdx).vecMin.X = (lQuadBaseX - (Width \ 2)) * CellSpacing
                mcbQuadBox(lQuadIdx).vecMin.Y = 0
                mcbQuadBox(lQuadIdx).vecMin.Z = (lQuadBaseZ - (Height \ 2)) * CellSpacing
                mcbQuadBox(lQuadIdx).vecMax.X = ((lQuadBaseX + mlVertsPerQuad) - (Width \ 2)) * CellSpacing
                mcbQuadBox(lQuadIdx).vecMax.Y = 256I * Me.ml_Y_Mult
                mcbQuadBox(lQuadIdx).vecMax.Z = ((lQuadBaseZ + mlVertsPerQuad) - (Height \ 2)) * CellSpacing

                mlQuadFOWXIndex(lQuadIdx) = CInt((Math.Floor(((lQuadBaseX - mlHalfWidth) / mlVertsPerQuad) / lQuadsPerFOWQuad)))
                mlQuadFOWZIndex(lQuadIdx) = CInt((Math.Floor(((lQuadBaseZ - mlHalfHeight) / mlVertsPerQuad) / lQuadsPerFOWQuad)))

                'Now, go thru normally... the quads inside our quad
                If Z = (mlQuads - 1) Then
                    lSubZMax = mlVertsPerQuad - 2
                Else : lSubZMax = mlVertsPerQuad - 1
                End If
                If X = (mlQuads - 1) Then
                    lSubXMax = mlVertsPerQuad - 1
                Else : lSubXMax = mlVertsPerQuad
                End If

                For lSubZ = 0 To lSubZMax
                    lTmpZ = (lQuadBaseZ + lSubZ) * Width

                    If lSubZ Mod 2 = 0 Then
                        For lSubX = 0 To lSubXMax
                            lTmpX = lQuadBaseX + lSubX
                            mlIndices(lIndex) = lTmpX + lTmpZ : lIndex += 1
                            mlIndices(lIndex) = lTmpX + lTmpZ + Width : lIndex += 1
                        Next lSubX
                        If lSubZ <> Height - 2 Then
                            lSubX -= 1
                            lTmpX = lQuadBaseX + lSubX
                            mlIndices(lIndex) = lTmpX + lTmpZ
                        End If
                    Else
                        For lSubX = lSubXMax To 0 Step -1
                            lTmpX = lQuadBaseX + lSubX
                            mlIndices(lIndex) = lTmpX + lTmpZ : lIndex += 1
                            mlIndices(lIndex) = lTmpX + lTmpZ + Width : lIndex += 1
                        Next lSubX
                        If lSubZ <> Height - 2 Then
                            lSubX += 1
                            lTmpX = lQuadBaseX + lSubX
                            mlIndices(lIndex) = lTmpX + lTmpZ
                        End If
                    End If
                Next lSubZ
            Next X
        Next Z

        'moIB.SetData(mlIndices, 0, LockFlags.None)

        'Now, for the east and west boxes...
        Dim lAdjustX As Int32 = Me.CellSpacing * TerrainClass.Width '(TerrainClass.Width - 1)
        For lTmpIdx As Int32 = 0 To mcbQuadBox.GetUpperBound(0)
            With mcbQB_East(lTmpIdx)
                .vecMax = mcbQuadBox(lTmpIdx).vecMax
                .vecMin = mcbQuadBox(lTmpIdx).vecMin
                .vecMax.X -= lAdjustX
                .vecMin.X -= lAdjustX
            End With
            With mcbQB_West(lTmpIdx)
                .vecMax = mcbQuadBox(lTmpIdx).vecMax
                .vecMin = mcbQuadBox(lTmpIdx).vecMin
                .vecMax.X += lAdjustX
                .vecMin.X += lAdjustX
            End With
        Next lTmpIdx

        Return mlIndices
    End Function

    Private Sub GetHeightMapColor(ByVal lHeight As Int32, ByRef lR As Int32, ByRef lG As Int32, ByRef lB As Int32)
        Dim fTw As Single
        Dim lAboveWaterHeight As Int32 = 255 - WaterHeight

        'Select Case MapType
        '    Case PlanetType.eAcidic
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.01)
        '                lR = 0 : lG = 128 : lB = 0
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.1)
        '                lR = 128 : lG = 164 : lB = 128
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.1), 255, lHeight)
        '                fTw *= 6
        '                fTw += 2
        '                If fTw < 3 Then
        '                    lR = 128 : lG = 128 : lB = 128
        '                ElseIf fTw < 4 Then
        '                    lR = 103 : lG = 164 : lB = 0
        '                ElseIf fTw < 5 Then
        '                    lR = 128 : lG = 150 : lB = 128
        '                ElseIf fTw < 6 Then
        '                    lR = 96 : lG = 148 : lB = 96
        '                ElseIf fTw < 7 Then
        '                    lR = 96 : lG = 164 : lB = 96
        '                Else
        '                    lR = 156 : lG = 192 : lB = 156
        '                End If
        '        End Select

        '    Case PlanetType.eAdaptable
        '        Select Case lHeight
        '            Case Is < CInt(lAboveWaterHeight * 0.15) + WaterHeight
        '                fTw = GetRangedTexCoord(WaterHeight, CInt(lAboveWaterHeight * 0.15) + WaterHeight, lHeight)
        '                lR = 72 + CInt(fTw * 20) + CInt(Rnd() * 10)
        '                lG = 32 + CInt(fTw * 20) + CInt(Rnd() * 10)
        '                lB = 32 + CInt(fTw * 10) + CInt(Rnd() * 10)
        '            Case Is < CInt(lAboveWaterHeight * 0.75) + WaterHeight
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.15), CInt(lAboveWaterHeight * 0.75) + WaterHeight, lHeight)
        '                lR = 128 + CInt(Rnd() * 10) - 5
        '                lG = 45 + CInt(fTw * 45) + CInt(Rnd() * 10) - 5
        '                lB = CInt(fTw * 96) + CInt(Rnd() * 5)
        '            Case Is < CInt(lAboveWaterHeight * 0.8) + WaterHeight
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.75), WaterHeight + CInt(lAboveWaterHeight * 0.8), lHeight)
        '                lR = 115 + CInt(fTw * 40) + CInt(Rnd() * 10) - 5
        '                lG = 80 + CInt(fTw * 70) + CInt(Rnd() * 10)
        '                lB = 80 + CInt(fTw * 70) + CInt(Rnd() * 10)
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.8), 255, lHeight)
        '                lR = 160 + CInt(70 * fTw)
        '                lB = lR + CInt(Rnd() * 10)
        '                lG = lR + CInt(Rnd() * 10)
        '                lR += CInt(Rnd() * 10)
        '        End Select
        '    Case PlanetType.eBarren
        '        fTw = (lHeight / 255.0F)
        '        Dim lVal As Int32 = CInt(128 + (fTw * 128) + ((Rnd() * 10) - 5))
        '        If lVal > 255 Then lVal = 255
        '        If lVal < 0 Then lVal = 255
        '        lR = lVal
        '        lG = lVal
        '        lB = lVal
        '        'lR = 32 + CInt(fTw * 200) + CInt(Rnd() * 10) - 5
        '        'lG = 32 + CInt(fTw * 200) + CInt(Rnd() * 10) - 5
        '        'lB = 32 + CInt(fTw * 200) + CInt(Rnd() * 10) - 5
        '    Case PlanetType.eDesert
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.1)
        '                fTw = GetRangedTexCoord(WaterHeight, WaterHeight + CInt(lAboveWaterHeight * 0.1), lHeight)
        '                lR = CInt(fTw * 128) + CInt(Rnd() * 10)
        '                lG = 64 + CInt((Rnd() * 10) - 5)
        '                lB = 64 - CInt(fTw * 64) + CInt(Rnd() * 10)
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.2)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.1), WaterHeight + CInt(lAboveWaterHeight * 0.2), lHeight)
        '                lR = 112 + CInt(fTw * 80) + CInt(Rnd() * 10)
        '                lG = 64 + CInt(fTw * 100) + CInt(Rnd() * 10)
        '                lB = CInt(fTw * 132) + CInt(Rnd() * 5)
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.75)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.2), CInt(lAboveWaterHeight * 0.75) + WaterHeight, lHeight)
        '                lR = 160 + CInt(fTw * 30) + CInt(Rnd() * 10)
        '                lG = 140 + CInt(fTw * 30) + CInt(Rnd() * 10)
        '                lB = 115 + CInt(fTw * 30) + CInt(Rnd() * 10)
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.8)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.75), CInt(lAboveWaterHeight * 0.8) + WaterHeight, lHeight)
        '                lR = 190 - CInt(fTw * 50) + CInt(Rnd() * 10)
        '                lG = 170 - CInt(fTw * 70) + CInt(Rnd() * 10)
        '                lB = 128 - CInt(fTw * 80) + CInt(Rnd() * 10)
        '            Case Else
        '                lR = 115 + CInt(Rnd() * 25)
        '                lG = 80 + CInt(Rnd() * 30)
        '                lB = 50 + CInt(Rnd() * 25)
        '        End Select
        '    Case PlanetType.eGeoPlastic
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.01)
        '                lR = 255 : lG = 64 : lB = 0
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.15)
        '                lR = 96 : lG = 64 : lB = 64
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.3)
        '                lR = 110 : lG = 96 : lB = 96
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.45)
        '                lR = 128 : lG = 110 : lB = 110
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.6)
        '                lR = 132 : lG = 128 : lB = 128
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.75)
        '                lR = 148 : lG = 132 : lB = 132
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.9)
        '                lR = 164 : lG = 148 : lB = 148
        '            Case Else
        '                lR = 196 : lG = 164 : lB = 164
        '        End Select
        '    Case PlanetType.eTerran
        Select Case lHeight
            Case Is < CInt(lAboveWaterHeight * 0.1) + WaterHeight
                fTw = GetRangedTexCoord(WaterHeight, CInt(lAboveWaterHeight * 0.1) + WaterHeight, lHeight)
                lR = CInt(fTw * 100) + 40 + CInt((Rnd() * 10) - 5)
                lG = CInt(fTw * 80) + 45 + CInt(Rnd() * 5)
                lB = 120 - CInt(fTw * 20)
            Case Is < CInt(lAboveWaterHeight * 0.15) + WaterHeight
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.1) + WaterHeight, CInt(lAboveWaterHeight * 0.15) + WaterHeight, lHeight)
                lR = 170 - CInt(fTw * 40) + CInt(Rnd() * 5)
                lG = 150 - CInt(fTw * 30) + CInt(Rnd() * 5)
                lB = 100 - CInt(fTw * 10)
            Case Is < CInt(lAboveWaterHeight * 0.51) + WaterHeight
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.15) + WaterHeight, CInt(lAboveWaterHeight * 0.51) + WaterHeight, lHeight)
                lR = 32 + CInt(fTw * 16) + CInt((Rnd() * 8) - 4)
                lG = 72 + CInt(fTw * 32) + CInt((Rnd() * 8) - 4)
                lB = 32 + CInt(fTw * 16) + CInt((Rnd() * 8) - 4)
            Case Is < CInt(lAboveWaterHeight * 0.7) + WaterHeight
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.51) + WaterHeight, CInt(lAboveWaterHeight * 0.7) + WaterHeight, lHeight)
                lR = 120 + CInt(40 * fTw) + CInt(Rnd() * 10)
                lG = 80 + CInt(40 * fTw) + CInt(Rnd() * 5)
                lB = 40 + CInt(40 * fTw) + CInt(Rnd() * 5)
            Case Else
                fTw = GetRangedTexCoord(CInt(lAboveWaterHeight * 0.7) + WaterHeight, 255, lHeight)
                lR = 160 + CInt(fTw * 80) + CInt((Rnd() * 10) - 5)
                lG = 160 + CInt(fTw * 80) + CInt((Rnd() * 10) - 5)
                lB = 160 + CInt(fTw * 88) + CInt((Rnd() * 10) - 5)
        End Select
        '    Case PlanetType.eTundra
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.1)
        '                fTw = GetRangedTexCoord(WaterHeight, WaterHeight + CInt(lAboveWaterHeight * 0.1), lHeight)
        '                lR = CInt((fTw * 192) + Rnd() * 10)
        '                lG = 220 - CInt(fTw * 45) + CInt(Rnd() * 5)
        '                lB = 220 - CInt(fTw * 22) + CInt(Rnd() * 5)
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.1), 255, lHeight)
        '                lR = 110 + CInt(fTw * 90) + CInt(Rnd() * 10) - 5
        '                lG = 128 + CInt(fTw * 90) + CInt(Rnd() * 10) - 5
        '                lB = 150 + CInt(fTw * 90) + CInt(Rnd() * 10) - 5

        '                lR = 132 + CInt(fTw * 108) + CInt(Rnd() * 10) - 5
        '                lG = 140 + CInt(fTw * 150)
        '                lB = 148 + CInt(fTw * 200)
        '                If lG > 250 Then lG = 250
        '                lG += CInt(Rnd() * 10) - 5
        '                If lB > 250 Then lB = 250
        '                lB += CInt(Rnd() * 5)
        '        End Select
        '    Case PlanetType.eWaterWorld
        '        Select Case lHeight
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.1)
        '                fTw = GetRangedTexCoord(WaterHeight, WaterHeight + CInt(lAboveWaterHeight * 0.1), lHeight)
        '                lR = CInt(fTw * 128) + CInt(Rnd() * 10)
        '                lG = 64 + CInt((Rnd() * 10) - 5)
        '                lB = 64 - CInt(fTw * 64) + CInt(Rnd() * 10)
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.2)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.1), WaterHeight + CInt(lAboveWaterHeight * 0.2), lHeight)
        '                lR = 112 + CInt(fTw * 80) + CInt(Rnd() * 10)
        '                lG = 64 + CInt(fTw * 100) + CInt(Rnd() * 10)
        '                lB = CInt(fTw * 132) + CInt(Rnd() * 5)
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.4)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.2), CInt(lAboveWaterHeight * 0.5) + WaterHeight, lHeight)
        '                lR = 160 + CInt(fTw * 30) + CInt(Rnd() * 10)
        '                lG = 140 + CInt(fTw * 30) + CInt(Rnd() * 10)
        '                lB = 115 + CInt(fTw * 30) + CInt(Rnd() * 10)
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.66)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.4), WaterHeight + CInt(lAboveWaterHeight * 0.66), lHeight)
        '                lR = 190 - CInt(fTw * 150) + CInt(Rnd() * 10)
        '                lG = 170 - CInt(fTw * 100) + CInt(Rnd() * 5)
        '                lB = 140 - CInt(fTw * 100) + CInt(Rnd() * 5)
        '            Case Is < WaterHeight + CInt(lAboveWaterHeight * 0.8)
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.66), CInt(lAboveWaterHeight * 0.8) + WaterHeight, lHeight)
        '                lR = 32 + CInt(fTw * 16) + CInt((Rnd() * 32) - 16)
        '                lG = 72 + CInt(fTw * 32) + CInt((Rnd() * 32) - 16)
        '                lB = 32 + CInt(fTw * 16) + CInt((Rnd() * 32) - 16)
        '            Case Else
        '                fTw = GetRangedTexCoord(WaterHeight + CInt(lAboveWaterHeight * 0.8), 255, lHeight)
        '                lR = 96 + CInt(fTw * 48) + CInt(Rnd() * 5)
        '                lB = lR
        '                lG = 112 + CInt(fTw * 48) + CInt(Rnd() * 5)
        '        End Select
        'End Select

    End Sub

    Protected Overrides Sub Finalize()
        CleanResources()
        MyBase.Finalize()
    End Sub

    'NOTE: DO NOT CALL THIS NORMALLY, THIS ROUTINE IS FOR SPEED MOD CALCULATIONS!!!
    Public Function GetTerrainNormalEx(ByVal fVertX As Single, ByVal fVertY As Single) As Vector3
        Dim fTX As Single   'translated X 
        Dim fTZ As Single   'translated Z 
        Dim lCol As Int32
        Dim lRow As Int32

        Dim dZ As Single
        Dim dX As Single

        Dim vN1 As Vector3
        Dim vN2 As Vector3
        Dim vN3 As Vector3
        Dim vN4 As Vector3

        Dim lIdx As Int32

        If HMNormals Is Nothing Then Return New Vector3(0.0F, 1.0F, 0.0F)

        fTX = fVertX
        fTZ = fVertY

        lCol = CInt(Math.Floor(fTX))
        lRow = CInt(Math.Floor(fTZ))
        dX = fTX - lCol
        dZ = fTZ - lRow

        lIdx = (lRow * Width) + lCol
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN1 = Vector3.Empty Else vN1 = HMNormals(lIdx)
        lIdx = (lRow * Width) + lCol + 1
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN2 = Vector3.Empty Else vN2 = HMNormals(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN3 = Vector3.Empty Else vN3 = HMNormals(lIdx)
        lIdx = ((lRow + 1) * Width) + lCol + 1
        If lIdx < 0 Or lIdx > HMNormals.Length - 1 Then vN4 = Vector3.Empty Else vN4 = HMNormals(lIdx)

        vN1.Multiply(1 - dX)
        vN2.Multiply(dX)
        Dim vA As Vector3 = Vector3.Add(vN1, vN2)
        vN3.Multiply(1 - dX)
        vN4.Multiply(dX)
        Dim vB As Vector3 = Vector3.Add(vN3, vN4)
        vA.Multiply(1 - dZ)
        vB.Multiply(dZ)
        Dim vF As Vector3 = Vector3.Add(vA, vB)
        vF.Normalize()

        Return vF
    End Function

    Public Function CreateFOWDisk(ByRef moDevice As Device, ByVal oVisibleColor As System.Drawing.Color, ByVal oShroudColor As System.Drawing.Color) As Mesh
        Dim lSlices As Int32 = 16
        Dim fInnerRadius As Single = 0.8F
        Dim fOuterRadius As Single = 1.0F

        Dim fPtOX As Single
        Dim fPtOZ As Single
        Dim fPtIX As Single
        Dim fPtIZ As Single
        Dim fAngle As Single
        Dim X As Int32
        Dim lIdx As Int32

        Dim lVerts As Int32 = (lSlices * 2) + 1
        Dim lRanks(0) As Int32

        Dim lInnerIdx() As Int32
        Dim lCurrentInner As Int32
        Dim lCenterIndex As Int32

        Dim pnt As CustomVertex.PositionColored

        Dim fTX As Single
        Dim fTZ As Single

        Dim iIndices() As Int16

        Device.IsUsingEventHandlers = False

        lRanks(0) = lVerts

        Dim oTmp As Mesh = New Mesh(lVerts * 3, lVerts, MeshFlags.Managed, CustomVertex.PositionColored.Format, moDevice)
        Dim data As System.Array = oTmp.VertexBuffer.Lock(0, (New CustomVertex.PositionColored()).GetType(), LockFlags.None, lRanks)

        ReDim iIndices(((lVerts - 1) * 3) * 2)

        fPtOX = 0
        fPtOZ = fOuterRadius
        fPtIX = 0
        fPtIZ = fInnerRadius
        fAngle = CSng((360 / lSlices) * (Math.PI / 180))

        lIdx = 0
        ReDim lInnerIdx(lSlices - 1)
        lCurrentInner = 0

        For X = 0 To lSlices - 1
            'get our inner ring coordinates
            fTX = fPtIX : fTZ = fPtIZ
            fPtIX = CSng((fTX * Math.Cos(fAngle)) + (fTZ * Math.Sin(fAngle)))
            fPtIZ = -CSng((fTX * Math.Sin(fAngle)) - (fTZ * Math.Cos(fAngle)))

            pnt = CType(data.GetValue(lIdx), CustomVertex.PositionColored)
            With pnt
                .X = fPtIX
                .Y = 1
                .Z = fPtIZ
                .Color = oVisibleColor.ToArgb
            End With
            data.SetValue(pnt, lIdx)
            lInnerIdx(lCurrentInner) = lIdx
            lCurrentInner += 1

            lIdx += 1

            'get our outer ring coordinates
            fTX = fPtOX : fTZ = fPtOZ
            fPtOX = CSng((fTX * Math.Cos(fAngle)) + (fTZ * Math.Sin(fAngle)))
            fPtOZ = -CSng((fTX * Math.Sin(fAngle)) - (fTZ * Math.Cos(fAngle)))

            pnt = CType(data.GetValue(lIdx), CustomVertex.PositionColored)
            With pnt
                .X = fPtOX
                .Y = 0
                .Z = fPtOZ
                .Color = oShroudColor.ToArgb
            End With
            data.SetValue(pnt, lIdx)
            lIdx += 1

        Next X
        pnt = CType(data.GetValue(lIdx), CustomVertex.PositionColored)
        pnt.X = 0 : pnt.Y = 1 : pnt.Z = 0 : pnt.Color = oVisibleColor.ToArgb
        data.SetValue(pnt, lIdx)
        lCenterIndex = lIdx
        oTmp.VertexBuffer.Unlock()

        'now set our index data...
        lIdx = 0
        For X = 0 To lSlices - 1
            iIndices(lIdx) = CShort(X * 2) : lIdx += 1
            iIndices(lIdx) = CShort((X * 2) + 1) : lIdx += 1

            If X <> lSlices - 1 Then
                iIndices(lIdx) = CShort((X * 2) + 2) : lIdx += 1

                iIndices(lIdx) = CShort((X * 2) + 2) : lIdx += 1
                iIndices(lIdx) = CShort((X * 2) + 1) : lIdx += 1
                iIndices(lIdx) = CShort((X * 2) + 3) : lIdx += 1
            Else
                iIndices(lIdx) = 0 : lIdx += 1

                iIndices(lIdx) = 0 : lIdx += 1
                iIndices(lIdx) = CShort((X * 2) + 1) : lIdx += 1
                iIndices(lIdx) = 1 : lIdx += 1
            End If
        Next X

        'Now, set our inner index data
        For X = 0 To lSlices - 1
            iIndices(lIdx) = CShort(lCenterIndex)
            lIdx += 1
            iIndices(lIdx) = CShort(lInnerIdx(X))
            lIdx += 1
            If X <> lSlices - 1 Then
                iIndices(lIdx) = CShort(lInnerIdx(X + 1))
            Else
                iIndices(lIdx) = CShort(lInnerIdx(0))
            End If
            lIdx += 1
        Next X
        oTmp.IndexBuffer.SetData(iIndices, 0, LockFlags.None)

        ' Optimize the mesh for this graphics card's vertex cache 
        ' so when rendering the mesh's triangle list the vertices will 
        ' cache hit more often so it won't have to re-execute the vertex shader 
        ' on those vertices so it will improve perf.     
        Dim adj() As Int32 = oTmp.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
        oTmp.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
        Erase adj

        data = Nothing

        Device.IsUsingEventHandlers = True

        Return oTmp
    End Function

    Private Sub SetVertexBufferDataDirtyOnly()
        'to contain the loadvertexbuffer2 code, essentially
        Try
            If mbSetVertexBufferDataCalled = False Then
                SetVertexBufferData()
                Return
            End If

            Dim k As Int32 = 0
            Dim oVerts(((mlVertsPerQuad + 1) * (mlVertsPerQuad + 1)) - 1) As terVert
            Dim lHlfQuad As Int32 = mlQuads \ 2

            Dim lQuadsPerFOWQuad As Int32 = CInt(Math.Floor(16384 / (mlVertsPerQuad * CellSpacing)))
            Dim fFOWTexUsage As Single = (lQuadsPerFOWQuad * mlVertsPerQuad * CellSpacing) / 16384.0F
            Device.IsUsingEventHandlers = False

            For lQuadX As Int32 = 0 To mlQuads - 1
                For lQuadZ As Int32 = 0 To mlQuads - 1
                    Dim lQuadIdx As Int32 = (lQuadZ * mlQuads) + lQuadX

                    If myQuadVBDirty(lQuadIdx) = 0 Then Continue For

                    If moQuadVB(lQuadIdx) Is Nothing = False Then moQuadVB(lQuadIdx).Dispose()
                    moQuadVB(lQuadIdx) = Nothing

                    If muSettings.BumpMapTerrain = True Then
                        moQuadVB(lQuadIdx) = New VertexBuffer(New terVertTBN().GetType, oVerts.Length, moDevice, Usage.WriteOnly, terVertFmtTBN, Pool.Managed)
                    Else
                        moQuadVB(lQuadIdx) = New VertexBuffer(New terVert().GetType, oVerts.Length, moDevice, Usage.WriteOnly, terVertFmt, Pool.Managed)
                    End If

                    Dim lVertX As Int32 = (lQuadX * mlVertsPerQuad)     'the first vert in this quad
                    Dim lVertZ As Int32 = (lQuadZ * mlVertsPerQuad)
                    Dim lOffsetX As Int32 = (lQuadX - lHlfQuad) * (mlVertsPerQuad * Me.CellSpacing)
                    Dim lOffsetZ As Int32 = (lQuadZ - lHlfQuad) * (mlVertsPerQuad * Me.CellSpacing)

                    'ok, we have our verts for this quad
                    k = 0
                    Dim iMaxI As Int32 = mlVertsPerQuad
                    Dim iMaxJ As Int32 = mlVertsPerQuad
                    If lQuadX = mlQuads - 1 Then iMaxJ -= 1
                    If lQuadZ = mlQuads - 1 Then iMaxI -= 1

                    For i As Int32 = 0 To iMaxI 'mlVertsPerQuad - 1
                        For j As Int32 = 0 To iMaxJ 'mlVertsPerQuad - 1
                            Dim lTmpIdx As Int32 = ((lVertZ + i) * Width) + (j + lVertX)
                            With oVerts(k)
                                .p = moVerts(lTmpIdx).p
                                .tu = moVerts(lTmpIdx).tu
                                .tv = moVerts(lTmpIdx).tv
                                .tw = moVerts(lTmpIdx).tw
                                .tu2 = moVerts(lTmpIdx).tu2
                                .tv2 = moVerts(lTmpIdx).tv2
                                .tw2 = 0
                            End With

                            k += 1
                        Next j
                    Next i

                    ComputeNormals2(oVerts, lVertX, lVertZ)

                    If muSettings.BumpMapTerrain = True Then
                        'If lQuadX = 8 AndAlso lQuadZ = 20 Then Stop
                        Dim uTmpVerts() As terVertTBN = DoTangentCalcs(oVerts, mlBaseIndices)
                        'WriteTmpVert(uTmpVerts, lQuadX, lQuadZ, True)
                        moQuadVB(lQuadIdx).SetData(uTmpVerts, 0, LockFlags.None)
                    Else
                        moQuadVB(lQuadIdx).SetData(oVerts, 0, LockFlags.None)
                    End If
                    myQuadVBDirty(lQuadIdx) = 0

                Next lQuadZ
            Next lQuadX
        Catch
            SetVertexBufferData()
        End Try

        Device.IsUsingEventHandlers = True
    End Sub

    Private mbSetVertexBufferDataCalled As Boolean = False
    Private Sub SetVertexBufferData()
        'to contain the loadvertexbuffer2 code, essentially
        mbSetVertexBufferDataCalled = True
        Device.IsUsingEventHandlers = False

        Try
            ReDim moQuadVB((mlQuads * mlQuads) - 1)
            ReDim myQuadVBDirty((mlQuads * mlQuads) - 1)

            Dim k As Int32 = 0
            Dim oVerts(((mlVertsPerQuad + 1) * (mlVertsPerQuad + 1)) - 1) As terVert
            Dim lTmp As Int32
            Dim lHlfQuad As Int32 = mlQuads \ 2
            Dim lCellRows As Int32 = mlVertsPerQuad - 1
            Dim lCellCols As Int32 = mlVertsPerQuad - 1
            Dim lTris As Int32 = (lCellRows + 1) * (lCellCols + 1) * 2
            Dim lIndices(lTris * 3) As Int32

            'Dim lTempIndices(lTris * 3) As Int32
            ReDim mlBaseIndices(lTris * 3)

            mlTris = lTris
            mlTris_XZ = lCellRows * lCellCols * 2
            mlTris_XOrZ = (lCellRows + 1) * lCellCols * 2
            Dim lQuadsPerFOWQuad As Int32 = CInt(Math.Floor(16384 / (mlVertsPerQuad * CellSpacing)))
            Dim fFOWTexUsage As Single = (lQuadsPerFOWQuad * mlVertsPerQuad * CellSpacing) / 16384.0F


            'generate the indices for this quad
            moQuadIB = New IndexBuffer(lTmp.GetType, lIndices.Length, moDevice, Usage.WriteOnly, Pool.Managed)
            k = 0
            Dim lPerRowOffset As Int32 = mlVertsPerQuad + 1
            For i As Int32 = 0 To lCellRows '- 1
                For j As Int32 = 0 To lCellCols '- 1
                    lIndices(k) = i * lPerRowOffset + j
                    lIndices(k + 1) = i * lPerRowOffset + j + 1
                    lIndices(k + 2) = (i + 1) * lPerRowOffset + j
                    lIndices(k + 3) = (i + 1) * lPerRowOffset + j
                    lIndices(k + 4) = i * lPerRowOffset + j + 1
                    lIndices(k + 5) = (i + 1) * lPerRowOffset + j + 1
                    k += 6
                Next j
            Next i

            For X As Int32 = 0 To lIndices.GetUpperBound(0)
                mlBaseIndices(X) = lIndices(X)
            Next X

            moQuadIB.SetData(lIndices, 0, LockFlags.None)
            moQuad_RIB = New IndexBuffer(lTmp.GetType, lIndices.Length, moDevice, Usage.WriteOnly, Pool.Managed)
            k = 0
            lPerRowOffset = mlVertsPerQuad
            For i As Int32 = 0 To lCellRows '- 1
                For j As Int32 = 0 To lCellCols - 1
                    lIndices(k) = i * lPerRowOffset + j
                    lIndices(k + 1) = i * lPerRowOffset + j + 1
                    lIndices(k + 2) = (i + 1) * lPerRowOffset + j
                    lIndices(k + 3) = (i + 1) * lPerRowOffset + j
                    lIndices(k + 4) = i * lPerRowOffset + j + 1
                    lIndices(k + 5) = (i + 1) * lPerRowOffset + j + 1
                    k += 6
                Next j
            Next i
            moQuad_RIB.SetData(lIndices, 0, LockFlags.None)


            For lQuadX As Int32 = 0 To mlQuads - 1
                For lQuadZ As Int32 = 0 To mlQuads - 1
                    Dim lQuadIdx As Int32 = (lQuadZ * mlQuads) + lQuadX

                    If moQuadVB(lQuadIdx) Is Nothing = False Then moQuadVB(lQuadIdx).Dispose()
                    moQuadVB(lQuadIdx) = Nothing

                    If muSettings.BumpMapTerrain = True Then
                        moQuadVB(lQuadIdx) = New VertexBuffer(New terVertTBN().GetType, oVerts.Length, moDevice, Usage.WriteOnly, terVertFmtTBN, Pool.Managed)
                    Else
                        moQuadVB(lQuadIdx) = New VertexBuffer(New terVert().GetType, oVerts.Length, moDevice, Usage.WriteOnly, terVertFmt, Pool.Managed)
                    End If

                    Dim lVertX As Int32 = (lQuadX * mlVertsPerQuad)     'the first vert in this quad
                    Dim lVertZ As Int32 = (lQuadZ * mlVertsPerQuad)
                    Dim lOffsetX As Int32 = (lQuadX - lHlfQuad) * (mlVertsPerQuad * Me.CellSpacing)
                    Dim lOffsetZ As Int32 = (lQuadZ - lHlfQuad) * (mlVertsPerQuad * Me.CellSpacing)

                    'ok, we have our verts for this quad
                    k = 0
                    Dim iMaxI As Int32 = mlVertsPerQuad
                    Dim iMaxJ As Int32 = mlVertsPerQuad
                    If lQuadX = mlQuads - 1 Then iMaxJ -= 1
                    If lQuadZ = mlQuads - 1 Then iMaxI -= 1

                    For i As Int32 = 0 To iMaxI 'mlVertsPerQuad - 1
                        For j As Int32 = 0 To iMaxJ 'mlVertsPerQuad - 1
                            Dim lTmpIdx As Int32 = ((lVertZ + i) * Width) + (j + lVertX)
                            With oVerts(k)
                                .p = moVerts(lTmpIdx).p
                                .tu = moVerts(lTmpIdx).tu
                                .tv = moVerts(lTmpIdx).tv
                                .tw = moVerts(lTmpIdx).tw
                                .tu2 = moVerts(lTmpIdx).tu2
                                .tv2 = moVerts(lTmpIdx).tv2
                                .tw2 = 0
                            End With

                            k += 1
                        Next j
                    Next i

                    ComputeNormals2(oVerts, lVertX, lVertZ)

                    If muSettings.BumpMapTerrain = True Then
                        'If lQuadX = 8 AndAlso lQuadZ = 20 Then Stop
                        Dim uTmpVerts() As terVertTBN = DoTangentCalcs(oVerts, mlBaseIndices)
                        'WriteTmpVert(uTmpVerts, lQuadX, lQuadZ, False)
                        moQuadVB(lQuadIdx).SetData(uTmpVerts, 0, LockFlags.None)
                    Else
                        moQuadVB(lQuadIdx).SetData(oVerts, 0, LockFlags.None)
                    End If
                    myQuadVBDirty(lQuadIdx) = 0

                Next lQuadZ
            Next lQuadX
        Catch
            mbSetVertexBufferDataCalled = False
        End Try

        Device.IsUsingEventHandlers = True
    End Sub
    'Public Sub SaveToFile()
    '    Dim oFS As New IO.FileStream("C:\terrain_" & Now.Minute & "_" & Now.Second & ".txt", IO.FileMode.Create)

    '    oFS.WriteByte(Me.WaterHeight)
    '    oFS.Write(HeightMap, 0, HeightMap.Length)
    '    oFS.Close()
    '    oFS.Dispose()
    'End Sub
End Class
 