Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class EpicaResourceManager

    Private Const ml_SHIELD_MESH_OFFSET As Int32 = 65535I

    Public Enum eGetTextureType As Byte
        NoSpecifics = 0     'indicates that the texture is a general purpose texture, no resizing used
        ModelTexture        'use the sizing from Model Texture Res
        UserInterface       'indicates that the texture is a user interface texture and will need to use the Purple Chroma Key
        StartupTexture
    End Enum

    Private moMeshes() As EpicaMesh
    Public MeshUB As Int32 = -1
    Private mlMeshIdx() As Int32

    Private moTextures() As Texture
    Public TextureUB As Int32 = -1
    Private msTextureNames() As String
    Private mlTextureType() As eGetTextureType
    Private msTexPak() As String

    Private msAppPath As String
    Private msMeshPath As String
    Private msTexturePath As String

    Private moDevice As Device

    Public ReadOnly Property MeshPath() As String
        Get
            Return msMeshPath
        End Get
    End Property

    Public Sub New(ByRef oD3DDevice As Device)
        Dim oINI As New InitFile()
        Dim sTemp As String

        moDevice = oD3DDevice
        msAppPath = AppDomain.CurrentDomain.BaseDirectory
        If Right$(msAppPath, 1) <> "\" Then msAppPath = msAppPath & "\"

        sTemp = oINI.GetString("GRAPHICS", "MeshPath", "Meshes\")
        msMeshPath = msAppPath & sTemp
        sTemp = oINI.GetString("GRAPHICS", "TexturePath", "Textures\")
        msTexturePath = msAppPath & sTemp
    End Sub

    Public Function GetMesh(ByVal lModelID As Int32) As EpicaMesh
        Dim X As Int32
        Dim lUnused As Int32 = -1

        For X = 0 To MeshUB
            If mlMeshIdx(X) = lModelID Then
                Return moMeshes(X)
            ElseIf lUnused = -1 AndAlso mlMeshIdx(X) = -1 Then
                lUnused = X
            End If
        Next X

        If lUnused = -1 Then
            'need to load new mesh and redim arrays
            MeshUB += 1
            lUnused = MeshUB
            ReDim Preserve moMeshes(MeshUB)
            ReDim Preserve mlMeshIdx(MeshUB)
        End If
        LoadMesh(lUnused, lModelID)

        Return moMeshes(lUnused)
    End Function

    Private Sub LoadMesh(ByVal lIndex As Int32, ByVal lModelID As Int32)
		'Dim oINI As InitFile = New InitFile(msMeshPath & "models.dat")
        Dim sModel As String
        Dim mtrlBuffer() As ExtendedMaterial = Nothing
        Dim X As Int32
        Dim sTemp As String
        'Dim lBoxSize As Int32

        Dim adj() As Int32

        Dim sPostFix As String

        Dim sModelHdr As String = "MODEL_" & lModelID

        Dim oMem As IO.MemoryStream
        Dim sPakFile As String
        Dim sModelPak As String
        'TODO: When we go live, we will want to make all of this stuff part of a message downloaded from the 
        ' Primary so that no one can hack them....

        'prepare our object...
        mlMeshIdx(lIndex) = lModelID

		sModel = "CommandCenter.x"
		sPakFile = ""
		sModelPak = "3DObjects.pak"

        moMeshes(lIndex) = New EpicaMesh()

        If sModel <> "" Then
            'sModel = msMeshPath & sModel
            'now set it up
            With moMeshes(lIndex)

                'If we are loading the mesh, then we need to load the shield mesh too
				'X = CInt(Val(oINI.GetString(sModelHdr, "ShieldSphereSize", "0")))
				'If X <> 0 Then .oShieldMesh = CreateShieldSphere(lModelID)

				'.ShieldXZRadius = CSng(Math.Max(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchX", "1")), Val(oINI.GetString(sModelHdr, "ShieldSphereStretchZ", "1"))))
				.ShieldXZRadius *= Math.Max(X, 1)

                oMem = GetResourceStream(sModel, msMeshPath & sModelPak)

                Device.IsUsingEventHandlers = False
                If oMem Is Nothing = False Then
                    .oMesh = Mesh.FromStream(oMem, MeshFlags.Managed, moDevice, mtrlBuffer)
                    oMem.Close()
                    oMem = Nothing
                Else
                    'TODO: we will want this to be more... graceful...
                    .oMesh = Mesh.Sphere(moDevice, 32, 32, 32)
                    .NumOfMaterials = 0
                    Exit Sub
                End If

                'Now, because our engine uses Normals and lighting, let's get that out of the way
                If (.oMesh.VertexFormat And VertexFormats.Normal) = 0 Then
                    Dim oTmpMesh As Mesh = .oMesh.Clone(.oMesh.Options.Value, .oMesh.VertexFormat Or VertexFormats.Normal, moDevice)
                    oTmpMesh.ComputeNormals()
                    .oMesh.Dispose()
                    .oMesh = oTmpMesh
                    oTmpMesh = Nothing
                End If

                ' Optimize the mesh for this graphics card's vertex cache 
                ' so when rendering the mesh's triangle list the vertices will 
                ' cache hit more often so it won't have to re-execute the vertex shader 
                ' on those vertices so it will improve perf.     
                adj = .oMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
                .oMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
                Erase adj

                Device.IsUsingEventHandlers = True

                ''Determine if we are using the new model render method or the old
                'If GFXEngine.mbSupportsNewModelMethod = True Then
                '    If sPakFile = "" Then sPakFile = "textures.pak"
                '    Dim lTemp As Int32 = sPakFile.LastIndexOf("."c)
                '    If lTemp < 1 Then
                '        sPakFile = "textures.pak"
                '        lTemp = sPakFile.LastIndexOf("."c)
                '    End If
                '    sPakFile = sPakFile.Substring(0, lTemp) & "2" & sPakFile.Substring(lTemp)

                '    'Initialize our materials and textures using the new method
                '    .NumOfMaterials = mtrlBuffer.Length
                '    ReDim .Materials(.NumOfMaterials - 1)
                '    ReDim .Textures(.NumOfMaterials - 1)
                '    ReDim .sTextures(.NumOfMaterials - 1)
                '    For X = 0 To mtrlBuffer.Length - 1
                '        .Materials(X) = mtrlBuffer(X).Material3D
                '        .Materials(X).Diffuse = Color.White
                '        .Materials(X).Ambient = .Materials(X).Diffuse
                '        If mtrlBuffer(X).TextureFilename <> "" Then
                '            sTemp = mtrlBuffer(X).TextureFilename
                '            sPostFix = Mid$(sTemp, sTemp.Length - 3, 4)
                '            sTemp = Mid$(sTemp, 1, sTemp.Length - 4) & ".dds"
                '            .Textures(X) = GetTexture(sTemp, eGetTextureType.ModelTexture, sPakFile)
                '            .sTextures(X) = sTemp
                '        End If
                '    Next X
                'Else
                'now initialize our materials and textures using the old method
                .NumOfMaterials = mtrlBuffer.Length
                ReDim .Materials(.NumOfMaterials - 1)
                ReDim .Textures((.NumOfMaterials * 4) - 1)          '*4, 1 for each relationship state
                ReDim .sTextures((.NumOfMaterials * 4) - 1)
                'Now load our textures and materials
                For X = 0 To mtrlBuffer.Length - 1
                    .Materials(X) = mtrlBuffer(X).Material3D
                    .Materials(X).Diffuse = Color.White
                    .Materials(X).Ambient = .Materials(X).Diffuse
                    If mtrlBuffer(X).TextureFilename <> "" Then
                        'sTemp is the NEUTRAL version
                        sTemp = mtrlBuffer(X).TextureFilename
                        sPostFix = Mid$(sTemp, sTemp.Length - 3, 4)

                        If sPostFix.ToUpper = ".DDS" Then sPostFix = ".bmp"
                        sTemp = Mid$(sTemp, 1, sTemp.Length - 4)

                        'Now, load our textures
                        .Textures((X * 4)) = GetTexture(sTemp & sPostFix, eGetTextureType.ModelTexture, sPakFile)
                        .sTextures(X * 4) = sTemp & sPostFix
                        .Textures((X * 4) + 1) = GetTexture(sTemp & "_mine" & sPostFix, eGetTextureType.ModelTexture, sPakFile)
                        .sTextures((X * 4) + 1) = sTemp & "_mine" & sPostFix
                        .Textures((X * 4) + 2) = GetTexture(sTemp & "_ally" & sPostFix, eGetTextureType.ModelTexture, sPakFile)
                        .sTextures((X * 4) + 2) = sTemp & "_ally" & sPostFix
                        .Textures((X * 4) + 3) = GetTexture(sTemp & "_enemy" & sPostFix, eGetTextureType.ModelTexture, sPakFile)
                        .sTextures((X * 4) + 3) = sTemp & "_enemy" & sPostFix
                    End If
                Next X
                'End If

                'Now, check if there is a Turret mesh...
				'sTemp = oINI.GetString(sModelHdr, "TurretFileName", "")
				'If sTemp <> "" Then
				'    sModel = sTemp

				'    Dim oTurretMatBuff() As ExtendedMaterial = Nothing

				'    oMem = GetResourceStream(sModel, msMeshPath & "3DObjects.pak")
				'    If oMem Is Nothing = False Then
				'        '.oTurretMesh = Mesh.FromFile(sModel, MeshFlags.Managed, moDevice, oTurretMatBuff)

				'        .oTurretMesh = Mesh.FromStream(oMem, MeshFlags.Managed, moDevice, oTurretMatBuff)

				'        'Release our memory buffer
				'        oMem.Close()
				'        oMem = Nothing

				'        'Ensure lighting and normals and stuff
				'        If (.oTurretMesh.VertexFormat And VertexFormats.Normal) = 0 Then
				'            Dim oTmpMesh As Mesh = .oTurretMesh.Clone(.oTurretMesh.Options.Value, .oTurretMesh.VertexFormat Or VertexFormats.Normal, moDevice)
				'            oTmpMesh.ComputeNormals()
				'            .oTurretMesh.Dispose()
				'            .oTurretMesh = oTmpMesh
				'            oTmpMesh = Nothing
				'        End If
				'        'Do some optimization
				'        adj = .oTurretMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
				'        .oTurretMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
				'        Erase adj

				'        'Now, it is assumed that a turret mesh uses the same materials/textures as the parent
				'        'So, we do not load them again... instead, simply get the remaining details
				'        .lTurretZOffset = CInt(Val(oINI.GetString(sModelHdr, "TurretOffsetZ", "0")))
				'        .bTurretMesh = True
				'    End If
				'End If
            End With
        Else
            'return a sphere... hehe
            moMeshes(lIndex).oMesh = Mesh.Sphere(moDevice, 32, 32, 32)
            moMeshes(lIndex).NumOfMaterials = 0
        End If

		'oINI = Nothing
    End Sub

    Public Sub DeleteMesh(ByVal lModelID As Int32)
        Dim X As Int32
        For X = 0 To MeshUB
            If mlMeshIdx(X) = lModelID Then
                mlMeshIdx(X) = -1
                moMeshes(X) = Nothing
            End If
        Next X
    End Sub

    Public Function GetTexture(ByVal sTextureName As String, ByVal lTextureType As eGetTextureType, Optional ByVal sPakFile As String = "") As Texture
        Dim X As Int32
        Dim lUnused As Int32 = -1

        For X = 0 To TextureUB
            If msTextureNames(X) = sTextureName Then
                If moTextures(X) Is Nothing OrElse moTextures(X).Disposed = True Then
                    moTextures(X) = Nothing
                    lUnused = X
                    Exit For
                End If
                Return moTextures(X)
            ElseIf msTextureNames(X) = "" AndAlso lUnused = -1 Then
                lUnused = X
            End If
        Next X

        If lUnused = -1 Then
            'need to load new texture and redim arrays
            TextureUB += 1
            lUnused = TextureUB
            ReDim Preserve moTextures(TextureUB)
            ReDim Preserve msTextureNames(TextureUB)
            ReDim Preserve mlTextureType(TextureUB)
            ReDim Preserve msTexPak(TextureUB)
        End If

        If sPakFile = "" Then
            LoadTexture(lUnused, sTextureName, lTextureType, "textures.pak")
        Else : LoadTexture(lUnused, sTextureName, lTextureType, sPakFile)
        End If



        Return moTextures(lUnused)
    End Function

    'Private Sub LoadTexture(ByVal lIndex As Int32, ByVal sName As String, ByVal lTextureType As eGetTextureType)
    '    Dim oINI As InitFile = New InitFile()
    '    Dim sTexture As String

    '    msTextureNames(lIndex) = sName

    '    sTexture = msTexturePath & sName
    '    Device.IsUsingEventHandlers = False
    '    Select Case lTextureType
    '        Case eGetTextureType.ModelTexture
    '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture, muSettings.ModelTextureResolution, muSettings.ModelTextureResolution, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Linear, 0)
    '        Case eGetTextureType.UserInterface
    '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture, 0, 0, 1, Usage.None, Format.Unknown, Pool.Default, Filter.None, Filter.None, System.Drawing.Color.FromArgb(255, 255, 0, 255).ToArgb)
    '        Case eGetTextureType.StartupTexture
    '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture, 0, 0, 0, Usage.None, Format.Unknown, Pool.Default, Filter.Triangle, Filter.Triangle, System.Drawing.Color.Black.ToArgb)
    '        Case Else       'includes NoSpecifics
    '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture)     'use defaults across the board
    '    End Select
    '    Device.IsUsingEventHandlers = True

    '    AddHandler moTextures(lIndex).Disposing, AddressOf TextureDispose
    'End Sub

    Private Sub LoadTexture(ByVal lIndex As Int32, ByVal sName As String, ByVal lTextureType As eGetTextureType, ByVal sTexturePak As String)
        Dim oINI As InitFile = New InitFile()
        'Dim sTexture As String

        msTextureNames(lIndex) = sName
        mlTextureType(lIndex) = lTextureType
        msTexPak(lIndex) = sTexturePak

        'Now, change the sName to what it needs to be... JPG = BMP
        If UCase$(sName).EndsWith(".JPG") = True Then
            sName = Replace$(sName, ".jpg", ".BMP", , , CompareMethod.Text)
        End If

        'sTexture = msTexturePath & sName
        Dim oMem As IO.MemoryStream = GetResourceStream(sName, msTexturePath & sTexturePak)
        If oMem Is Nothing = False Then
            'If UCase$(sName).EndsWith(".JPG") = True Then
            '    'Ok, this is a bit of a hack... make a new file
            '    sTexture = msTexturePath & "temp.jpg"

            '    Dim oFS As IO.FileStream = New IO.FileStream(sTexture, IO.FileMode.Create)
            '    Dim oBW As IO.BinaryWriter = New IO.BinaryWriter(oFS)
            '    Dim oBR As IO.BinaryReader = New IO.BinaryReader(oMem)

            '    oBW.Write(oBR.ReadBytes(CInt(oMem.Length)))
            '    oBR.Close() : oBR = Nothing
            '    oMem.Close() : oMem = Nothing
            '    oBW.Close() : oBW = Nothing
            '    oFS.Close() : oFS = Nothing

            '    Device.IsUsingEventHandlers = False
            '    Select Case lTextureType
            '        Case eGetTextureType.ModelTexture
            '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture, muSettings.ModelTextureResolution, muSettings.ModelTextureResolution, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Linear, 0)
            '        Case eGetTextureType.UserInterface
            '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture, 0, 0, 1, Usage.None, Format.Unknown, Pool.Default, Filter.None, Filter.None, System.Drawing.Color.FromArgb(255, 255, 0, 255).ToArgb)
            '        Case eGetTextureType.StartupTexture
            '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture, 0, 0, 0, Usage.None, Format.Unknown, Pool.Default, Filter.Triangle, Filter.Triangle, System.Drawing.Color.Black.ToArgb)
            '        Case Else       'includes NoSpecifics
            '            moTextures(lIndex) = TextureLoader.FromFile(moDevice, sTexture)     'use defaults across the board
            '    End Select
            '    Device.IsUsingEventHandlers = True
            'Else
            Device.IsUsingEventHandlers = False
            Select Case lTextureType
                Case eGetTextureType.ModelTexture
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, muSettings.ModelTextureResolution, muSettings.ModelTextureResolution, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Linear, 0)
                Case eGetTextureType.UserInterface
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, 0, 0, 1, Usage.None, Format.Unknown, Pool.Default, Filter.None, Filter.None, System.Drawing.Color.FromArgb(255, 255, 0, 255).ToArgb)
                Case eGetTextureType.StartupTexture
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, 0, 0, 0, Usage.None, Format.Unknown, Pool.Default, Filter.Triangle, Filter.Triangle, System.Drawing.Color.Black.ToArgb)
                Case Else       'includes NoSpecifics
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem)
            End Select
            Device.IsUsingEventHandlers = True
            'End If

            If moTextures(lIndex) Is Nothing = False Then
                AddHandler moTextures(lIndex).Disposing, AddressOf TextureDispose
            Else
                msTextureNames(lIndex) = ""
                moTextures(lIndex) = Nothing
            End If

        Else
            oMem = Nothing
            Exit Sub
        End If

    End Sub

    Private Sub TextureDispose(ByVal sender As Object, ByVal e As EventArgs)
        Dim X As Int32

        For X = 0 To TextureUB
            If msTextureNames(X) <> "" Then
                If moTextures(X) Is Nothing = False AndAlso moTextures(X).Disposed Then
                    msTextureNames(X) = ""
                    moTextures(X) = Nothing
                    mlTextureType(X) = eGetTextureType.NoSpecifics
                    msTexPak(X) = ""
                End If
            End If
        Next X
    End Sub

    Public Sub DeleteTexture(ByVal sName As String)
        Dim X As Int32
        For X = 0 To TextureUB
            If msTextureNames(X) = sName Then
                msTextureNames(X) = ""
                moTextures(X) = Nothing
                mlTextureType(X) = eGetTextureType.NoSpecifics
                msTexPak(X) = ""
            End If
        Next X
    End Sub

    'TODO: Added the bOffset flag to this function as an optional. We want to remove this once the models all default to 0
    Public Function CreateTexturedSphere(ByVal fRadius As Single, ByVal lSlices As Int32, ByVal lStacks As Int32, ByVal lWrapCount As Int32, Optional ByVal bOffsetY As Boolean = False) As Mesh

        Device.IsUsingEventHandlers = False

        Dim oTemp As Mesh = Mesh.Sphere(moDevice, fRadius, lSlices, lStacks)
        oTemp.ComputeNormals()
        Dim TexturedObject As Mesh = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, moDevice)
        'Dim TexturedObject As Mesh = oTemp.Clone(MeshFlags.Managed, VertexFormats.Position Or VertexFormats.Normal Or VertexFormats.Texture0, moDevice)

        ' Get the original mesh's vertex buffer.
        Dim ranks(0) As Integer
        ranks(0) = oTemp.NumberVertices
        Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

        ' Set the vertex buffer
        Dim data As System.Array = TexturedObject.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

        Dim phi As Single
        Dim u As Single
        Dim i As Integer

        For i = 0 To arr.Length - 1
            Dim pn As Direct3D.CustomVertex.PositionNormal = CType(arr.GetValue(i), CustomVertex.PositionNormal)
            Dim pnt As Direct3D.CustomVertex.PositionNormalTextured = CType(data.GetValue(i), CustomVertex.PositionNormalTextured)
            pnt.X = pn.X
            pnt.Y = pn.Y
            pnt.Z = pn.Z
            pnt.Nx = pn.Nx
            pnt.Ny = pn.Ny
            pnt.Nz = pn.Nz

            If lWrapCount = 0 Then
                phi = CSng(Math.Acos(pn.Nz))
                pnt.Tv = CSng(phi / Math.PI)
                u = CSng(Math.Acos(Math.Max(Math.Min(pnt.Ny / Math.Sin(phi), 1.0), -1.0)) / (2.0 * Math.PI))
                If pnt.Nx > 0 Then
                    pnt.Tu = u
                Else
                    pnt.Tu = 1 - u
                End If
            Else
                'TODO: At the moment, lWrapCount only determines if we wrap or not... I will want it to
                '  actually do more wraps with higher counts....
                pnt.Tu = CSng(Math.Asin(pnt.Nx) / gdPi + 0.5)
                pnt.Tv = CSng(Math.Asin(pnt.Ny) / gdPi + 0.5)
            End If

            If bOffsetY = True Then
                pnt.Y += fRadius
            End If

            data.SetValue(pnt, i)
        Next i

        TexturedObject.VertexBuffer.Unlock()
        oTemp.VertexBuffer.Unlock()

        ' Set the index buffer. 
        ranks(0) = oTemp.NumberFaces * 3
        arr = oTemp.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
        TexturedObject.IndexBuffer.SetData(arr, 0, LockFlags.None)

        ' Optimize the mesh for this graphics card's vertex cache 
        ' so when rendering the mesh's triangle list the vertices will 
        ' cache hit more often so it won't have to re-execute the vertex shader 
        ' on those vertices so it will improve perf.     
        Dim adj() As Int32 = TexturedObject.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
        TexturedObject.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
        Erase adj

        Device.IsUsingEventHandlers = True

        oTemp = Nothing
        arr = Nothing
        data = Nothing
        Return TexturedObject

    End Function

    Public Function CreateTexturedBox(ByVal fWidth As Single, ByVal fHeight As Single, ByVal fDepth As Single) As Mesh

        Device.IsUsingEventHandlers = False

        Dim oTemp As Mesh = Mesh.Box(moDevice, fWidth, fHeight, fDepth)
        oTemp.ComputeNormals()

        Dim TexturedObject As Mesh = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, moDevice)

        ' Get the original mesh's vertex buffer.
        Dim ranks(0) As Integer
        ranks(0) = oTemp.NumberVertices
        Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

        ' Set the vertex buffer
        Dim data As System.Array = TexturedObject.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

        'Dim u As Single
        Dim i As Integer

        Dim lFacing As Int32
        Dim bTopOrBottom As Boolean
        Dim fHalfHeight As Single = fHeight / 2

        For i = 0 To arr.Length - 1

            lFacing = CInt(Math.Floor(i / 4))
            bTopOrBottom = (lFacing = 1) OrElse (lFacing = 3)

            Dim pn As Direct3D.CustomVertex.PositionNormal = CType(arr.GetValue(i), CustomVertex.PositionNormal)
            Dim pnt As Direct3D.CustomVertex.PositionNormalTextured = CType(data.GetValue(i), CustomVertex.PositionNormalTextured)

            pnt.X = pn.X
            pnt.Y = pn.Y
            pnt.Z = pn.Z
            pnt.Nx = pn.Nx
            pnt.Ny = pn.Ny
            pnt.Nz = pn.Nz

            If bTopOrBottom = True Then
                If pnt.X < 0 Then
                    pnt.Tv = 0
                Else : pnt.Tv = 1
                End If

                If pnt.Z < 0 Then
                    pnt.Tu = 0
                Else : pnt.Tu = 1
                End If
            Else
                If pnt.Y < 0 Then
                    pnt.Tv = 0
                Else : pnt.Tv = 1
                End If

                If pnt.X < 0 Then
                    If pnt.Z < 0 Then
                        pnt.Tu = 0
                    Else : pnt.Tu = 1
                    End If
                Else
                    If pnt.Z < 0 Then
                        pnt.Tu = 1
                    Else : pnt.Tu = 0
                    End If
                End If
            End If

            'TODO: offset y here, but we may want to remove this later
            pnt.Y += fHalfHeight

            data.SetValue(pnt, i)
        Next i

        TexturedObject.VertexBuffer.Unlock()
        oTemp.VertexBuffer.Unlock()

        ' Set the index buffer. 
        ranks(0) = oTemp.NumberFaces * 3
        arr = oTemp.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
        TexturedObject.IndexBuffer.SetData(arr, 0, LockFlags.None)

        ' Optimize the mesh for this graphics card's vertex cache 
        ' so when rendering the mesh's triangle list the vertices will 
        ' cache hit more often so it won't have to re-execute the vertex shader 
        ' on those vertices so it will improve perf.     
        Dim adj() As Int32 = TexturedObject.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
        TexturedObject.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
        Erase adj

        Device.IsUsingEventHandlers = True

        oTemp = Nothing
        arr = Nothing
        data = Nothing
        Return TexturedObject

    End Function

    Public Function CreatePlanetRing(ByVal fInnerRadius As Single, ByVal fOuterRadius As Single, ByVal lSlices As Int32) As Mesh
        Dim fPtOX As Single
        Dim fPtOZ As Single
        Dim fPtIX As Single
        Dim fPtIZ As Single
        Dim fAngle As Single
        Dim X As Int32
        Dim lIdx As Int32

        Dim lVerts As Int32 = lSlices * 2
        Dim lRanks(0) As Int32

        Dim fTu As Single
        Dim pnt As CustomVertex.PositionNormalTextured

        Dim fTX As Single
        Dim fTZ As Single

        Dim iIndices() As Int16

        lRanks(0) = lVerts

        Device.IsUsingEventHandlers = False

        Dim oTmp As Mesh = New Mesh(lVerts * 3, lVerts, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, moDevice)
        Dim data As System.Array = oTmp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, lRanks)

        ReDim iIndices(lVerts * 3)

        fPtOX = 0
        fPtOZ = fOuterRadius
        fPtIX = 0
        fPtIZ = fInnerRadius
        fAngle = CSng((360 / lSlices) * (Math.PI / 180))

        lIdx = 0
        fTu = 0

        For X = 0 To lSlices - 1
            'get our inner ring coordinates
            fTX = fPtIX : fTZ = fPtIZ
            fPtIX = CSng((fTX * Math.Cos(fAngle)) + (fTZ * Math.Sin(fAngle)))
            fPtIZ = -CSng((fTX * Math.Sin(fAngle)) - (fTZ * Math.Cos(fAngle)))

            pnt = CType(data.GetValue(lIdx), CustomVertex.PositionNormalTextured)
            With pnt
                .X = fPtIX
                .Y = 0
                .Z = fPtIZ
                .Nx = 0
                .Ny = 1
                .Nz = 0
                .Tu = fTu
                .Tv = 0     'always 0
            End With
            data.SetValue(pnt, lIdx)

            lIdx += 1

            'get our outer ring coordinates
            fTX = fPtOX : fTZ = fPtOZ
            fPtOX = CSng((fTX * Math.Cos(fAngle)) + (fTZ * Math.Sin(fAngle)))
            fPtOZ = -CSng((fTX * Math.Sin(fAngle)) - (fTZ * Math.Cos(fAngle)))

            pnt = CType(data.GetValue(lIdx), CustomVertex.PositionNormalTextured)
            With pnt
                .X = fPtOX
                .Y = 0
                .Z = fPtOZ
                .Nx = 0
                .Ny = 1
                .Nz = 0
                .Tu = fTu
                .Tv = 1     'always 1
            End With
            data.SetValue(pnt, lIdx)
            lIdx += 1

            If fTu = 0 Then fTu = 1 Else fTu = 0
        Next X
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
        oTmp.IndexBuffer.SetData(iIndices, 0, LockFlags.None)

        oTmp.ComputeNormals()

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

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Function GetVolTerrainTexture(ByVal sTextureName As String) As VolumeTexture
        'NOTE: This particular function is a SCRATCH surface, meaning, when it is gone, it is gone.

        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap Then
            Dim oMem As IO.MemoryStream = GetResourceStream(sTextureName, msTexturePath & "VolTerr.pak")

            Device.IsUsingEventHandlers = False
            'Dim oTmp As VolumeTexture = TextureLoader.FromVolumeFile(moDevice, msTexturePath & sTextureName, 0, 0, 0, 1, Usage.None, Format.X8R8G8B8, Pool.Managed, Filter.Linear, Filter.Linear, 0)
            Dim oTmp As VolumeTexture = TextureLoader.FromVolumeStream(moDevice, oMem, 0, 0, 0, 1, Usage.None, Format.X8R8G8B8, Pool.Managed, Filter.Linear, Filter.Linear, 0)
            Device.IsUsingEventHandlers = True

            oMem.Close()
            oMem = Nothing

            Return oTmp
        Else
            Return Nothing
        End If
    End Function

    Public Function GetVolNormalTexture(ByVal sTextureName As String) As VolumeTexture
        'NOTE: This particular function is a SCRATCH surface, meaning, when it is gone, it is gone.

        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap Then
            Dim oMem As IO.MemoryStream = GetResourceStream(sTextureName, msTexturePath & "terrdin.pak")
            If oMem Is Nothing Then Return Nothing

            Device.IsUsingEventHandlers = False
            'Dim oTmp As VolumeTexture = TextureLoader.FromVolumeFile(moDevice, msTexturePath & sTextureName, 0, 0, 0, 1, Usage.None, Format.X8R8G8B8, Pool.Managed, Filter.Linear, Filter.Linear, 0)
            Dim oTmp As VolumeTexture = TextureLoader.FromVolumeStream(moDevice, oMem, 0, 0, 0, 1, Usage.None, Format.X8R8G8B8, Pool.Managed, Filter.Linear, Filter.Linear, 0)
            Device.IsUsingEventHandlers = True

            oMem.Close()
            oMem = Nothing

            Return oTmp
        Else
            Return Nothing
        End If
    End Function

    Public Function GetNVTerrainTexture(ByVal sTextureName As String) As Texture
        'NOTE: This particular function is a SCRATCH surface, meaning, when it is gone, it is gone.
        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = False Then
            Dim oMem As IO.MemoryStream = GetResourceStream(sTextureName, msTexturePath & "NVTerr.pak")

            Device.IsUsingEventHandlers = False
            'Dim oTmp As Texture = TextureLoader.FromFile(moDevice, msTexturePath & sTextureName)
            Dim oTmp As Texture = TextureLoader.FromStream(moDevice, oMem)
            Device.IsUsingEventHandlers = True

            oMem.Close()
            oMem = Nothing

            Return oTmp
        Else
            Return Nothing
        End If
    End Function

    Public Function CreateFOWDisk(ByVal oVisibleColor As System.Drawing.Color, ByVal oShroudColor As System.Drawing.Color) As Mesh
        Dim lSlices As Int32 = 16
        Dim fInnerRadius As Single = 0.9F
        Dim fOuterRadius As Single = 1.1F

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

    'Should only be called from within resource manager...
    Private Function CreateShieldSphere(ByVal lModelID As Int32) As Mesh
        'ByVal lWrapCount As Int32, ByVal fStretchX As Single, ByVal fStretchY As Single, ByVal fStretchZ As Single, ByVal fShiftX As Single, ByVal fShiftY As Single, ByVal fShiftZ As Single
        Dim oINI As InitFile = New InitFile(msMeshPath & "models.dat")
        Dim sModelHdr As String = "MODEL_" & lModelID

        Dim fRadius As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereSize", "0")))
        Dim fShiftX As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereShiftX", "0")))
        Dim fShiftY As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereShiftY", "0")))
        Dim fShiftZ As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereShiftZ", "0")))
        Dim fStretchX As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchX", "0")))
        Dim fStretchY As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchY", "0")))
        Dim fStretchZ As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchZ", "0")))

        oINI = Nothing

        'Dim phi As Single
        'Dim u As Single
        Dim i As Integer

        Device.IsUsingEventHandlers = False

        Dim oTemp As Mesh = Mesh.Sphere(moDevice, fRadius, 32, 32)      'TODO: we could make an improvement here by possibly reducing stacks/slices to 16
        oTemp.ComputeNormals()
        Dim TexturedObject As Mesh = New Mesh(oTemp.NumberFaces, oTemp.NumberVertices, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, moDevice)

        ' Get the original mesh's vertex buffer.
        Dim ranks(0) As Integer
        ranks(0) = oTemp.NumberVertices
        Dim arr As System.Array = oTemp.VertexBuffer.Lock(0, (New CustomVertex.PositionNormal()).GetType(), LockFlags.None, ranks)

        ' Set the vertex buffer
        Dim data As System.Array = TexturedObject.VertexBuffer.Lock(0, (New CustomVertex.PositionNormalTextured()).GetType(), LockFlags.None, ranks)

        For i = 0 To arr.Length - 1
            Dim pn As Direct3D.CustomVertex.PositionNormal = CType(arr.GetValue(i), CustomVertex.PositionNormal)
            Dim pnt As Direct3D.CustomVertex.PositionNormalTextured = CType(data.GetValue(i), CustomVertex.PositionNormalTextured)
            pnt.X = pn.X + fShiftX
            pnt.Y = pn.Y + fShiftY
            pnt.Z = pn.Z + fShiftZ
            pnt.Nx = pn.Nx
            pnt.Ny = pn.Ny
            pnt.Nz = pn.Nz

            'Assumed WrapCount = 1
            pnt.Tu = CSng(Math.Asin(pnt.Nx) / Math.PI + 0.5F)
            pnt.Tv = CSng(Math.Asin(pnt.Ny) / Math.PI + 0.5F)

            'Now, do our stretch
            pnt.X *= fStretchX
            pnt.Y *= fStretchY
            pnt.Z *= fStretchZ

            data.SetValue(pnt, i)
        Next i

        TexturedObject.VertexBuffer.Unlock()
        oTemp.VertexBuffer.Unlock()

        ' Set the index buffer. 
        ranks(0) = oTemp.NumberFaces * 3
        arr = oTemp.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
        TexturedObject.IndexBuffer.SetData(arr, 0, LockFlags.None)

        ' Optimize the mesh for this graphics card's vertex cache 
        ' so when rendering the mesh's triangle list the vertices will 
        ' cache hit more often so it won't have to re-execute the vertex shader 
        ' on those vertices so it will improve perf.     
        Dim adj() As Int32 = TexturedObject.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
        TexturedObject.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
        Erase adj

        Device.IsUsingEventHandlers = True

        oTemp = Nothing
        arr = Nothing
        data = Nothing
        Return TexturedObject

    End Function

    Private Function GetResourceStream(ByVal sName As String, ByVal sResourceFile As String) As IO.MemoryStream
        If Exists(sResourceFile) = False Then Return Nothing

        Dim oFS As IO.FileStream = Nothing
        Dim oReader As IO.BinaryReader

        Dim yData() As Byte
        Dim yTOC() As Byte
        Dim lTOCLen As Int32

        Dim X As Int32
        Dim yEntryName(19) As Byte
        Dim sFile As String
        Dim lPos As Int32

        Dim lEntryStart As Int32 = -1
        Dim lEntryLen As Int32 = -1

        'Open the resource file
        Dim lAttemptCntr As Int32 = 0
        While oFS Is Nothing AndAlso lAttemptCntr < 100
            Try
                oFS = New IO.FileStream(sResourceFile, IO.FileMode.Open)
            Catch ex As Exception
                'do nothing
            End Try
            lAttemptCntr += 1
            Threading.Thread.Sleep(10)
        End While
        If oFS Is Nothing Then Return Nothing

        'Create our reader
        oReader = New IO.BinaryReader(oFS)

        'Get the first 4 bytes which is the unencrypted length of our Table of Contents portion
        yData = oReader.ReadBytes(4)
        lTOCLen = System.BitConverter.ToInt32(yData, 0)

        'read in the TOC contents
        yTOC = oReader.ReadBytes(lTOCLen)

        'Decrypt the TOC
        yTOC = DecBytes(yTOC)

        lPos = 0
        'Find out entry...
        For X = 0 To CInt(Math.Floor(lTOCLen / 28) - 1)
            Array.Copy(yTOC, X * 28, yEntryName, 0, 20)
            lPos += 20
            sFile = BytesToString(yEntryName)
            If UCase$(sFile) = UCase$(sName) Then
                lEntryStart = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                lEntryLen = System.BitConverter.ToInt32(yTOC, lPos) : lPos += 4
                Exit For
            Else
                lPos += 8
            End If
        Next X

        If lEntryStart > -1 AndAlso lEntryLen > -1 Then
            'Ok, found our entry... let's get the data
            'oReader.ReadBytes(lEntryStart - 1)  'should be right...
            oReader.BaseStream.Seek(lEntryStart + lTOCLen + 4, IO.SeekOrigin.Begin)
            Dim yEntry() As Byte
            yData = oReader.ReadBytes(lEntryLen)
            'Now, decrypt it
            yData = DecBytes(yData)
            ReDim yEntry(lEntryLen - 1)
            Array.Copy(yData, 0, yEntry, 0, lEntryLen)

            'Now, create our response
            Dim oMemStream As IO.MemoryStream = New IO.MemoryStream(yEntry.Length - 1)
            Dim oFSWriter As IO.BinaryWriter = New IO.BinaryWriter(oMemStream)
            oFSWriter.Write(yEntry)
            'Flush and Nothing, but do not close as it will close the base stream too
            oFSWriter.Flush()
            oFSWriter = Nothing

            'Now, seek the memory stream to the beginning
            oMemStream.Seek(0, IO.SeekOrigin.Begin)

            'Close everything else
            oReader.Close()
            oFS.Close()

            Return oMemStream
        Else
            oReader.Close()
            oFS.Close()
            Return Nothing
        End If

    End Function

    Private Function DecBytes(ByVal yBytes() As Byte) As Byte()
        Const ml_ENCRYPT_SEED As Int32 = 777

        'Now, we do the exact opposite...
        Dim lLen As Int32 = yBytes.GetUpperBound(0)
        Dim lKey As Int32
        'Dim lOffset As Int32
        Dim X As Int32
        Dim yFinal(lLen + 1) As Byte
        Dim lChrCode As Int32
        Dim lMod As Int32

        'Get our key value...
        lKey = yBytes(0)

        'set up our seed
        Rnd(-1)
        Call Randomize(ml_ENCRYPT_SEED + lKey)

        For X = 1 To lLen
            'Now, find out what we got here...
            lChrCode = yBytes(X)
            'now, subtract our value... 1 to 5
            lMod = CInt(Int(Rnd() * 5) + 1)
            lChrCode = lChrCode - lMod
            If lChrCode < 0 Then lChrCode = 256 + lChrCode
            yFinal(X - 1) = CByte(lChrCode)
        Next X
        DecBytes = yFinal
    End Function

    Private Function BytesToString(ByVal yBytes() As Byte) As String
        Dim lLen As Int32 = yBytes.Length
        Dim X As Int32

        For X = 0 To yBytes.Length - 1
            If yBytes(X) = 0 Then
                lLen = X
                Exit For
            End If
        Next X

        Return Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yBytes), 1, lLen)
    End Function

    Public Sub ReleaseDefaultPoolTextures()
        'TODO: 'release' whatever that means the default pool textures
    End Sub

    Public Sub RecreateDefaultPoolTextures()
        'TODO: 'recreate' the default pool textures
    End Sub

    Public Sub UpdateModelTextureResolution()
        For X As Int32 = 0 To TextureUB
            If msTextureNames(X) <> "" AndAlso mlTextureType(X) = eGetTextureType.ModelTexture Then

                Dim sTexture As String = msTextureNames(X)
                Dim sPak As String = msTexPak(X)

                If moTextures(X) Is Nothing = False Then moTextures(X).Dispose()
                moTextures(X) = Nothing
                LoadTexture(X, sTexture, eGetTextureType.ModelTexture, sPak)
            End If
        Next X

        For X As Int32 = 0 To MeshUB
            If mlMeshIdx(X) <> -1 AndAlso moMeshes(X) Is Nothing = False AndAlso moMeshes(X).Textures Is Nothing = False Then
                For Y As Int32 = 0 To moMeshes(X).Textures.GetUpperBound(0)
                    '*SHOULD* return everything
                    moMeshes(X).Textures(Y) = Me.GetTexture(moMeshes(X).sTextures(Y), eGetTextureType.ModelTexture, "")
                Next Y
            End If
        Next X
    End Sub

    Public Function LoadScratchMeshNoTextures(ByVal sMeshName As String, ByVal sMeshPak As String) As Mesh
        Dim mtrlBuffer() As ExtendedMaterial = Nothing
        Dim adj() As Int32

        Dim oMem As IO.MemoryStream
        Dim oMesh As Mesh

        oMem = GetResourceStream(sMeshName, msMeshPath & sMeshPak)

        Device.IsUsingEventHandlers = False
        If oMem Is Nothing = False Then
            oMesh = Mesh.FromStream(oMem, MeshFlags.Managed, moDevice, mtrlBuffer)
            oMem.Close()
            oMem = Nothing
        Else
            'TODO: we will want this to be more... graceful...
            oMesh = Mesh.Sphere(moDevice, 32, 32, 32)
            Return oMesh
        End If

        'Now, because our engine uses Normals and lighting, let's get that out of the way
        If (oMesh.VertexFormat And VertexFormats.Normal) = 0 Then
            Dim oTmpMesh As Mesh = oMesh.Clone(oMesh.Options.Value, oMesh.VertexFormat Or VertexFormats.Normal, moDevice)
            oTmpMesh.ComputeNormals()
            oMesh.Dispose()
            oMesh = oTmpMesh
            oTmpMesh = Nothing
        End If

        ' Optimize the mesh for this graphics card's vertex cache 
        ' so when rendering the mesh's triangle list the vertices will 
        ' cache hit more often so it won't have to re-execute the vertex shader 
        ' on those vertices so it will improve perf.     
        adj = oMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
        oMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
        Erase adj

        Device.IsUsingEventHandlers = True

        Return oMesh
    End Function
End Class
