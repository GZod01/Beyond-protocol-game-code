Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Structure BoundingSphere
    Public SphereCenter As Vector3
    Public SphereRadius As Single
End Structure

Public Structure BurnSide
    Public lBurnFXCnt As Int32

    Public vecBurnOffset() As Vector3
    Public lBurnPCnt() As Int32

End Structure

Public Structure FireFromLoc
    Public ySideID As Byte      'the arc we are relative to
    Public lOffsetX As Int32    'from unit locx
    Public lOffsetY As Int32    'from unit locy
    Public lOffsetZ As Int32    'from unit locz

    Public Function GetFireFromLoc(ByRef oObj As RenderObject) As Vector3
        'Dim fX As Single = lOffsetX
        'Dim fZ As Single = lOffsetZ
        'RotatePoint(0, 0, fX, fZ, oObj.LocAngle / 10.0F)
        'Return New Vector3(fX + oObj.LocX, lOffsetY + oObj.LocY, fZ + oObj.LocZ)

        Dim fX As Single = lOffsetX + oObj.LocX
        Dim fZ As Single = lOffsetZ + oObj.LocZ
        Dim iAngle As Short = oObj.LocAngle - 900S
        If iAngle < 0 Then iAngle += 3600S

        RotatePoint(CInt(oObj.LocX), CInt(oObj.LocZ), fX, fZ, iAngle / 10.0F)
        Return New Vector3(fX, lOffsetY + oObj.LocY, fZ)
        'Return New Vector3(lOffsetX + oObj.LocX, lOffsetY + oObj.LocY, lOffsetZ + oObj.LocZ)

    End Function
End Structure

Public Structure LightInfo

    Public lLightType As Int32              'Type of the light
    Public Color As System.Drawing.Color
    Public Source As Vector3                'The Coordinates of the texture
    Public Direction As Vector3             'The direction it's pointing
    Public bSearchLight As Boolean          'If true the Dest is altered over time to sweep from Start to End.
    Public SearchLightCycles As Int32       'How many cycles between start and finish
    Public SearchLightFormula As Int32      'A bitmask that tells me the format of travel.  straight line.  Do a sweeping arc.  etc.
    'One option might be from the normal Direction -> Start point then do a simple circle.  SLEnd is optional in this case.
    Public SearchLightStart As Vector3      'The vector of the endpoint of the light at Search start
    Public SearchLightEnd As Vector3        'The vector of the endpoint of the light at Search end

    Public LightTex As Texture
    Public oLight As Light
    Public iLightID As Int32
    Public TheSprite As Sprite
End Structure

Public Class GFXResourceManager

    Private Const ml_SHIELD_MESH_OFFSET As Int32 = 65535I

    Public Enum eGetTextureType As Byte
        NoSpecifics = 0     'indicates that the texture is a general purpose texture, no resizing used
        ModelTexture        'use the sizing from Model Texture Res
        UserInterface       'indicates that the texture is a user interface texture and will need to use the Purple Chroma Key
		StartupTexture
		IlluminationMap
		WaterTexture
    End Enum

    Private moMeshes() As BaseMesh
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

    Public moGHTex() As Texture

    Public Shared oDefaultSpecularTex As Texture
    Public Class SharedTexture
        Public sSearchableName As String        'just the ucase of sDiffuseTexture
        Public sDiffuseTexture As String
        Public sDiffusePak As String
        Public sNormalTexture(3) As String
        Public sNormalPak(3) As String
        Public sIllumTexture(1) As String
        Public sIllumPak(1) As String
        Public sSpecTexture As String
        Public sSpecPak As String

        Public oDiffuseTexture As Texture
        Public oNormalTexture(3) As Texture
        Public oIllumTexture(1) As Texture
        Public oSpecTexture As Texture
    End Class
    Private moSharedTex() As SharedTexture
    Private mlSharedTexUB As Int32 = -1
    Public Function GetSharedTexture(ByVal sName As String) As SharedTexture
        sName = sName.ToUpper
        For X As Int32 = 0 To mlSharedTexUB
            If moSharedTex(X).sSearchableName = sName Then Return moSharedTex(X)
        Next X
        Return Nothing
    End Function
    Private Sub LoadAllSharedTextures()
        'get our ini some how here...
        If UnpackLocalDataFile("md.pak", "sharedtex.txt") = False Then Return

        Dim sFile As String = msMeshPath
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "sharedtex.txt"

        oDefaultSpecularTex = GetTexture("test.dds", eGetTextureType.ModelTexture, "mttex.pak")

        Dim oINI As New InitFile(sFile)
        mlSharedTexUB = -1
        For X As Int32 = 1 To 100
            Dim sHdr As String = "TEX" & X
            Dim sDiff As String = oINI.GetString(sHdr, "Diffuse", "").Trim
            If sDiff = "" Then Exit For

            mlSharedTexUB += 1
            ReDim Preserve moSharedTex(mlSharedTexUB)
            moSharedTex(mlSharedTexUB) = New SharedTexture
            With moSharedTex(mlSharedTexUB)
                .sSearchableName = sDiff.ToUpper.Trim
                .sDiffuseTexture = sDiff
                .sDiffusePak = oINI.GetString(sHdr, "DiffusePak", "texturesdin.pak")
                For Y As Int32 = 1 To 4
                    .sNormalTexture(Y - 1) = oINI.GetString(sHdr, "Nrm" & Y, "")
                    .sNormalPak(Y - 1) = oINI.GetString(sHdr, "Nrm" & Y & "Pak", "")
                Next Y
                For Y As Int32 = 1 To 2
                    .sIllumTexture(Y - 1) = oINI.GetString(sHdr, "Illum" & Y, "")
                    .sIllumPak(Y - 1) = oINI.GetString(sHdr, "Illum" & Y & "Pak", "")
                Next Y
                .sSpecTexture = oINI.GetString(sHdr, "Spec", "")
                .sSpecPak = oINI.GetString(sHdr, "SpecPak", "")
            End With
        Next X
        oINI = Nothing

        For X As Int32 = 0 To mlSharedTexUB
            With moSharedTex(X)
                .oDiffuseTexture = GetTexture(.sDiffuseTexture, eGetTextureType.ModelTexture, .sDiffusePak)
                For Y As Int32 = 0 To 3
                    If .sNormalTexture(Y) <> "" Then .oNormalTexture(Y) = GetTexture(.sNormalTexture(Y), eGetTextureType.ModelTexture, .sNormalPak(Y))
                Next Y
                For Y As Int32 = 0 To 1
                    If .sIllumTexture(Y) <> "" Then .oIllumTexture(Y) = GetTexture(.sIllumTexture(Y), eGetTextureType.IlluminationMap, .sIllumPak(Y))
                Next Y
                If .sSpecTexture <> "" Then .oSpecTexture = GetTexture(.sSpecTexture, eGetTextureType.ModelTexture, .sSpecPak)
            End With
        Next X

        If Exists(msMeshPath & "sharedtex.txt") = True Then Kill(msMeshPath & "sharedtex.txt")

        ReDim moGHTex(2)
        moGHTex(0) = GetTexture("TC.dds", eGetTextureType.IlluminationMap, "gh.pak")
        moGHTex(1) = GetTexture("BH.dds", eGetTextureType.IlluminationMap, "gh.pak")
        moGHTex(2) = GetTexture("QSC.dds", eGetTextureType.IlluminationMap, "gh.pak")
    End Sub

    Public Sub ClearAllResources()
        For X As Int32 = 0 To MeshUB
            mlMeshIdx(X) = -1
            moMeshes(X) = Nothing
        Next X
        For X As Int32 = 0 To TextureUB
            msTextureNames(X) = ""
            moTextures(X) = Nothing
            mlTextureType(X) = eGetTextureType.NoSpecifics
            msTexPak(X) = ""
        Next X
        For X As Int32 = 0 To mlSharedTexUB
            moSharedTex(X) = Nothing
        Next X
        If moGHTex Is Nothing = False Then
            For X As Int32 = 0 To moGHTex.GetUpperBound(0)
                If moGHTex(X) Is Nothing = False Then moGHTex(X).Dispose()
                moGHTex(X) = Nothing
            Next X
        End If
    End Sub

    Public ReadOnly Property MeshPath() As String
        Get
            Return msMeshPath
        End Get
    End Property
    Public ReadOnly Property TexturePath() As String
        Get
            Return msTexturePath
        End Get
    End Property

    Public Sub New()
        Dim oINI As New InitFile()
        Dim sTemp As String

        msAppPath = AppDomain.CurrentDomain.BaseDirectory
        If Right$(msAppPath, 1) <> "\" Then msAppPath = msAppPath & "\"

        sTemp = oINI.GetString("GRAPHICS", "MeshPath", "Meshes\")
        msMeshPath = msAppPath & sTemp
        sTemp = oINI.GetString("GRAPHICS", "TexturePath", "Textures\")
        msTexturePath = msAppPath & sTemp

        LoadAllSharedTextures()
    End Sub

    Public Function ClearAndGetMesh(ByVal lModelID As Int32) As BaseMesh
        'MSC: We only care about the mesh portion of the modelid
        lModelID = (lModelID And 255)
        'ok, basically, something bad happened so we're going to clear the mesh if it is there and then reload it
        Dim oResult As BaseMesh = Nothing

        For X As Int32 = 0 To MeshUB
            If mlMeshIdx(X) = lModelID Then
                'ok, found it
                LoadMesh(X, lModelID)
                oResult = moMeshes(X)
                Exit For
            End If
        Next X

        If oResult Is Nothing Then Return GetMesh(lModelID)

        'Now, anyone that has this model id, needs to set their mesh to the new mesh
        Try
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                Dim lCurUB As Int32 = -1
                If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If oEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False AndAlso oEntity.oMesh Is Nothing = False AndAlso oEntity.oMesh.lModelID = lModelID Then
                            oEntity.oMesh = oResult
                        End If
                    End If
                Next X
            End If
        Catch
        End Try

        Return oResult
    End Function

    Public Function GetMesh(ByVal lModelID As Int32) As BaseMesh
        'MSC: we only care about the mesh portion of the modelid
        lModelID = (lModelID And 255)

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

    'Private Sub LoadMesh(ByVal lIndex As Int32, ByVal lModelID As Int32)
    '    Dim oINI As InitFile = New InitFile(msMeshPath & "models.dat")
    '    Dim sModel As String
    '    Dim mtrlBuffer() As ExtendedMaterial
    '    Dim X As Int32
    '    Dim sTemp As String
    '    Dim lBoxSize As Int32

    '    Dim adj() As Int32

    '    Dim sPostFix As String

    '    Dim sModelHdr As String = "MODEL_" & lModelID

    '    'TODO: When we go live, we will want to make all of this stuff part of a message downloaded from the 
    '    ' Primary so that no one can hack them....

    '    'prepare our object...
    '    moMeshes(lIndex) = New EpicaMesh()
    '    mlMeshIdx(lIndex) = lModelID

    '    sModel = oINI.GetString(sModelHdr, "FileName", "")

    '    If sModel <> "" Then
    '        sModel = msMeshPath & sModel
    '        If Dir$(sModel) <> "" Then
    '            'now set it up
    '            With moMeshes(lIndex)

    '                'If we are loading the mesh, then we need to load the shield mesh too
    '                X = Val(oINI.GetString(sModelHdr, "ShieldSphereSize", "0"))
    '                If X <> 0 Then .oShieldMesh = CreateShieldSphere(lModelID)

    '                .ShieldXZRadius = Math.Max(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchX", "1")), Val(oINI.GetString(sModelHdr, "ShieldSphereStretchZ", "1")))
    '                .ShieldXZRadius *= Math.Max(X, 1)

    '                moDevice.IsUsingEventHandlers = False
    '                .oMesh = .oMesh.FromFile(sModel, MeshFlags.Managed, moDevice, mtrlBuffer)
    '                .PlanetYAdjust = Val(oINI.GetString(sModelHdr, "PlanetYAdjust", "0"))
    '                .YMidPoint = Val(oINI.GetString(sModelHdr, "YMidPoint", "0"))
    '                .bLandBased = (Val(oINI.GetString(sModelHdr, "LandBased", "0")) <> 0)
    '                .lModelID = lModelID
    '                .RangeOffset = Val(oINI.GetString(sModelHdr, "RangeOffset", "0"))

    '                .EngineFXCnt = Val(oINI.GetString(sModelHdr, "EngineFXCnt", "0"))
    '                ReDim .EngineFXOffset(.EngineFXCnt)
    '                ReDim .EngineFXPCnt(.EngineFXCnt)
    '                For X = 0 To .EngineFXCnt - 1
    '                    .EngineFXOffset(X).X = oINI.GetString(sModelHdr, "Engine" & X & "_OffsetX", "0")
    '                    .EngineFXOffset(X).Y = oINI.GetString(sModelHdr, "Engine" & X & "_OffsetY", "0")
    '                    .EngineFXOffset(X).Z = oINI.GetString(sModelHdr, "Engine" & X & "_OffsetZ", "0")
    '                    .EngineFXPCnt(X) = oINI.GetString(sModelHdr, "Engine" & X & "_PCnt", "0")
    '                Next X

    '                'Now, because our engine uses Normals and lighting, let's get that out of the way
    '                If (.oMesh.VertexFormat And VertexFormats.Normal) = 0 Then
    '                    Dim oTmpMesh As Mesh = .oMesh.Clone(.oMesh.Options.Value, .oMesh.VertexFormat Or VertexFormats.Normal, moDevice)
    '                    oTmpMesh.ComputeNormals()
    '                    .oMesh.Dispose()
    '                    .oMesh = oTmpMesh
    '                    oTmpMesh = Nothing
    '                End If

    '                ' Optimize the mesh for this graphics card's vertex cache 
    '                ' so when rendering the mesh's triangle list the vertices will 
    '                ' cache hit more often so it won't have to re-execute the vertex shader 
    '                ' on those vertices so it will improve perf.     
    '                adj = .oMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
    '                .oMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
    '                Erase adj

    '                moDevice.IsUsingEventHandlers = True

    '                'now initialize our materials and textures
    '                .NumOfMaterials = mtrlBuffer.Length
    '                ReDim .Materials(.NumOfMaterials - 1)
    '                ReDim .Textures((.NumOfMaterials * 4) - 1)          '*4, 1 for each relationship state
    '                'Now load our textures and materials
    '                For X = 0 To mtrlBuffer.Length - 1
    '                    .Materials(X) = mtrlBuffer(X).Material3D
    '                    .Materials(X).Ambient = .Materials(X).Diffuse
    '                    If mtrlBuffer(X).TextureFilename <> "" Then
    '                        'sTemp is the NEUTRAL version
    '                        sTemp = mtrlBuffer(X).TextureFilename
    '                        sPostFix = Mid$(sTemp, Len(sTemp) - 3, 4)
    '                        sTemp = Mid$(sTemp, 1, Len(sTemp) - 4)

    '                        'Now, load our textures
    '                        .Textures((X * 4)) = GetTexture(sTemp & sPostFix, eGetTextureType.ModelTexture)
    '                        .Textures((X * 4) + 1) = GetTexture(sTemp & "_mine" & sPostFix, eGetTextureType.ModelTexture)
    '                        .Textures((X * 4) + 2) = GetTexture(sTemp & "_ally" & sPostFix, eGetTextureType.ModelTexture)
    '                        .Textures((X * 4) + 3) = GetTexture(sTemp & "_enemy" & sPostFix, eGetTextureType.ModelTexture)

    '                    End If
    '                Next X

    '                'Now, check if there is a Turret mesh...
    '                sTemp = oINI.GetString(sModelHdr, "TurretFileName", "")
    '                If sTemp <> "" Then
    '                    sModel = msMeshPath & sTemp
    '                    If Dir$(sModel) <> "" Then
    '                        Dim oTurretMatBuff() As ExtendedMaterial
    '                        .oTurretMesh = Mesh.FromFile(sModel, MeshFlags.Managed, moDevice, oTurretMatBuff)
    '                        'Ensure lighting and normals and stuff
    '                        If (.oTurretMesh.VertexFormat And VertexFormats.Normal) = 0 Then
    '                            Dim oTmpMesh As Mesh = .oTurretMesh.Clone(.oTurretMesh.Options.Value, .oTurretMesh.VertexFormat Or VertexFormats.Normal, moDevice)
    '                            oTmpMesh.ComputeNormals()
    '                            .oTurretMesh.Dispose()
    '                            .oTurretMesh = oTmpMesh
    '                            oTmpMesh = Nothing
    '                        End If
    '                        'Do some optimization
    '                        adj = .oTurretMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
    '                        .oTurretMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
    '                        Erase adj

    '                        'Now, it is assumed that a turret mesh uses the same materials/textures as the parent
    '                        'So, we do not load them again... instead, simply get the remaining details
    '                        .lTurretZOffset = Val(oINI.GetString(sModelHdr, "TurretOffsetZ", "0"))
    '                        .bTurretMesh = True
    '                    End If
    '                End If
    '            End With
    '        Else
    '            'return a sphere... hehe
    '            moMeshes(lIndex).oMesh = Mesh.Sphere(moDevice, 32, 32, 32)
    '            moMeshes(lIndex).NumOfMaterials = 0
    '        End If
    '    Else
    '        'return a sphere... hehe
    '        moMeshes(lIndex).oMesh = Mesh.Sphere(moDevice, 32, 32, 32)
    '        moMeshes(lIndex).NumOfMaterials = 0
    '    End If

    '    oINI = Nothing
    'End Sub

	Private Sub LoadMesh(ByVal lIndex As Int32, ByVal lModelID As Int32)
        If Debugger.IsAttached = False OrElse Not Exists(msMeshPath & "models.dat") Then If UnpackLocalDataFile("md.pak", "Models.dat") = False Then Return

        'MSC: Only concerned with the mesh portion of the modelid
        lModelID = (lModelID And 255)

		Try

            Dim moDevice As Device = GFXEngine.moDevice

			Dim oINI As InitFile = New InitFile(msMeshPath & "models.dat")
			Dim sModel As String
			Dim mtrlBuffer() As ExtendedMaterial = Nothing
			Dim X As Int32
			Dim sTemp As String

			Dim adj() As Int32

			Dim sPostFix As String

			Dim sModelHdr As String = "MODEL_" & lModelID

			Dim oMem As IO.MemoryStream
			Dim sPakFile As String
			Dim sModelPak As String 

			'prepare our object...
			moMeshes(lIndex) = New BaseMesh()
			mlMeshIdx(lIndex) = lModelID

			sModel = oINI.GetString(sModelHdr, "FileName", "")
			sPakFile = oINI.GetString(sModelHdr, "TexturePak", "")
			sModelPak = oINI.GetString(sModelHdr, "ModelPak", "3DObjects.pak")


			If sModel <> "" Then
				'sModel = msMeshPath & sModel
				'now set it up
				With moMeshes(lIndex)

					Dim fRadius As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereSize", "0")))
					Dim fShiftX As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereShiftX", "0")))
					Dim fShiftY As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereShiftY", "0")))
					Dim fShiftZ As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereShiftZ", "0")))
					Dim fStretchX As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchX", "0")))
					Dim fStretchY As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchY", "0")))
					Dim fStretchZ As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchZ", "0")))


					'If we are loading the mesh, then we need to load the shield mesh too
					X = CInt(Val(oINI.GetString(sModelHdr, "ShieldSphereSize", "0")))
					If X <> 0 Then .oShieldMesh = CreateShieldSphere(fRadius, fShiftX, fShiftY, fShiftZ, fStretchX, fStretchY, fStretchZ)
					.ShieldXZRadius = CSng(Math.Max(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchX", "1")), Val(oINI.GetString(sModelHdr, "ShieldSphereStretchZ", "1"))))
					.ShieldXZRadius *= Math.Max(X, 1)

					'Now, our Y size values...
					Dim fTempHt As Single = 100.0F
					If X <> 0 Then
						'Dim fStretchY As Single = CSng(Val(oINI.GetString(sModelHdr, "ShieldSphereStretchY", "0")))

						fTempHt = CSng(Math.Ceiling(X * fStretchY))
					Else
						'ok, try our midpoint
						Dim fTempVal As Single = CSng(oINI.GetString(sModelHdr, "YMidPoint", "0"))
						If fTempVal < 1 Then
							'ok, use our death seq y
							fTempHt = CInt(oINI.GetString(sModelHdr, "DeathSeqSize_Y", "100"))
						Else
							fTempVal *= 2
							fTempHt = CSng(Math.Ceiling(fTempVal))
						End If
					End If
					fTempHt *= 2
					.YLevelEntityHeight = CInt(fTempHt + fShiftY)
					.YLevelRectSize = CInt(.ShieldXZRadius * 2)			 'factor of 2

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
					'.oMesh = .oMesh.FromFile(sModel, MeshFlags.Managed, moDevice, mtrlBuffer)
					.PlanetYAdjust = CInt(Val(oINI.GetString(sModelHdr, "PlanetYAdjust", "0")))
					.YMidPoint = CInt(Val(oINI.GetString(sModelHdr, "YMidPoint", "0")))
					.bLandBased = (Val(oINI.GetString(sModelHdr, "LandBased", "0")) <> 0)
					.lModelID = lModelID
                    .RangeOffset = CInt(Val(oINI.GetString(sModelHdr, "RangeOffset", "0")))

                    'Dim lDivisor As Int32 = 1
                    'If .bLandBased = False Then lDivisor = 2

					.EngineFXCnt = CInt(Val(oINI.GetString(sModelHdr, "EngineFXCnt", "0")))
					ReDim .EngineFXOffset(.EngineFXCnt)
					ReDim .EngineFXPCnt(.EngineFXCnt)
					For X = 0 To .EngineFXCnt - 1
                        .EngineFXOffset(X).X = CInt(oINI.GetString(sModelHdr, "Engine" & X & "_OffsetX", "0")) '\ lDivisor
                        .EngineFXOffset(X).Y = CInt(oINI.GetString(sModelHdr, "Engine" & X & "_OffsetY", "0")) '\ lDivisor
                        .EngineFXOffset(X).Z = CInt(oINI.GetString(sModelHdr, "Engine" & X & "_OffsetZ", "0")) '\ lDivisor
						.EngineFXPCnt(X) = CInt(oINI.GetString(sModelHdr, "Engine" & X & "_PCnt", "0"))
					Next X

					'Load the Fire From Locs
					Dim lSide As Int32
					Dim lMax As Int32
					For lSide = 0 To 3
						lMax = CInt(Val(oINI.GetString(sModelHdr, "FireFromCnt_" & lSide, "0"))) - 1
						ReDim .muFireLocs(lSide)(lMax)
						For X = 0 To lMax
                            .muFireLocs(lSide)(X).lOffsetX = CInt(Val(oINI.GetString(sModelHdr, "FireFromLoc_" & lSide & "_" & X & "_X", "0"))) '\ lDivisor
                            .muFireLocs(lSide)(X).lOffsetY = CInt(Val(oINI.GetString(sModelHdr, "FireFromLoc_" & lSide & "_" & X & "_Y", "0"))) '\ lDivisor
                            .muFireLocs(lSide)(X).lOffsetZ = CInt(Val(oINI.GetString(sModelHdr, "FireFromLoc_" & lSide & "_" & X & "_Z", "0"))) '\ lDivisor
							.muFireLocs(lSide)(X).ySideID = CByte(lSide)
						Next X
                    Next lSide

                    'Load the Light effects
                    Dim lLight As Int32
                    Dim lLights As Int32 = CInt(Val(oINI.GetString(sModelHdr, "Lights", "0"))) - 1
                    ReDim .Lights(lLights)
                    For lLight = 0 To lLights
                        .bHasLights = True
                        With .Lights(lLight)
                            .iLightID = -1
                            .lLightType = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_type", "0")))
                            Select Case .lLightType
                                Case 1
                                    .LightTex = TextureLoader.FromFile(moDevice, msAppPath & "textures\Spotlight.dds")
                            End Select
                            Dim A As Int32 = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_A", "0")))
                            Dim R As Int32 = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_R", "0")))
                            Dim G As Int32 = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_G", "0")))
                            Dim B As Int32 = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_B", "0")))
                            .Color = System.Drawing.Color.FromArgb(A, R, G, B)

                            Dim locX As Int32 = 0
                            Dim locY As Int32 = 0
                            Dim locZ As Int32 = 0
                            locX = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_S_X", "0")))
                            locY = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_S_Y", "0")))
                            locZ = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_S_Z", "0")))
                            .Source = New Vector3(locX, locY, locZ)

                            locX = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_E_X", "0")))
                            locY = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_E_Y", "0")))
                            locZ = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_E_Z", "0")))
                            .Direction = New Vector3(locX, locY, locZ)

                            .bSearchLight = CBool(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "SearchLight", "0")) <> 0)
                            If .bSearchLight Then
                                .SearchLightCycles = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_Cycles", "0")))
                                .SearchLightFormula = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_Formula", "0")))

                                locX = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_SL_SX", "0")))
                                locY = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_SL_SY", "0")))
                                locZ = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_SL_SZ", "0")))
                                .SearchLightStart = New Vector3(locX, locY, locZ)

                                locX = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_SL_EX", "0")))
                                locY = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_SL_EY", "0")))
                                locZ = CInt(Val(oINI.GetString(sModelHdr, "Light_" & lLight & "_SL_EZ", "0")))
                                .SearchLightEnd = New Vector3(locX, locY, locZ)
                            End If
                        End With
                    Next lLight

                    If oINI.GetString(sModelHdr, "mttexCapable", "0") = "1" Then
                        'Ok, we are mttexcapable, so lets load up our textures for that...
                        .bMTTexCapable = True
                        Dim lTmpUB As Int32 = -1
                        For X = 0 To 100
                            Dim sTex As String = oINI.GetString(sModelHdr, "mt" & X & "_tex", "")
                            If sTex = "" Then Exit For

                            lTmpUB += 1
                            ReDim Preserve .oSharedTex(lTmpUB)
                            ReDim Preserve .lSharedTexOpt(lTmpUB)
                            ReDim Preserve .sSharedTexName(lTmpUB)
                            .oSharedTex(lTmpUB) = GetSharedTexture(sTex)
                            .sSharedTexName(lTmpUB) = sTex
                            .lSharedTexOpt(lTmpUB) = CInt(oINI.GetString(sModelHdr, "mt" & X & "_opt", "0"))
                        Next X
                        .lDefaultSharedTexIdx = CInt(oINI.GetString(sModelHdr, "startmttex", "0"))
                    Else
                        .bMTTexCapable = False
                    End If

                    'Load our mesh points for where weapon shots will hit...
                    'If .bLandBased = False Then .oMesh = ScaleMesh(.oMesh)
                    .LoadMeshPoints()

					'Load the Burn Sides...
					Dim lBurnFXCnt As Int32 = CInt(Val(oINI.GetString(sModelHdr, "BurnFXCnt", "0")))
					ReDim .muBurnLocs(lBurnFXCnt - 1)
					For X = 0 To lBurnFXCnt - 1
						.muBurnLocs(X).lSide = CInt(Val(oINI.GetString(sModelHdr, "BurnSide_" & X, "0")))
                        .muBurnLocs(X).lOffsetX = CInt(Val(oINI.GetString(sModelHdr, "Burn" & X & "_X", "0"))) ' \ lDivisor
                        .muBurnLocs(X).lOffsetY = CInt(Val(oINI.GetString(sModelHdr, "Burn" & X & "_Y", "0"))) '\ lDivisor
                        .muBurnLocs(X).lOffsetZ = CInt(Val(oINI.GetString(sModelHdr, "Burn" & X & "_Z", "0"))) '\ lDivisor
						.muBurnLocs(X).lPCnt = CInt(Val(oINI.GetString(sModelHdr, "Burn" & X & "_PCnt", "0")))
						.muBurnLocs(X).lEmitterType = CInt(Val(oINI.GetString(sModelHdr, "Burn" & X & "_Type", "0")))
					Next X

					'Load the DeathSequenceData...
                    .vecDeathSeqSize.X = CInt(Val(oINI.GetString(sModelHdr, "DeathSeqSize_X", CInt(.ShieldXZRadius).ToString))) '\ lDivisor
                    .vecDeathSeqSize.Y = CInt(Val(oINI.GetString(sModelHdr, "DeathSeqSize_Y", .YMidPoint.ToString))) '\ lDivisor
                    .vecDeathSeqSize.Z = CInt(Val(oINI.GetString(sModelHdr, "DeathSeqSize_Z", .vecDeathSeqSize.X.ToString))) '\ lDivisor

					'Set our half vector
					.vecHalfDeathSeqSize = New Vector3(.vecDeathSeqSize.X / 2.0F, .vecDeathSeqSize.Y / 2.0F, .vecDeathSeqSize.Z / 2.0F)

					.lDeathSeqExpCnt = CInt(Val(oINI.GetString(sModelHdr, "DeathSeqExpCnt", "1")))
					.bDeathSeqFinale = (Val(oINI.GetString(sModelHdr, "DeathSeqFinale", "0")) <> 0)
					X = CInt(Val(oINI.GetString(sModelHdr, "DeathSeqWAV_Cnt", "1")))
					ReDim .sDeathSeqWav(X - 1)
					For X = 0 To .sDeathSeqWav.GetUpperBound(0)
						.sDeathSeqWav(X) = oINI.GetString(sModelHdr, "DeathSeqWav_" & X, "")
					Next X

                    If lModelID = 37 OrElse lModelID = 4 OrElse lModelID = 3 OrElse lModelID = 90 OrElse lModelID = 95 OrElse lModelID = 36 Then
                        ' Get the original mesh's index buffer.
                        Dim ranks(0) As Int32
                        Dim arr As System.Array

                        ranks(0) = .oMesh.NumberFaces * 3
                        arr = .oMesh.LockIndexBuffer((New Short()).GetType(), LockFlags.None, ranks)
                        Array.Reverse(arr)
                        .oMesh.IndexBuffer.SetData(arr, 0, LockFlags.None)

                        arr = Nothing
                    End If


					'Now, because our engine uses Normals and lighting, let's get that out of the way
					If (.oMesh.VertexFormat And VertexFormats.Normal) = 0 Then
						Dim oTmpMesh As Mesh = .oMesh.Clone(.oMesh.Options.Value, .oMesh.VertexFormat Or VertexFormats.Normal, moDevice)
                        oTmpMesh.ComputeNormals()
						.oMesh.Dispose()
						.oMesh = oTmpMesh
						oTmpMesh = Nothing
					End If

                    .oMesh.ComputeNormals()
					If muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap Then
						MakeMeshNormalMapAble(.oMesh)
					End If


                    'Now, for the sound fx file
                    Select Case lModelID
                        Case 3, 66, 67, 68, 69, 71, 2, 45, 46, 70, 73, 74
                            .sRoarSFX = "Unit Sounds\SmallShipRoar2.wav"
                        Case 4, 36, 53, 55, 59, 64, 76, 85, 97
                            .sRoarSFX = "Unit Sounds\MediumShipRoar4.wav"
                        Case 81, 82, 87, 44, 61, 63, 88
                            .sRoarSFX = "Unit Sounds\LargeShipRoar2.wav"
                        Case 50, 43, 89, 98, 5
                            .sRoarSFX = "Unit Sounds\LargeShipRoar1.wav"
                        Case 18, 19, 90, 93, 99, 101, 92
                            .sRoarSFX = "Unit Sounds\MediumShipRoar3.wav"
                        Case 37, 47, 72, 96, 100
                            .sRoarSFX = "Unit Sounds\SmallShipRoar1.wav"
                        Case 52, 80, 84, 91, 58, 60, 62, 75, 77, 78, 83
                            .sRoarSFX = "Unit Sounds\MediumShipRoar2.wav"
                        Case 48, 49, 51, 54, 56, 57, 65, 79, 86, 94, 95
                            .sRoarSFX = "Unit Sounds\MediumShipRoar1.wav"
                        Case Else
                            .sRoarSFX = ""
                    End Select


					' Optimize the mesh for this graphics card's vertex cache 
					' so when rendering the mesh's triangle list the vertices will 
					' cache hit more often so it won't have to re-execute the vertex shader 
					' on those vertices so it will improve perf.     
					adj = .oMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
					.oMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
					Erase adj

					Device.IsUsingEventHandlers = True

					If muSettings.LightQuality = EngineSettings.LightQualitySetting.VSPS1 Then
                        'Determine if we are using the new model render method or the old
                        If GFXEngine.mbSupportsNewModelMethod = True Then
                            If sPakFile = "" Then sPakFile = "textures.pak"
                            Dim lTemp As Int32 = sPakFile.LastIndexOf("."c)
                            If lTemp < 1 Then
                                sPakFile = "textures.pak"
                                lTemp = sPakFile.LastIndexOf("."c)
                            End If
                            sPakFile = sPakFile.Substring(0, lTemp) & "2" & sPakFile.Substring(lTemp)

                            'Initialize our materials and textures using the new method
                            .NumOfMaterials = mtrlBuffer.Length
                            ReDim .Materials(.NumOfMaterials - 1)
                            ReDim .Textures(.NumOfMaterials - 1)
                            ReDim .sTextures(.NumOfMaterials - 1)
                            ReDim .sTexturePak(.NumOfMaterials - 1)
                            For X = 0 To mtrlBuffer.Length - 1
                                .Materials(X) = mtrlBuffer(X).Material3D
                                .Materials(X).Diffuse = Color.White
                                .Materials(X).Ambient = .Materials(X).Diffuse
                                If mtrlBuffer(X).TextureFilename <> "" Then
                                    sTemp = oINI.GetString(sModelHdr, "Texture1Name", "")
                                    If sTemp = "" Then sTemp = mtrlBuffer(X).TextureFilename

                                    Dim lDotIdx As Int32 = sTemp.LastIndexOf("."c)
                                    sPostFix = sTemp.Substring(lDotIdx + 1) ' Mid$(sTemp, sTemp.Length - 3, 4)
                                    sTemp = sTemp.Substring(0, lDotIdx) & ".dds" ' Mid$(sTemp, 1, sTemp.Length - 4) & ".dds"
                                    .Textures(X) = GetTexture(sTemp, eGetTextureType.ModelTexture, sPakFile)
                                    .sTextures(X) = sTemp
                                    .sTexturePak(X) = sPakFile
                                End If
                            Next X
                        Else
                            'now initialize our materials and textures using the old method
                            .NumOfMaterials = mtrlBuffer.Length
                            ReDim .Materials(.NumOfMaterials - 1)
                            ReDim .Textures((.NumOfMaterials * 4) - 1)          '*4, 1 for each relationship state
                            ReDim .sTextures((.NumOfMaterials * 4) - 1)
                            ReDim .sTexturePak((.NumOfMaterials * 4) - 1)
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
                                    .sTexturePak(X * 4) = sPakFile
                                    .Textures((X * 4) + 1) = GetTexture(sTemp & "_mine" & sPostFix, eGetTextureType.ModelTexture, sPakFile)
                                    .sTextures((X * 4) + 1) = sTemp & "_mine" & sPostFix
                                    .sTexturePak((X * 4) + 1) = sPakFile
                                    .Textures((X * 4) + 2) = GetTexture(sTemp & "_ally" & sPostFix, eGetTextureType.ModelTexture, sPakFile)
                                    .sTextures((X * 4) + 2) = sTemp & "_ally" & sPostFix
                                    .sTexturePak((X * 4) + 2) = sPakFile
                                    .Textures((X * 4) + 3) = GetTexture(sTemp & "_enemy" & sPostFix, eGetTextureType.ModelTexture, sPakFile)
                                    .sTextures((X * 4) + 3) = sTemp & "_enemy" & sPostFix
                                    .sTexturePak((X * 4) + 3) = sPakFile
                                End If
                            Next X
                        End If

					Else
						'New shader level... which includes PerPixel shading, Illumination Maps and Normal Maps...
						If sPakFile = "" Then sPakFile = "textures.pak"
						Dim lTemp As Int32 = sPakFile.LastIndexOf("."c)
						If lTemp < 1 Then
							sPakFile = "textures.pak"
							lTemp = sPakFile.LastIndexOf("."c)
						End If
						sPakFile = sPakFile.Substring(0, lTemp) & "din" & sPakFile.Substring(lTemp)

						Dim sTexName As String = oINI.GetString(sModelHdr, "TextureDef", "")

						'Initialize our materials and textures using the new method
						.NumOfMaterials = mtrlBuffer.Length
						ReDim .Materials(.NumOfMaterials - 1)
						ReDim .Textures(.NumOfMaterials - 1)
						ReDim .sTextures(.NumOfMaterials - 1)
						ReDim .sTexturePak(.NumOfMaterials - 1)
						For X = 0 To mtrlBuffer.Length - 1
							.Materials(X) = mtrlBuffer(X).Material3D
							.Materials(X).Diffuse = Color.White
							.Materials(X).Ambient = .Materials(X).Diffuse
							If mtrlBuffer(X).TextureFilename <> "" Then
								If sTexName <> "" Then
									sTemp = sTexName & ".dds"
								Else
									sTemp = mtrlBuffer(X).TextureFilename
									sPostFix = Mid$(sTemp, sTemp.Length - 3, 4)
									sTemp = Mid$(sTemp, 1, sTemp.Length - 4) & ".dds"
								End If
								.Textures(X) = GetTexture(sTemp, eGetTextureType.ModelTexture, sPakFile)
								.sTextures(X) = sTemp
								.sTexturePak(X) = sPakFile
							End If
						Next X

						If sTexName <> "" Then
							If muSettings.IlluminationMap <> -1 Then
								sTemp = sTexName & "_i.dds"
								.sIllumMap = sTemp
								.sIllumMapPak = sPakFile
								.oIllumMap = GetTexture(sTemp, eGetTextureType.IlluminationMap, sPakFile)
							End If
							If muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap Then
								sTemp = sTexName & "_n.dds"
								.sNormalMap = sTemp
								.sNormalMapPak = sPakFile
								.oNormalMap = GetTexture(sTemp, eGetTextureType.ModelTexture, sPakFile)
							End If
						End If

					End If

					'Now, check if there is a Turret mesh...
					sTemp = oINI.GetString(sModelHdr, "TurretFileName", "")
					If sTemp <> "" Then
						sModel = sTemp

						Dim oTurretMatBuff() As ExtendedMaterial = Nothing

						oMem = GetResourceStream(sModel, msMeshPath & "3DObjects.pak")
						If oMem Is Nothing = False Then
							'.oTurretMesh = Mesh.FromFile(sModel, MeshFlags.Managed, moDevice, oTurretMatBuff)

							.oTurretMesh = Mesh.FromStream(oMem, MeshFlags.Managed, moDevice, oTurretMatBuff)

							'Release our memory buffer
							oMem.Close()
							oMem = Nothing

							'Ensure lighting and normals and stuff
							If (.oTurretMesh.VertexFormat And VertexFormats.Normal) = 0 Then
								Dim oTmpMesh As Mesh = .oTurretMesh.Clone(.oTurretMesh.Options.Value, .oTurretMesh.VertexFormat Or VertexFormats.Normal, moDevice)
								oTmpMesh.ComputeNormals()
								.oTurretMesh.Dispose()
								.oTurretMesh = oTmpMesh
								oTmpMesh = Nothing
                            End If

                            .oTurretMesh.ComputeNormals()
                            If muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap Then
                                MakeMeshNormalMapAble(.oTurretMesh)
                            End If

							'Do some optimization
							adj = .oTurretMesh.ConvertPointRepsToAdjacency(CType(Nothing, GraphicsStream))
							.oTurretMesh.OptimizeInPlace(MeshFlags.OptimizeVertexCache, adj)
							Erase adj

							'Now, it is assumed that a turret mesh uses the same materials/textures as the parent
							'So, we do not load them again... instead, simply get the remaining details
							.lTurretZOffset = CInt(Val(oINI.GetString(sModelHdr, "TurretOffsetZ", "0")))
							.bTurretMesh = True
						End If
					End If
				End With
			Else
				'return a sphere... hehe
				moMeshes(lIndex).oMesh = Mesh.Sphere(moDevice, 32, 32, 32)
				moMeshes(lIndex).NumOfMaterials = 0
			End If

			oINI = Nothing
		Catch
			'TODO: could trap this error
        Finally
            Try
                If Debugger.IsAttached = False Then If Exists(msMeshPath & "models.dat") = True Then Kill(msMeshPath & "models.dat")
            Catch
            End Try
        End Try

    End Sub

	'Private Sub InvertNormals(ByRef oMesh As Mesh)
	'	Dim elems() As VertexElement = oMesh.Declaration

	'	If elems Is Nothing = False Then

	'		'Get our offset for the normal
	'		Dim lOffset As Int32
	'		For X As Int32 = 0 To elems.GetUpperBound(0)
	'			If elems(X).DeclarationUsage = DeclarationUsage.Normal Then
	'				lOffset = elems(X).Offset
	'				Exit For
	'			End If
	'		Next X

	'		' Get the original mesh's vertex buffer.
	'		Dim ranks(0) As Integer
	'		ranks(0) = oMesh.NumberVertices

	'		Dim arr As System.Array = oMesh.VertexBuffer.Lock(0, New Byte().GetType, LockFlags.None, ranks)

	'		'Now, go through and get our stuff...
	'		For X As Int32 = 0 To arr.Length - oMesh.NumberBytesPerVertex + 1 Step oMesh.NumberBytesPerVertex
	'			'Ok, now, get our index
	'			Dim lIdx As Int32 = X + lOffset

	'			Dim yTemp(11) As Byte
	'			For Y As Int32 = 0 To 11
	'				yTemp(Y) = CByte(arr.GetValue(lIdx + Y))
	'			Next Y
	'			Dim fX As Single = System.BitConverter.ToSingle(yTemp, 0)
	'			Dim fY As Single = System.BitConverter.ToSingle(yTemp, 4)
	'			Dim fZ As Single = System.BitConverter.ToSingle(yTemp, 8)
	'			Dim vecTemp As New Vector3(fX, -fY, fZ)
	'			vecTemp.Normalize()

	'			System.BitConverter.GetBytes(vecTemp.X).CopyTo(yTemp, 0)
	'			System.BitConverter.GetBytes(vecTemp.Y).CopyTo(yTemp, 4)
	'			System.BitConverter.GetBytes(vecTemp.Z).CopyTo(yTemp, 8)
	'			For Y As Int32 = 0 To 11
	'				arr.SetValue(yTemp(Y), lIdx + Y)
	'			Next Y
	'		Next X

	'		oMesh.VertexBuffer.Unlock()
	'	End If

	'End Sub

	Private Sub MakeMeshNormalMapAble(ByRef oMesh As Mesh)
		Dim elems(5) As VertexElement
		'position, normal, texture coords, tangent, binormal.

		elems(0) = New VertexElement(0, 0, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Position, 0)
		elems(1) = New VertexElement(0, 12, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Normal, 0)
		elems(2) = New VertexElement(0, 24, DeclarationType.Float2, DeclarationMethod.Default, DeclarationUsage.TextureCoordinate, 0)
		elems(3) = New VertexElement(0, 32, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.Tangent, 0)
		elems(4) = New VertexElement(0, 44, DeclarationType.Float3, DeclarationMethod.Default, DeclarationUsage.BiNormal, 0)
		elems(5) = VertexElement.VertexDeclarationEnd

        Dim clonedmesh As Mesh = oMesh.Clone(MeshFlags.Managed, elems, GFXEngine.moDevice)
		oMesh.Dispose()
		oMesh = clonedmesh
		clonedmesh = Nothing

		oMesh = Geometry.ComputeTangentFrame(oMesh, DeclarationUsage.TextureCoordinate, 0, DeclarationUsage.Tangent, 0, DeclarationUsage.BiNormal, 0, _
		DeclarationUsage.Normal, 0, 0, Nothing, 0.0F, 0.0F, 0.0F, Nothing)

	End Sub

    Public Sub DeleteMesh(ByVal lModelID As Int32)
        'MSC: we only care about the mesh portion of the model
        lModelID = (lModelID And 255)
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
        Dim moDevice As Device = GFXEngine.moDevice

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
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, 0, 0, 1, Usage.None, Format.Unknown, Pool.Managed, Filter.None, Filter.None, System.Drawing.Color.FromArgb(255, 255, 0, 255).ToArgb)
                Case eGetTextureType.StartupTexture
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, 0, 0, 1, Usage.None, Format.Unknown, Pool.Managed, Filter.Triangle, Filter.Triangle, 0)
				Case eGetTextureType.IlluminationMap
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, muSettings.IlluminationMap, muSettings.IlluminationMap, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Box, 0)
				Case eGetTextureType.WaterTexture
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, muSettings.WaterTextureRes, muSettings.WaterTextureRes, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Box, 0)
				Case Else		'includes NoSpecifics
                    moTextures(lIndex) = TextureLoader.FromStream(moDevice, oMem, 0, 0, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Box, Filter.Box, 0)
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

    'Private Function ScaleMesh(ByRef oMesh As Mesh) As Mesh
    '    'Ok, we need to lock the mesh... let's see what the declaration of the mesh is...
    '    Dim decl() As VertexElement = oMesh.Declaration
    '    Dim lPosOffset As Int32 = 0
    '    Dim lNormalOffset As Int32 = 0
    '    Dim lTotal As Int32 = 0
    '    If decl Is Nothing = False Then
    '        For X As Int32 = 0 To decl.GetUpperBound(0)
    '            If Object.Equals(decl(X), VertexElement.VertexDeclarationEnd) Then

    '                If X <> 0 Then
    '                    lTotal = decl(X - 1).Offset
    '                    Select Case decl(X - 1).DeclarationType
    '                        Case DeclarationType.Float2
    '                            lTotal += 8
    '                        Case DeclarationType.Float3
    '                            lTotal += 12
    '                        Case DeclarationType.Float4
    '                            lTotal += 16
    '                    End Select
    '                End If

    '                Exit For
    '            End If
    '            If decl(X).DeclarationUsage = DeclarationUsage.Position Then
    '                lPosOffset = decl(X).Offset
    '            ElseIf decl(X).DeclarationUsage = DeclarationUsage.Normal Then
    '                lNormalOffset = decl(X).Offset
    '            End If
    '        Next X
    '    End If

    '    Dim yData() As Byte = Nothing
    '    Dim lTotalElems As Int32
    '    Try
    '        Dim ranks(0) As Int32
    '        ranks(0) = oMesh.NumberVertices * lTotal
    '        yData = CType(oMesh.VertexBuffer.Lock(0, New Byte().GetType(), LockFlags.None, ranks), Byte())
    '        lTotalElems = yData.Length \ lTotal
    '        If lTotalElems <> oMesh.NumberVertices Then lTotalElems = Math.Min(lTotalElems, oMesh.NumberVertices)
    '    Catch
    '        Return oMesh
    '    End Try

    '    'For storing which index in the array we are on...
    '    Dim fScaleFactor As Single = 0.5F
    '    For i As Int32 = 0 To lTotalElems - 1
    '        Try
    '            Dim lVertexPos As Int32 = i * lTotal
    '            Dim lPos As Int32 = lVertexPos + lPosOffset
    '            Dim lBasePos As Int32 = lPos
    '            Dim fLocX As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
    '            Dim fLocY As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
    '            Dim fLocZ As Single = System.BitConverter.ToSingle(yData, lPos) : lPos += 4

    '            fLocX *= fScaleFactor
    '            fLocY *= fScaleFactor
    '            fLocZ *= fScaleFactor
    '            System.BitConverter.GetBytes(fLocX).CopyTo(yData, lBasePos) : lBasePos += 4
    '            System.BitConverter.GetBytes(fLocY).CopyTo(yData, lBasePos) : lBasePos += 4
    '            System.BitConverter.GetBytes(fLocZ).CopyTo(yData, lBasePos) : lBasePos += 4

    '        Catch
    '        End Try
    '    Next i

    '    oMesh.VertexBuffer.Unlock()

    '    Return oMesh
    'End Function

    'TODO: Added the bOffset flag to this function as an optional. We want to remove this once the models all default to 0
    Public Function CreateTexturedSphere(ByVal fRadius As Single, ByVal lSlices As Int32, ByVal lStacks As Int32, ByVal lWrapCount As Int32, Optional ByVal bOffsetY As Boolean = False) As Mesh

        Device.IsUsingEventHandlers = False

        Dim moDevice As Device = GFXEngine.moDevice

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

        'ReDim adj(3 * TexturedObject.NumberFaces)

        ''Now, generate our adjacency
        'TexturedObject.GenerateAdjacency(0, adj)
        ''Now, go back and set our materials
        'Dim mtrlBuffer(0) As ExtendedMaterial
        'Dim mat As Material
        'With mat
        '    .Ambient = Color.White
        '    .Diffuse = Color.White
        'End With
        'mtrlBuffer(0).Material3D = mat
        ''save it
        'TexturedObject.Save("C:\Test.x", adj, mtrlBuffer, XFileFormat.Text)

        Return TexturedObject

    End Function

	Public Function CreateTexturedBox(ByVal fWidth As Single, ByVal fHeight As Single, ByVal fDepth As Single) As Mesh

		Device.IsUsingEventHandlers = False
        Dim moDevice As Device = GFXEngine.moDevice
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

        Dim oTmp As Mesh = New Mesh(lVerts * 3, lVerts, MeshFlags.Managed, CustomVertex.PositionNormalTextured.Format, GFXEngine.moDevice)
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
        Dim moDevice As Device = GFXEngine.moDevice
        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap Then
            Dim oMem As IO.MemoryStream = GetResourceStream(sTextureName, msTexturePath & "VolTerr.pak")
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
	Public Function GetVolNormalTexture(ByVal sTextureName As String) As VolumeTexture
		'NOTE: This particular function is a SCRATCH surface, meaning, when it is gone, it is gone.
        Dim moDevice As Device = GFXEngine.moDevice
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
        Dim moDevice As Device = GFXEngine.moDevice
        If moDevice.DeviceCaps.TextureCaps.SupportsVolumeMap = False Then
            Dim oMem As IO.MemoryStream = GetResourceStream(sTextureName, msTexturePath & "NVTerr.pak")

            Device.IsUsingEventHandlers = False
            'Dim oTmp As Texture = TextureLoader.FromFile(moDevice, msTexturePath & sTextureName)
            Dim oTmp As Texture = TextureLoader.FromStream(moDevice, oMem, 0, 0, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Linear, Filter.Linear, 0)
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

        Dim moDevice As Device = GFXEngine.moDevice

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
	Private Function CreateShieldSphere(ByVal fRadius As Single, ByVal fShiftX As Single, ByVal fShiftY As Single, ByVal fShiftZ As Single, ByVal fStretchX As Single, ByVal fStretchY As Single, ByVal fStretchZ As Single) As Mesh
		'Dim phi As Single
		'Dim u As Single
		Dim i As Integer

		Device.IsUsingEventHandlers = False

        Dim moDevice As Device = GFXEngine.moDevice

		Dim oTemp As Mesh = Mesh.Sphere(moDevice, fRadius, 32, 32)		'TODO: we could make an improvement here by possibly reducing stacks/slices to 16
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
            Array.Copy(yData, 0, yEntry, 0, lEntryLen - 1)

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
		Dim yFinal(lLen - 1) As Byte
        Dim lChrCode As Int32
        Dim lMod As Int32

        'Get our key value...
        lKey = yBytes(0)

        VBMath.Rnd(-1)
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
        ''release' whatever that means the default pool textures
        For X As Int32 = 0 To TextureUB
            Try
                If moTextures(X) Is Nothing = False AndAlso moTextures(X).Disposed = False AndAlso moTextures(X).GetLevelDescription(0).Pool = Pool.Default Then
                    moTextures(X).Dispose()
                    moTextures(X) = Nothing
                    msTextureNames(X) = ""
                    msTexPak(X) = ""
                End If
            Catch
            End Try
        Next X
    End Sub

    Public Sub RecreateDefaultPoolTextures()
        'TODO: 'recreate' the default pool textures
    End Sub

    Public Sub UpdateModelTextureResolution()
        LoadAllSharedTextures()
        For X As Int32 = 0 To TextureUB
            If msTextureNames(X) <> "" AndAlso mlTextureType(X) = eGetTextureType.ModelTexture Then

                Dim sTexture As String = msTextureNames(X)
                Dim sPak As String = msTexPak(X)

                Try
                    If moTextures(X) Is Nothing = False Then moTextures(X).Dispose()
                Catch
                End Try

                moTextures(X) = Nothing
                LoadTexture(X, sTexture, eGetTextureType.ModelTexture, sPak)
            End If
        Next X

        For X As Int32 = 0 To MeshUB
            If mlMeshIdx(X) <> -1 AndAlso moMeshes(X) Is Nothing = False AndAlso moMeshes(X).Textures Is Nothing = False Then
                For Y As Int32 = 0 To moMeshes(X).Textures.GetUpperBound(0)
                    '*SHOULD* return everything
                    Dim sPak As String = ""
                    If moMeshes(X).sTexturePak Is Nothing = False AndAlso moMeshes(X).sTexturePak.GetUpperBound(0) >= Y Then
                        sPak = moMeshes(X).sTexturePak(Y)
                    End If
                    moMeshes(X).Textures(Y) = Me.GetTexture(moMeshes(X).sTextures(Y), eGetTextureType.ModelTexture, sPak)

                    If muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap Then
                        If moMeshes(X).sNormalMap Is Nothing = False AndAlso moMeshes(X).sNormalMap <> "" Then
                            moMeshes(X).oNormalMap = GetTexture(moMeshes(X).sNormalMap, eGetTextureType.ModelTexture, moMeshes(X).sNormalMapPak)
                        End If
                    End If
                Next Y

                'RELOAD SHARED TEXS
                If moMeshes(X).oSharedTex Is Nothing = False Then
                    For Y As Int32 = 0 To moMeshes(X).oSharedTex.GetUpperBound(0)
                        moMeshes(X).oSharedTex(Y) = GetSharedTexture(moMeshes(X).sSharedTexName(Y))
                    Next Y
                End If
            End If
        Next X
    End Sub

    Public Sub UpdateIllumTextureResolution()
        For X As Int32 = 0 To TextureUB
            If msTextureNames(X) <> "" AndAlso mlTextureType(X) = eGetTextureType.IlluminationMap Then

                Dim sTexture As String = msTextureNames(X)
                Dim sPak As String = msTexPak(X)

                If moTextures(X) Is Nothing = False Then moTextures(X).Dispose()
                moTextures(X) = Nothing
                LoadTexture(X, sTexture, eGetTextureType.IlluminationMap, sPak)
            End If
        Next X

        For X As Int32 = 0 To mlSharedTexUB
            If moSharedTex(X) Is Nothing = False AndAlso moSharedTex(X).sIllumTexture Is Nothing = False Then
                For Y As Int32 = 0 To moSharedTex(X).sIllumTexture.GetUpperBound(0)
                    If moSharedTex(X).oIllumTexture(Y) Is Nothing = False Then
                        moSharedTex(X).oNormalTexture(Y).Dispose()
                        moSharedTex(X).oNormalTexture(Y) = Nothing
                    End If
                    If moSharedTex(X).sIllumTexture(Y) <> "" Then
                        moSharedTex(X).oIllumTexture(Y) = GetTexture(moSharedTex(X).sIllumTexture(Y), eGetTextureType.IlluminationMap, moSharedTex(X).sIllumPak(Y))
                    End If
                Next Y
            End If
        Next X

        For X As Int32 = 0 To MeshUB
            If mlMeshIdx(X) <> -1 AndAlso moMeshes(X) Is Nothing = False Then
                If moMeshes(X).oIllumMap Is Nothing = False Then
                    moMeshes(X).oIllumMap.Dispose()
                End If
                moMeshes(X).oIllumMap = Me.GetTexture(moMeshes(X).sIllumMap, eGetTextureType.IlluminationMap, moMeshes(X).sIllumMapPak)
            End If
        Next X
    End Sub

	Public Sub UpdateWaterTextureResolution()
		For X As Int32 = 0 To TextureUB
			If msTextureNames(X) <> "" AndAlso mlTextureType(X) = eGetTextureType.WaterTexture Then

				Dim sTexture As String = msTextureNames(X)
				Dim sPak As String = msTexPak(X)

				If moTextures(X) Is Nothing = False Then moTextures(X).Dispose()
				moTextures(X) = Nothing
				LoadTexture(X, sTexture, eGetTextureType.WaterTexture, sPak)
			End If
		Next X
		OceanRender.bRefreshTexture = True
	End Sub

    Public Function LoadScratchMeshNoTextures(ByVal sMeshName As String, ByVal sMeshPak As String) As Mesh
        Dim mtrlBuffer() As ExtendedMaterial = Nothing
        Dim adj() As Int32

        Dim oMem As IO.MemoryStream
        Dim oMesh As Mesh

        Dim moDevice As Device = GFXEngine.moDevice

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

    Public Function LoadScratchTexture(ByVal sTextureName As String, ByVal sPakFile As String) As Texture
        'NOTE: This particular function is a SCRATCH surface, meaning, when it is gone, it is gone.

        Dim oMem As IO.MemoryStream = GetResourceStream(sTextureName, msTexturePath & sPakFile)
		If oMem Is Nothing Then Return Nothing
        Device.IsUsingEventHandlers = False
        'Dim oTmp As Texture = TextureLoader.FromStream(moDevice, oMem)
        Dim oTmp As Texture = TextureLoader.FromStream(GFXEngine.moDevice, oMem, 0, 0, 0, Usage.None, Format.Unknown, Pool.Managed, Filter.Linear, Filter.Linear, 0)
        Device.IsUsingEventHandlers = True

        oMem.Close()
        oMem = Nothing

        Return oTmp  
	End Function

	Public Function UnpackLocalDataFile(ByVal sPakFile As String, ByVal sDATFile As String) As Boolean
		Dim oMem As IO.MemoryStream = GetResourceStream(sDATFile, msMeshPath & sPakFile)
		If oMem Is Nothing = False Then
			Try
				Dim oFS As New IO.FileStream(msMeshPath & sDATFile, IO.FileMode.Create)
				oFS.Write(oMem.GetBuffer, 0, CInt(oMem.Length - 2))
				oFS.Close()
				oFS.Dispose()
				oFS = Nothing
			Catch
			End Try
			Return True
		Else : Return False
		End If
	End Function

	Public Function UnpackAndGetDataFile(ByVal sPakFile As String, ByVal sDATFile As String) As IO.Stream
		Dim oMem As IO.MemoryStream = GetResourceStream(sDATFile, msMeshPath & sPakFile)
		Return oMem
	End Function
End Class
