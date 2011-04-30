Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class SimpleShader
    Private moEffect As Effect = Nothing
    Private moDevice As Device

    Private moDiffuse As Texture
    Private moNormal As Texture
    Private moBump As Texture
    Private moGlowMap As Texture

	Private mhndlWorldInv As EffectHandle
	Private mhndlWVP As EffectHandle
	Private mhndlEyePos As EffectHandle 
	Private mhndlLightDir As EffectHandle


	Public Sub RenderMesh(ByRef oMesh As Mesh, ByVal lCameraAtX As Int32, ByVal lCameraAtY As Int32, ByVal lCameraAtZ As Int32, ByVal vecLight As Vector3)

		'uniform extern float4x4 gWorldInv : WorldInverse;
		'uniform extern float4x4 gWVP : WorldViewProjection;
		'uniform extern float3   gEyePosW : CameraPosition;
		vecLight.Normalize()
		moEffect.SetValue(mhndlLightDir, New Vector4(vecLight.X, vecLight.Y, vecLight.Z, 1.0F))

		Dim oMatWorldInvert As Matrix = moDevice.Transform.World
		oMatWorldInvert.Invert()
		moEffect.SetValue(mhndlWorldInv, oMatWorldInvert)
		Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection) ' Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
		moEffect.SetValue(mhndlWVP, matWVP)

		Dim vecEyePos As Vector4 = New Vector4(lCameraAtX, lCameraAtY, lCameraAtZ, 1.0F)
		'vecEyePos.Normalize()
		moEffect.SetValue(mhndlEyePos, vecEyePos)



		Dim lPasses As Int32 = moEffect.Begin(FX.None)
		For X As Int32 = 0 To lPasses - 1
			moEffect.BeginPass(X)
			oMesh.DrawSubset(0)
			moEffect.EndPass()
		Next X
		moEffect.End()


	End Sub
	Public Sub New(ByRef oDevice As Device)
		moDevice = oDevice
		Device.IsUsingEventHandlers = False
		moEffect = Effect.FromFile(oDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\ShaderLibrary\NormalMap.fx", Nothing, Nothing, ShaderFlags.None, Nothing)

		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\3DRT\Naval\naval-01.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\battleship1\battleship1-texture.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser3\Cruiser3-texture.jpg")
		moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser1\Cruiser1-texture.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Battleships\octopus copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Battleships\x-63 copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Fighters\wasp\WASP  uvmap copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Fighters\Mantis Fighter\Mantis mat.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\crocodile copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\StarBlaster map.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\scorponok copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\Rhino copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\HammerHead copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\Avenger copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Space Cruisers\Ortas copy.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser2\Cruiser2-texture.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\ModelFinal\sp_st.bmp")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser6\Cruiser6-texture.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser4\Cruiser4-texture.jpg")
		'moNormal = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\MeshFormatter\bin\sp_st_normal.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\3DRT\Naval\nav-bump_NRM.jpg")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\battleship1\battleship1-texture_NRM.jpg")
		moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser1\BumpNormal2.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Battleships\octopus_Bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Battleships\x-63_Bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Fighters\wasp\WASP_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Fighters\Mantis Fighter\mantis_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\Croc_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\StarBlaster_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\scorp_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\rhino_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\hammerhead_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\Avenger_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Space Cruisers\Ortos_bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser6\Cruiser6-bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser4\Cruiser4-bump.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\ModelFinal\sp_st_bump_nrm.bmp")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\battleship1-texture_NRM.dds")
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser3\Cruiser3-bump.bmp")

		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Good Articles\Tutorial - Normal Mapping\bin\Debug\rockwall_normal.dds")
		'moGlowMap = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\3DRT\Naval\nav-illum.jpg")
		'moGlowMap = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\ModelFinal\sp_st-illum.jpg")
		'moGlowMap = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Battleships\x63_illum.bmp")
		'moGlowMap = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Frigates\Avenger_illum.bmp")
		'moGlowMap = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\7Keys\FinalSubmission_061207\Space Cruisers\Ortas_illum.bmp")

		moEffect.SetValue("gTex", moDiffuse)
		moEffect.SetValue("gNormalMap", moBump)
		'moEffect.SetValue("I", moBump)
		moEffect.SetValue("gIllumMap", moGlowMap)

		mhndlLightDir = moEffect.GetParameter(Nothing, "lightDir")

		moEffect.Technique = "NormalMapTech"

		mhndlWorldInv = moEffect.GetParameter(Nothing, "gWorldInv")
		mhndlWVP = moEffect.GetParameter(Nothing, "gWVP")
		mhndlEyePos = moEffect.GetParameter(Nothing, "gEyePosW")

		moEffect.SetValue("shininess", 30.0F)
		Dim vecTemp As Vector4

		Dim vecLightInv As Vector3 = moDevice.Lights(0).Direction
		vecLightInv.Normalize()
		With vecLightInv
			vecTemp.X = .X : vecTemp.Y = .Y : vecTemp.Z = .Z : vecTemp.W = 1.0F
		End With
		moEffect.SetValue("lightDir", vecTemp)
		vecTemp.X = 1.0F : vecTemp.Y = 1.0F : vecTemp.Z = 1.0F : vecTemp.W = 1.0F
		moEffect.SetValue("lightColor", vecTemp)
		moEffect.SetValue("lightSpecular", vecTemp)
		moEffect.SetValue("materialDiffuse", vecTemp)
		moEffect.SetValue("materialSpecular", vecTemp)
		moEffect.SetValue("materialAmbient", vecTemp)
		'> = {0.19f, 0.19f, 0.19f, 1.0f};
		vecTemp.X = 0.19F : vecTemp.Y = 0.19F : vecTemp.Z = 0.19F : vecTemp.W = 1.0F
		moEffect.SetValue("lightAmbient", vecTemp)
		Device.IsUsingEventHandlers = True
	End Sub

	Protected Overrides Sub Finalize()
		'If moEffect Is Nothing = False Then moEffect.Dispose()
		moEffect = Nothing
		MyBase.Finalize()
	End Sub
End Class

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


    Public Sub ExecutePostProcess()

        Dim oOriginal As Surface = moDevice.GetBackBuffer(0, 0, BackBufferType.Mono)

        moDevice.StretchRectangle(oOriginal, New Rectangle(0, 0, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight), _
          moSceneMap.GetSurfaceLevel(0), New System.Drawing.Rectangle(0, 0, moDevice.PresentationParameters.BackBufferWidth, moDevice.PresentationParameters.BackBufferHeight), TextureFilter.None)

        'Turn off the Z buffer
        moDevice.RenderState.ZBufferEnable = False
        moDevice.RenderState.ZBufferWriteEnable = False
        'And alpha blending/light
        moDevice.RenderState.AlphaBlendEnable = False
        moDevice.RenderState.Lighting = False

        Dim fWS(1) As Single
        fWS(0) = moDevice.PresentationParameters.BackBufferWidth
        fWS(1) = moDevice.PresentationParameters.BackBufferHeight
        If mhdlWindowSize Is Nothing = False Then moEffect.SetValue(mhdlWindowSize, fWS)
        If mhdlSceneMap Is Nothing = False Then moEffect.SetValue(mhdlSceneMap, moSceneMap)
        If mhdlDownsample Is Nothing = False Then moEffect.SetValue(mhdlDownsample, moDownsampleMap)
        If mhdlBlurMap1 Is Nothing = False Then moEffect.SetValue(mhdlBlurMap1, moBlurMap1)
        If mhdlBlurMap2 Is Nothing = False Then moEffect.SetValue(mhdlBlurMap2, moBlurMap2)


        Dim lPasses As Int32 = moEffect.Begin(FX.None)
        For X As Int32 = 0 To lPasses - 1
            If X = 0 Then
                'downsample
                moDevice.SetRenderTarget(0, moDownsampleMap.GetSurfaceLevel(0))
            ElseIf X = 1 Then
                'blurmap1
                moDevice.SetRenderTarget(0, moBlurMap1.GetSurfaceLevel(0))
            ElseIf X = 2 Then
                'blurmap2
                moDevice.SetRenderTarget(0, moBlurMap2.GetSurfaceLevel(0))
            Else
                'reset render target
                moDevice.SetRenderTarget(0, oOriginal)
            End If

            moEffect.BeginPass(X)
            RenderQuad(moDevice)
            moEffect.EndPass()
        Next X
        moEffect.End()
        'Turn on the Z buffer
        moDevice.RenderState.ZBufferEnable = True
        moDevice.RenderState.ZBufferWriteEnable = True
        'And alpha blending/light
        moDevice.RenderState.AlphaBlendEnable = True
        moDevice.RenderState.Lighting = True
        oOriginal = Nothing
    End Sub

    Public Sub New(ByRef oDevice As Device)
        moDevice = oDevice

        moEffect = Effect.FromFile(oDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\TempCode\MeshFormatter\bin\PostScreenGlow.fx", Nothing, Nothing, ShaderFlags.None, Nothing)

        moSceneMap = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth, oDevice.PresentationParameters.BackBufferHeight, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)
        moDownsampleMap = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth \ 4, oDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)
        moBlurMap1 = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth \ 4, oDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)
        moBlurMap2 = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth \ 4, oDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)

        mhdlWindowSize = moEffect.GetParameter(Nothing, "windowSize")
        mhdlSceneMap = moEffect.GetParameter(Nothing, "sceneMap")
        mhdlDownsample = moEffect.GetParameter(Nothing, "downsampleMap")
        mhdlBlurMap1 = moEffect.GetParameter(Nothing, "blurMap1")
        mhdlBlurMap2 = moEffect.GetParameter(Nothing, "blurMap2")

        moEffect.Technique = "ScreenGlow"
    End Sub

    Private Shared mvbQuad As VertexBuffer = Nothing
    Private Shared Sub RenderQuad(ByRef oDevice As Device)
        If mvbQuad Is Nothing Then
            If InitVBQuad(oDevice) = False Then Return
        End If
        oDevice.VertexFormat = CustomVertex.PositionTextured.Format
        oDevice.SetStreamSource(0, mvbQuad, 0, CustomVertex.PositionTextured.StrideSize)
        oDevice.DrawPrimitives(PrimitiveType.TriangleStrip, 0, 2)
    End Sub
    Private Shared Function InitVBQuad(ByRef oDevice As Device) As Boolean
        Dim uVert As CustomVertex.PositionTextured
        If mvbQuad Is Nothing Then mvbQuad = New VertexBuffer(uVert.GetType, 4, oDevice, Usage.WriteOnly, CustomVertex.PositionTextured.Format, Pool.Managed)
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

End Class