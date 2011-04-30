Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class SimpleShader
    Private moEffect As Effect = Nothing
    Private moDevice As Device

    Private moDiffuse As Texture
    Private moBump As Texture
    Private moGlowMap As Texture
    Private moSpecMap As Texture

	Private mhndlWorldInv As EffectHandle
	Private mhndlWVP As EffectHandle
	Private mhndlEyePos As EffectHandle 
	Private mhndlLightDir As EffectHandle
	Private mhndlSpecPwr As EffectHandle
	Private mhndlClrRel As EffectHandle
	Private mhndlWorld As EffectHandle
	Private mhndlWorldInvTrans As EffectHandle
	Public Sub SetDiffuseColor(ByVal clrValue As System.Drawing.Color)
		Dim vec4 As Vector4 = New Vector4(clrValue.R / 255.0F, clrValue.G / 255.0F, clrValue.B / 255.0F, 1.0F)
		'moEffect.SetValue("materialDiffuse", vec4)
	End Sub
	Public Sub SetSpecularColor(ByVal clrValue As System.Drawing.Color)
		Dim vec4 As Vector4 = New Vector4(clrValue.R / 255.0F, clrValue.G / 255.0F, clrValue.B / 255.0F, 1.0F)
		'moEffect.SetValue("materialSpecular", vec4)
	End Sub

	Public Sub SetDiffuseTexture(ByRef oTex As Texture)
		moDiffuse = oTex
		moEffect.SetValue("gTex", moDiffuse)
	End Sub
	Public Sub SetBumpTexture(ByRef oTex As Texture)
		moBump = oTex
		moEffect.SetValue("gNormalMap", moBump)
	End Sub
	Public Sub SetIllumTexture(ByRef oTex As Texture)
		moGlowMap = oTex
		moEffect.SetValue("gIllumMap", moGlowMap)
    End Sub
    Public Sub SetSpecTexture(ByRef oTex As Texture)
        moSpecMap = oTex
        moEffect.SetValue("gSpecMap", moSpecMap)
    End Sub
	Public Sub ClearTextures()
		moDiffuse = Nothing
		moBump = Nothing
        moGlowMap = Nothing
        moSpecMap = Nothing
		moEffect.SetValue("gTex", moDiffuse)
		moEffect.SetValue("gNormalMap", moBump)
        moEffect.SetValue("gIllumMap", moGlowMap)
        moEffect.SetValue("gSpecMap", moSpecMap)
	End Sub
	Public Sub SetSpecPower(ByVal fPower As Single)
		moEffect.SetValue(mhndlSpecPwr, fPower)
	End Sub

	Public sTechnique As String = "NormalMapFull"

	Private mlPasses As Int32 = -1
	Public Sub PrepareToRender()
		moEffect.Technique = sTechnique
		mlPasses = moEffect.Begin(FX.None)
	End Sub
	Public Sub EndRender()
		moEffect.End()
	End Sub

	Public Sub RenderMesh(ByRef oMesh As Mesh, ByVal lCameraAtX As Int32, ByVal lCameraAtY As Int32, ByVal lCameraAtZ As Int32, ByVal vecLight As Vector3, ByVal clrRel As System.Drawing.Color)

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

		Dim vecEyePos As Vector4 = New Vector4(lCameraAtX, lCameraAtY, lCameraAtZ, 0.0F)
		'vecEyePos.Normalize()
		moEffect.SetValue(mhndlEyePos, vecEyePos)

		moEffect.SetValue(mhndlClrRel, Vector4.Normalize((New Vector4(clrRel.R, clrRel.G, clrRel.B, 1.0F))))

		moEffect.SetValue(mhndlWorld, moDevice.Transform.World)
		Dim matWorldInvTrans As Matrix = moDevice.Transform.World
		matWorldInvTrans.Invert()
		matWorldInvTrans = Matrix.TransposeMatrix(matWorldInvTrans)
		moEffect.SetValue(mhndlWorldInvTrans, matWorldInvTrans)


		moEffect.SetValue(mhndlSpecPwr, 30)
		moEffect.SetValue("lightColor", New Vector4(1, 1, 1, 1))
		moEffect.SetValue("lightSpecular", New Vector4(1, 1, 1, 1))
        moEffect.SetValue("lightAmbient", New Vector4(0.19F, 0.19F, 0.19F, 0.19F))
		'moEffect.SetValue("RelColor", New Vector4(0, 1, 0, 0))
		moEffect.SetValue("EntityColorAdjust", New Vector4(1, 1, 1, 1))



		moEffect.CommitChanges()


		'Dim matWVP As Matrix
		'With matWVP
		'	.M11 = 1.36354589 : .M12 = -0.06492641 : .M13 = -0.291395575 : .M14 = -0.291395575
		'	.M21 = 0 : .M22 = 2.40399647 : .M23 = -0.09190115 : .M24 = -0.09190115
		'	.M31 = 0.4172868 : .M32 = 0.212156564 : .M33 = 0.952177763 : .M34 = 0.952177763
		'	.M41 = -14.9726686 : .M42 = -10.0906572 : .M43 = 378 : .M44 = 379
		'End With
		'moEffect.SetValue(mhndlWVP, matWVP)

		'Dim matWorldInverse As Matrix
		'With matWorldInverse
		'	.M11 = 0.939692736 : .M12 = 0.0 : .M13 = 0.342020124 : .M14 = 0.0
		'	.M21 = 0 : .M22 = 1 : .M23 = 0 : .M24 = 0
		'	.M31 = -0.342020124 : .M32 = 0 : .M33 = 0.939692736 : .M34 = 0
		'	.M41 = 5331064.5 : .M42 = 0 : .M43 = -303817.4 : .M44 = 1
		'End With
		'moEffect.SetValue(mhndlWorldInv, matWorldInverse)

		'moEffect.SetValue(moEffect.GetParameter(Nothing, "EntityColorAdjust"), New Vector4(1, 1, 1, 1))


		'Dim lPasses As Int32 = moEffect.Begin(FX.None)
		For X As Int32 = 0 To mlPasses - 1
			moEffect.BeginPass(X)
			oMesh.DrawSubset(0)
			moEffect.EndPass()
		Next X
		'moEffect.End()


	End Sub
	Public Sub New(ByRef oDevice As Device)
		moDevice = oDevice
		Device.IsUsingEventHandlers = False

		'moEffect = Effect.FromFile(oDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\ShaderLibrary\NormalMap.fx", Nothing, Nothing, ShaderFlags.None, Nothing)
		Dim oAssembly As System.Reflection.Assembly
		oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
		Dim sBase As String = AppDomain.CurrentDomain.FriendlyName
		Dim lIdx As Int32 = sBase.IndexOf("."c)
		sBase = sBase.Substring(0, lIdx)
		Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("MeshFormatter.NormalMap.fx")
		Device.IsUsingEventHandlers = False
		'moEffect = Effect.FromStream(moDevice, oStream, Nothing, Nothing, ShaderFlags.None, Nothing) 'Effect.FromStream(oDevice, oStream, Nothing, ShaderFlags.None, Nothing)
		moEffect = Effect.FromStream(moDevice, oStream, Nothing, ShaderFlags.None, Nothing)	'Effect.FromStream(oDevice, oStream, Nothing, ShaderFlags.None, Nothing)
		Device.IsUsingEventHandlers = True

		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\3DRT\Naval\naval-01.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\battleship1\battleship1-texture.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser3\Cruiser3-texture.jpg")
		'moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser1\Cruiser1-texture.jpg")
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
		'moBump = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\Models\MontMarx\cruiser1\BumpNormal2.bmp")
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
        moEffect.SetValue("gIllumMap", moGlowMap)
        'moEffect.SetValue("gSpecMap", moSpecMap)

		mhndlLightDir = moEffect.GetParameter(Nothing, "lightDir")

		moEffect.Technique = "NormalMapFull"

		mhndlWorldInv = moEffect.GetParameter(Nothing, "gWorldInv")
		mhndlWVP = moEffect.GetParameter(Nothing, "gWVP")
		mhndlEyePos = moEffect.GetParameter(Nothing, "gEyePosW")

		mhndlSpecPwr = moEffect.GetParameter(Nothing, "shininess")
		moEffect.SetValue(mhndlSpecPwr, 30.0F)
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
		'moEffect.SetValue("materialDiffuse", vecTemp)
		'moEffect.SetValue("materialSpecular", vecTemp)
		'moEffect.SetValue("materialAmbient", vecTemp)
		'> = {0.19f, 0.19f, 0.19f, 1.0f};
		vecTemp.X = 0.19F : vecTemp.Y = 0.19F : vecTemp.Z = 0.19F : vecTemp.W = 1.0F
		moEffect.SetValue("lightAmbient", vecTemp)

		mhndlClrRel = moEffect.GetParameter(Nothing, "RelColor")
		mhndlWorldInvTrans = moEffect.GetParameter(Nothing, "worldInverseTranspose")
		mhndlWorld = moEffect.GetParameter(Nothing, "gWorld")

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
	Private mhdlGlowValue As EffectHandle

	Public Sub SetGlowValue(ByVal lValue As Int32)
		moEffect.SetValue(mhdlGlowValue, lValue / 10.0F)
	End Sub

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

		Device.IsUsingEventHandlers = False

		Dim oAssembly As System.Reflection.Assembly
		oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
		Dim sBase As String = AppDomain.CurrentDomain.FriendlyName
		Dim lIdx As Int32 = sBase.LastIndexOf("."c)
		sBase = sBase.Substring(0, lIdx)
		Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("MeshFormatter.PostScreenGlow.fx")
		Device.IsUsingEventHandlers = False
		moEffect = Effect.FromStream(moDevice, oStream, Nothing, ShaderFlags.None, Nothing)	'Effect.FromStream(oDevice, oStream, Nothing, ShaderFlags.None, Nothing)
		Device.IsUsingEventHandlers = True

        moSceneMap = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth, oDevice.PresentationParameters.BackBufferHeight, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)
        moDownsampleMap = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth \ 4, oDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)
        moBlurMap1 = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth \ 4, oDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)
        moBlurMap2 = New Texture(oDevice, oDevice.PresentationParameters.BackBufferWidth \ 4, oDevice.PresentationParameters.BackBufferHeight \ 4, 1, Usage.RenderTarget, oDevice.PresentationParameters.BackBufferFormat, Pool.Default)

        mhdlWindowSize = moEffect.GetParameter(Nothing, "windowSize")
        mhdlSceneMap = moEffect.GetParameter(Nothing, "sceneMap")
        mhdlDownsample = moEffect.GetParameter(Nothing, "downsampleMap")
        mhdlBlurMap1 = moEffect.GetParameter(Nothing, "blurMap1")
		mhdlBlurMap2 = moEffect.GetParameter(Nothing, "blurMap2")
		mhdlGlowValue = moEffect.GetParameter(Nothing, "GlowIntensity")

		moEffect.Technique = "ScreenGlow"

		Device.IsUsingEventHandlers = True
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

Public Class CloakShader
    Private moEffect As Effect = Nothing
    Private moDevice As Device

    Private moDiffuse As Texture

    Private mhndlModel As EffectHandle
    Private mhndlIllum As EffectHandle
    Private mhndlWVP As EffectHandle
    Private mhndlScaleVal As EffectHandle

    Public sTechnique As String = "textured"

    Private mlPasses As Int32 = -1
    Public Sub PrepareToRender()
        moEffect.Technique = sTechnique
        moDevice.RenderState.CullMode = Cull.None
        moDevice.RenderState.AlphaBlendEnable = True
        moDevice.RenderState.DestinationBlend = Blend.One
        moDevice.RenderState.SourceBlend = Blend.SourceAlpha
        'wrap0=7; 
        'wrap1=7; 

        mlPasses = moEffect.Begin(FX.None)
    End Sub
    Public Sub EndRender()
        moEffect.End()
    End Sub
    Public Sub RenderMesh(ByRef oMesh As Mesh, ByVal fScaleVal As Single, ByVal oModelTex As Texture, ByVal oIllumTex As Texture)
        Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection) ' Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
        moEffect.SetValue(mhndlWVP, matWVP)
        moEffect.SetValue(mhndlScaleVal, fScaleVal)
        moEffect.SetValue(mhndlModel, oModelTex)
        moEffect.SetValue(mhndlIllum, oIllumTex)
        moEffect.CommitChanges()

        For X As Int32 = 0 To mlPasses - 1
            moEffect.BeginPass(X)
            oMesh.DrawSubset(0)
            moEffect.EndPass()
        Next X
    End Sub

    Public Sub New(ByRef oDevice As Device)
        moDevice = oDevice
        Device.IsUsingEventHandlers = False

        Device.IsUsingEventHandlers = False
        moEffect = Effect.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\cloak.fx", Nothing, ShaderFlags.None, Nothing) 'Effect.FromStream(oDevice, oStream, Nothing, ShaderFlags.None, Nothing)
        Device.IsUsingEventHandlers = True

        moDiffuse = TextureLoader.FromFile(moDevice, "C:\Documents and Settings\Matthew Campbell\Desktop\cloak.bmp")
        moEffect.SetValue("diffuseTexture", moDiffuse)

        mhndlWVP = moEffect.GetParameter(Nothing, "worldViewProj")
        mhndlScaleVal = moEffect.GetParameter(Nothing, "scaleVal")
        mhndlModel = moEffect.GetParameter(Nothing, "modelTexture")
        mhndlIllum = moEffect.GetParameter(Nothing, "modelIllum")

        Device.IsUsingEventHandlers = True
    End Sub

    Protected Overrides Sub Finalize()
        'If moEffect Is Nothing = False Then moEffect.Dispose()
        moEffect = Nothing
        MyBase.Finalize()
    End Sub
End Class
