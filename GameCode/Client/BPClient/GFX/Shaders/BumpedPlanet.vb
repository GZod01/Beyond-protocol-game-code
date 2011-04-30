Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class BumpedPlanet

	Private moEffect As Effect = Nothing

	'Texture Samplers
	Private mhndlDiffuseMap As EffectHandle
	Private mhndlNormalMap As EffectHandle
	Private mhndlIllumMap As EffectHandle 

	'Set in the SetStaticData stuff
	Private mhndlEyePos As EffectHandle
	Private mhndlLightDir As EffectHandle
	Private mhndlLightClr As EffectHandle
	Private mhndlLightAmb As EffectHandle 

	'Used for Bump Mapping Technique
	Private mhndlWorldInv As EffectHandle		'WorldInverse
	Private mhndlWVP As EffectHandle			'WorldViewProjection 

	'shininess
	Private mhndlSpecPwr As EffectHandle 

	Private mlPasses As Int32
	Private mbFinalFogEnable As Boolean
	Public Sub PrepareToRender(ByVal vecLight As Vector3, ByVal clrDiffuse As System.Drawing.Color, ByVal clrAmbient As System.Drawing.Color, ByVal clrSpecular As System.Drawing.Color)
		moEffect.Technique = "NormalMapTech"

		moEffect.SetValue(mhndlSpecPwr, 3.0F)

        Dim moDevice As Device = GFXEngine.moDevice

		mbFinalFogEnable = moDevice.RenderState.FogEnable
        moDevice.RenderState.FogEnable = False 

		'gEyePosW
		Dim vecEyePos As Vector4 = New Vector4(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ, 0.0F)
		'vecEyePos.Normalize()
		moEffect.SetValue(mhndlEyePos, vecEyePos)

		'lightDir
		'vecLight = New Vector3(0, -1, 0)
		vecLight.Normalize()
        moEffect.SetValue(mhndlLightDir, New Vector4(vecLight.X, vecLight.Y, vecLight.Z, 1))

		'lightColor
		moEffect.SetValue(mhndlLightClr, New Vector4((clrDiffuse.R / 255.0F), (clrDiffuse.G / 255.0F), (clrDiffuse.B / 255.0F), 1.0F))
		'moEffect.SetValue(mhndlLightClr, New Vector4(1, 1, 1, 1))
		'lightAmbient
        'moEffect.SetValue(mhndlLightAmb, New Vector4(clrAmbient.R / 255.0F, clrAmbient.G / 255.0F, clrAmbient.B / 255.0F, 1.0F))
        Dim fVal As Single = muSettings.AmbientLevel * 0.01F
        moEffect.SetValue(mhndlLightAmb, New Vector4(fVal, fVal, fVal, 1.0F))

		mlPasses = moEffect.Begin(FX.None)
		moEffect.BeginPass(0)
	End Sub

	Public Sub EndRender()
		moEffect.EndPass()
		moEffect.End()

        GFXEngine.moDevice.RenderState.FogEnable = mbFinalFogEnable
	End Sub

	Public Sub SetTextures(ByRef oDiffuse As Texture, ByRef oNormal As Texture, ByRef oIllum As Texture)
		moEffect.SetValue(mhndlDiffuseMap, oDiffuse)
		moEffect.SetValue(mhndlNormalMap, oNormal)
		moEffect.SetValue(mhndlIllumMap, oIllum)
	End Sub

    Public Sub UpdateMatrices()
        Dim moDevice As Device = GFXEngine.moDevice
        Dim matWorld As Matrix = moDevice.Transform.World

        'Now, let's set our variables
        Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
        moEffect.SetValue(mhndlWVP, matWVP)
        Dim matWorldInv As Matrix = matWorld
        matWorldInv.Invert()
        moEffect.SetValue(mhndlWorldInv, matWorldInv)
        moEffect.CommitChanges()

    End Sub
    
    Public Sub New()
        Dim moDevice As Device = GFXEngine.moDevice
        Device.IsUsingEventHandlers = False

        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim sBase As String = AppDomain.CurrentDomain.FriendlyName
        Dim lIdx As Int32 = sBase.IndexOf("."c)
        sBase = sBase.Substring(0, lIdx)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.PlanetBump.fx")
        Device.IsUsingEventHandlers = False

        moEffect = Effect.FromStream(moDevice, oStream, Nothing, ShaderFlags.None, Nothing)

        mhndlDiffuseMap = moEffect.GetParameter(Nothing, "gTex")
        mhndlNormalMap = moEffect.GetParameter(Nothing, "gNormalMap")
        mhndlIllumMap = moEffect.GetParameter(Nothing, "gIllumMap")

        mhndlEyePos = moEffect.GetParameter(Nothing, "gEyePosW")
        mhndlLightDir = moEffect.GetParameter(Nothing, "lightDir")
        mhndlLightClr = moEffect.GetParameter(Nothing, "lightColor")
        mhndlLightAmb = moEffect.GetParameter(Nothing, "lightAmbient")

        mhndlWorldInv = moEffect.GetParameter(Nothing, "gWorldInv")
        mhndlWVP = moEffect.GetParameter(Nothing, "gWVP")

        mhndlSpecPwr = moEffect.GetParameter(Nothing, "shininess")
        moEffect.SetValue(mhndlSpecPwr, 3.0F)

        Device.IsUsingEventHandlers = True
    End Sub

    Public Sub DisposeMe()
        If moEffect Is Nothing = False Then moEffect.Dispose()
        moEffect = Nothing
    End Sub
End Class
