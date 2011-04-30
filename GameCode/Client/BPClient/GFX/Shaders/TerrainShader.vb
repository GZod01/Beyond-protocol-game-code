Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class TerrainShader

	Private moEffect As Effect = Nothing

	'Texture Samplers
	Private mhndlDiffuseMap As EffectHandle
	Private mhndlNormalMap As EffectHandle
	Private mhndlIllumMap As EffectHandle
	Private mhndlFOWMap As EffectHandle

	'Set in the SetStaticData stuff
	Private mhndlEyePos As EffectHandle
	Private mhndlLightDir As EffectHandle
	Private mhndlLightClr As EffectHandle
	Private mhndlLightAmb As EffectHandle
	Private mhndlLightSpec As EffectHandle

	'Used for Bump Mapping Technique
	Private mhndlWorldInv As EffectHandle		'WorldInverse
	Private mhndlWVP As EffectHandle			'WorldViewProjection 

	'shininess
	Private mhndlSpecPwr As EffectHandle
	Private mhndlSpecMult As EffectHandle

	'Fog Settings
	Private mhndlFogEnable As EffectHandle
	Private mhndlFogStart As EffectHandle
	Private mhndlFogRange As EffectHandle
    Private mhndlFogColor As EffectHandle

    Private mhndlHalfSize As EffectHandle
    Private mhndlFullSize As EffectHandle

	Private mlPasses As Int32 
    Private mbFinalFogEnable As Boolean
    Private mbInRender As Boolean = False
    Public Sub PrepareToRender(ByVal vecLight As Vector3, ByVal clrDiffuse As System.Drawing.Color, ByVal clrAmbient As System.Drawing.Color, ByVal clrSpecular As System.Drawing.Color, ByVal fHalfSize As Single, ByVal fFullSize As Single)
        Try
            Dim moDevice As Device = GFXEngine.moDevice
            If muSettings.ShowFOWTerrainShading = True Then
                If muSettings.IlluminationMapTerrain = False Then
                    moEffect.Technique = "NormalMapOnly"
                Else
                    moEffect.Technique = "NormalMapFull"
                End If
            Else
                If muSettings.IlluminationMapTerrain = False Then
                    moEffect.Technique = "NormalMapOnlyNoFOW"
                Else
                    moEffect.Technique = "NormalMapFullNoFOW"
                End If
            End If

            moEffect.SetValue(mhndlHalfSize, fHalfSize)
            moEffect.SetValue(mhndlFullSize, fFullSize)

            moEffect.SetValue(mhndlSpecPwr, 30.0F)

            Dim fTimeOfDayMult As Single = 1.0F

            If glCurrentEnvirView = CurrentView.ePlanetView Then
                moEffect.SetValue(mhndlFogEnable, moDevice.RenderState.FogEnable)
                moEffect.SetValue(mhndlFogStart, moDevice.RenderState.FogStart)
                moEffect.SetValue(mhndlFogRange, moDevice.RenderState.FogEnd)
                With moDevice.RenderState.FogColor
                    moEffect.SetValue(mhndlFogColor, New Vector4(.R / 255.0F, .G / 255.0F, .B / 255.0F, 1.0F))
                End With
                mbFinalFogEnable = moDevice.RenderState.FogEnable
                moDevice.RenderState.FogEnable = False

                'Now, affect our light...
                If vecLight.Y > 0 Then
                    fTimeOfDayMult = 0
                Else
                    Dim vecTmp As Vector3 = vecLight
                    vecTmp.Normalize()
                    fTimeOfDayMult = Math.Min(1, Math.Abs(vecTmp.Y) * 3)
                End If
            Else
                moEffect.SetValue(mhndlFogEnable, False)
            End If

            'gEyePosW
            Dim vecEyePos As Vector4 = New Vector4(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ, 0.0F)
            'vecEyePos.Normalize()
            moEffect.SetValue(mhndlEyePos, vecEyePos)

            'lightDir
            vecLight.Normalize()
            moEffect.SetValue(mhndlLightDir, New Vector4(vecLight.X, vecLight.Y, vecLight.Z, 1))

            'lightColor
            'moEffect.SetValue(mhndlLightClr, New Vector4((clrDiffuse.R / 255.0F) * fTimeOfDayMult, (clrDiffuse.G / 255.0F) * fTimeOfDayMult, (clrDiffuse.B / 255.0F) * fTimeOfDayMult, 1.0F))
            Dim fLtR As Single = (((clrDiffuse.R / 255.0F) * 0.7F) + 0.3F) * fTimeOfDayMult
            Dim fLtG As Single = (((clrDiffuse.G / 255.0F) * 0.7F) + 0.3F) * fTimeOfDayMult
            Dim fLtB As Single = (((clrDiffuse.B / 255.0F) * 0.7F) + 0.3F) * fTimeOfDayMult


            moEffect.SetValue(mhndlLightClr, New Vector4(fLtR, fLtG, fLtB, fTimeOfDayMult))
            'lightAmbient
            Dim fVal As Single = muSettings.AmbientLevel * 0.01F
            'moEffect.SetValue(mhndlLightAmb, New Vector4(0.19F, 0.19F, 0.19F, 0.19F))
            moEffect.SetValue(mhndlLightAmb, New Vector4(fVal, fVal, fVal, fVal))

            'lightSpecular
            'moEffect.SetValue(mhndlLightSpec, New Vector4((clrSpecular.R / 255.0F) * fTimeOfDayMult, (clrSpecular.G / 255.0F) * fTimeOfDayMult, (clrSpecular.B / 255.0F) * fTimeOfDayMult, 1.0F))
            fLtR = (clrSpecular.R / 255.0F) * fTimeOfDayMult
            fLtG = (clrSpecular.G / 255.0F) * fTimeOfDayMult
            fLtB = (clrSpecular.B / 255.0F) * fTimeOfDayMult
            moEffect.SetValue(mhndlLightSpec, New Vector4(fLtR, fLtG, fLtB, fTimeOfDayMult))

            mlPasses = moEffect.Begin(FX.None)
            moEffect.BeginPass(0)
            mbInRender = True
        Catch
        End Try
    End Sub

    Public Sub EndRender()
        If mbInRender = True Then
            moEffect.EndPass()
            moEffect.End()
        End If

        GFXEngine.moDevice.RenderState.FogEnable = mbFinalFogEnable
    End Sub

    Public Sub SetTextures(ByRef oDiffuse As VolumeTexture, ByRef oNormal As VolumeTexture, ByRef oIllum As Texture, ByRef oFOW As Texture)
        moEffect.SetValue(mhndlDiffuseMap, oDiffuse)
        moEffect.SetValue(mhndlNormalMap, oNormal)
        moEffect.SetValue(mhndlIllumMap, oIllum)
        If oFOW Is Nothing = False Then moEffect.SetValue(mhndlFOWMap, oFOW)
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
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("BPClient.BumpTerrain.fx")
        Device.IsUsingEventHandlers = False
        moEffect = Effect.FromStream(moDevice, oStream, Nothing, ShaderFlags.None, Nothing) 'Effect.FromStream(oDevice, oStream, Nothing, ShaderFlags.None, Nothing)

        mhndlDiffuseMap = moEffect.GetParameter(Nothing, "gTex")
        mhndlNormalMap = moEffect.GetParameter(Nothing, "gNormalMap")
        mhndlIllumMap = moEffect.GetParameter(Nothing, "gIllumMap")

        mhndlEyePos = moEffect.GetParameter(Nothing, "gEyePosW")
        mhndlLightDir = moEffect.GetParameter(Nothing, "lightDir")
        mhndlLightClr = moEffect.GetParameter(Nothing, "lightColor")
        mhndlLightAmb = moEffect.GetParameter(Nothing, "lightAmbient")
        mhndlLightSpec = moEffect.GetParameter(Nothing, "lightSpecular")

        mhndlWorldInv = moEffect.GetParameter(Nothing, "gWorldInv")
        mhndlWVP = moEffect.GetParameter(Nothing, "gWVP")

        mhndlSpecPwr = moEffect.GetParameter(Nothing, "shininess")
        moEffect.SetValue(mhndlSpecPwr, 30.0F)

        mhndlSpecMult = moEffect.GetParameter(Nothing, "gfSpecMult")

        mhndlFogEnable = moEffect.GetParameter(Nothing, "gbRenderFog")
        mhndlFogStart = moEffect.GetParameter(Nothing, "gFogStart")
        mhndlFogRange = moEffect.GetParameter(Nothing, "gFogRange")
        mhndlFogColor = moEffect.GetParameter(Nothing, "gFogColor")

        mhndlFOWMap = moEffect.GetParameter(Nothing, "gFOWMap")

        mhndlFullSize = moEffect.GetParameter(Nothing, "gfFullSize")
        mhndlHalfSize = moEffect.GetParameter(Nothing, "gfHalfSize")

        Device.IsUsingEventHandlers = True
    End Sub

    Public Sub ReleaseMe()
        If moEffect Is Nothing = False Then moEffect.Dispose()
        moEffect = Nothing
    End Sub

End Class
