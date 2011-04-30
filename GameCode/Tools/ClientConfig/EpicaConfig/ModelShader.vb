Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class ModelShader

    Private Enum RenderFlags As Int32
        RenderNormalMap = 1
        RenderIllumMap = 2
        RenderPlayerCustom = 4
    End Enum

    Private moEffect As Effect = Nothing
    Private moDevice As Device

    'Texture Samplers
    Private mhndlDiffuseMap As EffectHandle
    Private mhndlNormalMap As EffectHandle
    Private mhndlIllumMap As EffectHandle

    'Set in the SetStaticData stuff
    Private mhndlEyePos As EffectHandle
    Private mhndlLightDir As EffectHandle
    Private mhndlLightClr As EffectHandle
    Private mhndlLightAmb As EffectHandle
    Private mhndlLightSpec As EffectHandle

    'Used for Bump Mapping Technique
    Private mhndlWorldInv As EffectHandle       'WorldInverse
    Private mhndlWVP As EffectHandle            'WorldViewProjection

    'Used for Per Pixel Only Technique
    Private mhndlWorld As EffectHandle          'World
    Private mhndlWorldInvTrans As EffectHandle  'WorldInverseTranspose

    'EntityColorAdjust
    Private mhndlEntityColor As EffectHandle
    'shininess
    Private mhndlSpecPwr As EffectHandle

    'RelColor
    Private mhndlRelColor As EffectHandle

    Private mhndlFogEnable As EffectHandle
    Private mhndlFogStart As EffectHandle
    Private mhndlFogRange As EffectHandle
    Private mhndlFogColor As EffectHandle

    Private mlPasses As Int32
    Private mlRenderFlags As RenderFlags
    Private mbFinalFogEnable As Boolean

    Public Sub PrepareToRender(ByVal vecLight As Vector3, ByVal clrDiffuse As System.Drawing.Color, ByVal clrAmbient As System.Drawing.Color, ByVal clrSpecular As System.Drawing.Color)

        mlRenderFlags = 0

        If muSettings.LightQuality = EngineSettings.LightQualitySetting.BumpMap Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderNormalMap
        If muSettings.RenderPlayerCustomization = True Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderPlayerCustom
        If muSettings.IlluminationMap <> -1 Then mlRenderFlags = mlRenderFlags Or RenderFlags.RenderIllumMap

        'Now, determine what technique to use...
        Try
            If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
                'NormalMap
                Select Case (mlRenderFlags And (RenderFlags.RenderIllumMap Or RenderFlags.RenderPlayerCustom))
                    Case RenderFlags.RenderIllumMap
                        moEffect.Technique = "NormalMapWithIllum"
                    Case RenderFlags.RenderPlayerCustom
                        moEffect.Technique = "NormalMapWithEntityColor"
                    Case (RenderFlags.RenderIllumMap Or RenderFlags.RenderPlayerCustom)
                        moEffect.Technique = "NormalMapFull"
                    Case Else
                        moEffect.Technique = "NormalMapOnly"
                End Select
            Else
                'PerPixel
                Select Case (mlRenderFlags And (RenderFlags.RenderIllumMap Or RenderFlags.RenderPlayerCustom))
                    Case RenderFlags.RenderIllumMap
                        moEffect.Technique = "PerPixelWithIllum"
                    Case RenderFlags.RenderPlayerCustom
                        moEffect.Technique = "PerPixelWithEntityColor"
                    Case (RenderFlags.RenderIllumMap Or RenderFlags.RenderPlayerCustom)
                        moEffect.Technique = "PerPixelFull"
                    Case Else
                        moEffect.Technique = "PerPixelOnly"
                End Select
            End If
        Catch
            'If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to set shader technique. Possible incompatibility. Auto-adjusted settings.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            muSettings.LightQuality = EngineSettings.LightQualitySetting.VSPS1
            muSettings.RenderPlayerCustomization = False
            muSettings.IlluminationMap = -1
            Return
        End Try

        moEffect.SetValue(mhndlSpecPwr, 30.0F)

        Dim fTimeOfDayMult As Single = 1.0F

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
        End If '

        'gEyePosW
        Dim vecEyePos As Vector4 = New Vector4(goCamera.mlCameraX, goCamera.mlCameraY, goCamera.mlCameraZ, 0.0F)
        'vecEyePos.Normalize()
        moEffect.SetValue(mhndlEyePos, vecEyePos)

        'lightDir
        'vecLight.Y = 0
        vecLight.Normalize()
        'Dim vecTemp As Vector3 = New Vector3(0, -1, 0)
        'vecTemp.Normalize()
        moEffect.SetValue(mhndlLightDir, New Vector4(vecLight.X, vecLight.Y, vecLight.Z, 1))
        'moEffect.SetValue(mhndlLightDir, New Vector4(vecTemp.X, vecTemp.Y, vecTemp.Z, 1))

        'lightColor
        moEffect.SetValue(mhndlLightClr, New Vector4((clrDiffuse.R / 255.0F) * fTimeOfDayMult, (clrDiffuse.G / 255.0F) * fTimeOfDayMult, (clrDiffuse.B / 255.0F) * fTimeOfDayMult, 1.0F))
        'moEffect.SetValue(mhndlLightClr, New Vector4(fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult))
        'lightAmbient
        moEffect.SetValue(mhndlLightAmb, New Vector4(Math.Max(0.19F, clrAmbient.R / 255.0F), Math.Max(0.19F, clrAmbient.G / 255.0F), Math.Max(0.19F, clrAmbient.B / 255.0F), 1.0F))
        'moEffect.SetValue(mhndlLightAmb, New Vector4(0, 0, 0, 0))
        'moEffect.SetValue(mhndlLightAmb, New Vector4(0.19F, 0.19F, 0.19F, 0.19F))
        'lightSpecular
        moEffect.SetValue(mhndlLightSpec, New Vector4((clrSpecular.R / 255.0F) * fTimeOfDayMult, (clrSpecular.G / 255.0F) * fTimeOfDayMult, (clrSpecular.B / 255.0F) * fTimeOfDayMult, 1.0F))
        'moEffect.SetValue(mhndlLightSpec, New Vector4(fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult, fTimeOfDayMult))

        mlPasses = moEffect.Begin(FX.None)
        'moEffect.BeginPass(0)
    End Sub

    Public Sub EndRender()
        'moEffect.EndPass()
        moEffect.End()

        moDevice.RenderState.FogEnable = mbFinalFogEnable
    End Sub

    Public Sub RenderMesh(ByVal oMesh As Mesh, ByVal oTexture As Texture, ByVal oNormalMap As Texture, ByVal oIllumMap As Texture)

        moEffect.SetValue(mhndlRelColor, muSettings.MyAssetColor)

        'Now, let's set our variables
        Dim matWVP As Matrix = Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection) ' Matrix.Multiply(Matrix.Multiply(moDevice.Transform.World, moDevice.Transform.View), moDevice.Transform.Projection)
        moEffect.SetValue(mhndlWVP, matWVP)
        Dim matWorldInv As Matrix = moDevice.Transform.World
        matWorldInv.Invert()
        moEffect.SetValue(mhndlWorldInv, matWorldInv)

        If (mlRenderFlags And RenderFlags.RenderNormalMap) <> 0 Then
            'Ok, use the Normal Map shader, that uses... mhndlWorldInv and mhndlWVP and the normal map
            moEffect.SetValue(mhndlNormalMap, oNormalMap)
        Else
            'ok, use the pixel shader, mhndlWorld and mhndlWorldInvTrans
            moEffect.SetValue(mhndlWorld, moDevice.Transform.World)
            Dim matWorldInvTrans As Matrix = moDevice.Transform.World
            matWorldInvTrans.Invert()
            matWorldInvTrans = Matrix.TransposeMatrix(matWorldInvTrans)
            moEffect.SetValue(mhndlWorldInvTrans, matWorldInvTrans)
        End If

        If (mlRenderFlags And RenderFlags.RenderIllumMap) <> 0 Then
            moEffect.SetValue(mhndlIllumMap, oIllumMap)
        End If
        If (mlRenderFlags And RenderFlags.RenderPlayerCustom) <> 0 Then
            moEffect.SetValue(mhndlEntityColor, New Vector4(1, 1, 1, 1))
        End If
        moEffect.CommitChanges()

        'For lPass As Int32 = 0 To mlPasses - 1
        'mlPasses = moEffect.Begin(FX.None)
        moEffect.BeginPass(0)

        Try
            moEffect.SetValue(mhndlDiffuseMap, oTexture)
            moEffect.CommitChanges()
            oMesh.DrawSubset(0)
        Catch
        End Try

        moEffect.EndPass()
        'moEffect.End()
        'Next lPass  
    End Sub

    Public Sub New(ByRef oDevice As Device)
        moDevice = oDevice
        Device.IsUsingEventHandlers = False

        Dim oAssembly As System.Reflection.Assembly
        oAssembly = System.Reflection.Assembly.LoadFrom(Application.ExecutablePath)
        Dim sBase As String = AppDomain.CurrentDomain.FriendlyName
        Dim lIdx As Int32 = sBase.IndexOf("."c)
        sBase = sBase.Substring(0, lIdx)
        Dim oStream As System.IO.Stream = oAssembly.GetManifestResourceStream("ClientConfig.NormalMap.fx")
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

        mhndlWorld = moEffect.GetParameter(Nothing, "gWorld")
        mhndlWorldInvTrans = moEffect.GetParameter(Nothing, "worldInverseTranspose")

        mhndlEntityColor = moEffect.GetParameter(Nothing, "EntityColorAdjust")

        mhndlSpecPwr = moEffect.GetParameter(Nothing, "shininess")
        moEffect.SetValue(mhndlSpecPwr, 30.0F)

        mhndlRelColor = moEffect.GetParameter(Nothing, "RelColor")


        mhndlFogEnable = moEffect.GetParameter(Nothing, "gbRenderFog")
        mhndlFogStart = moEffect.GetParameter(Nothing, "gFogStart")
        mhndlFogRange = moEffect.GetParameter(Nothing, "gFogRange")
        mhndlFogColor = moEffect.GetParameter(Nothing, "gFogColor")

        'mhndlMatDiffuse = moEffect.GetParameter(Nothing, "materialDiffuse")
        'mhndlMatSpecular = moEffect.GetParameter(Nothing, "materialSpecular")
        'mhndlMatAmbient = moEffect.GetParameter(Nothing, "materialAmbient")

        Device.IsUsingEventHandlers = True
    End Sub

    Protected Overrides Sub Finalize()
        mhndlDiffuseMap = Nothing
        mhndlNormalMap = Nothing
        mhndlIllumMap = Nothing
        mhndlEyePos = Nothing
        mhndlLightDir = Nothing
        mhndlLightClr = Nothing
        mhndlLightAmb = Nothing
        mhndlLightSpec = Nothing
        mhndlWorldInv = Nothing       'WorldInverse
        mhndlWVP = Nothing            'WorldViewProjection
        mhndlWorld = Nothing          'World
        mhndlWorldInvTrans = Nothing  'WorldInverseTranspose
        mhndlEntityColor = Nothing
        mhndlSpecPwr = Nothing
        mhndlRelColor = Nothing
        mhndlFogEnable = Nothing
        mhndlFogStart = Nothing
        mhndlFogRange = Nothing
        mhndlFogColor = Nothing

        If moEffect Is Nothing = False Then
            Try
                moEffect.Dispose()
            Catch
            End Try
        End If
        moEffect = Nothing

        moDevice = Nothing

        MyBase.Finalize()
    End Sub

    Public Sub DisposeMe()
        If moEffect Is Nothing = False Then moEffect.Dispose()
        moEffect = Nothing
    End Sub
End Class
